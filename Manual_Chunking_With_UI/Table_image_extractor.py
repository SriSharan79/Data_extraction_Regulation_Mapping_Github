from __future__ import annotations

import logging
import math
import os
import shutil
import sys
import tempfile
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path
from typing import Dict, List, Tuple, Optional

try:
    import imagehash  # optional: enables perceptual de-duplication of images
except ImportError:
    imagehash = None

# NOTE: pandas, tqdm, docling and docling_core are heavyweight and only needed
# when an extraction actually runs. They are imported lazily inside the methods
# below so this module (and any UI that imports it) loads even when those
# packages are not installed. `from __future__ import annotations` keeps the
# docling type hints (e.g. PictureItem) from being evaluated at import time.


logger = logging.getLogger(__name__)


_IS_TTY = sys.stdout.isatty()


def _tqdm(iterable, **kwargs):
    """Wrap an iterable in a tqdm progress bar if tqdm is available; otherwise
    return it unchanged (progress bars are cosmetic, not required)."""
    try:
        from tqdm import tqdm
        return tqdm(iterable, **kwargs)
    except ImportError:
        return iterable

def bbox_to_dict(bbox) -> Optional[Dict]:
    """
    Convert Docling BoundingBox to JSON-serializable dict.

    Args:
        bbox: Docling bbox object

    Returns:
        Dict with l, t, r, b keys or None
    """
    if bbox is None:
        return None

    try:
        return {
            "l": bbox.l,
            "t": bbox.t,
            "r": bbox.r,
            "b": bbox.b
        }
    except AttributeError:
        try:
            return {
                "l": getattr(bbox, 'x0', None),
                "t": getattr(bbox, 'y0', None),
                "r": getattr(bbox, 'x1', None),
                "b": getattr(bbox, 'y1', None)
            }
        except Exception:
            return None

class DoclingExtractor:
    """
    Extract tables, images, and headings using Docling.
    Designed to run in parallel with PyMuPDF extraction.
    """

    def __init__(
        self,
        input_path: str,
        tables_output_path: str,
        images_output_path: str,
        enable_ocr: bool = None,
        image_resolution_scale: float = None,
    ):
        self.input_path = Path(input_path)
        self.tables_output_path = Path(tables_output_path)
        self.images_output_path = Path(images_output_path)
        self.document_name = self.input_path.stem
        self.image_hashes = set()

        _ocr = enable_ocr if enable_ocr is not None else True
        _scale = image_resolution_scale if image_resolution_scale is not None else 1.0

        # Heavy docling imports happen here (construction), not at module import.
        from docling.datamodel.pipeline_options import PdfPipelineOptions
        from docling.document_converter import DocumentConverter, PdfFormatOption
        from docling.datamodel.base_models import InputFormat

        pipeline_options = PdfPipelineOptions()
        pipeline_options.do_ocr = _ocr
        pipeline_options.images_scale = _scale
        pipeline_options.generate_page_images = False
        pipeline_options.generate_picture_images = True

        self.doc_converter = DocumentConverter(
            format_options={
                InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
            }
        )

    def extract_all(self) -> Dict:
        """
        Extract tables, images, and headings in one pass.

        Returns:
            Dictionary with 'tables', 'images', and 'headings' keys.
            Image entries include a 'hash' field for cross-chunk deduplication.
        """
        logger.info(f"[Docling] Starting extraction from {self.input_path}")
        print(f"\n[Docling] Converting PDF chunk: {self.input_path.name} ...", flush=True)

        doc = self.doc_converter.convert(self.input_path)

        tables = self._extract_tables(doc)
        images = self._extract_images(doc)
        headings = self._extract_headings(doc)

        logger.info(
            f"[Docling] Extraction complete: {len(tables)} tables, "
            f"{len(images)} images, {len(headings)} headings"
        )

        return {"tables": tables, "images": images, "headings": headings}

    def _extract_tables(self, doc) -> List[Dict]:
        """Extract all tables and save as CSV files."""
        import pandas as pd

        tables_data = []
        self.tables_output_path.mkdir(parents=True, exist_ok=True)

        tables = doc.document.tables
        for table_ix, table in enumerate(_tqdm(tables, desc="[Docling] Extracting tables",
                                               unit="table", leave=False, disable=not _IS_TTY)):
            try:
                df: pd.DataFrame = table.export_to_dataframe(doc)
            except Exception:
                df = pd.DataFrame()

            if df.empty and hasattr(table, 'data') and table.data and table.data.grid:
                rows = [[cell.text if cell else '' for cell in row] for row in table.data.grid]
                if rows:
                    df = pd.DataFrame(rows)

            page_no = table.prov[0].page_no if getattr(table, "prov", None) else "unknown"
            bbox = getattr(table.prov[0], "bbox", None) if getattr(table, "prov", None) else None

            if df.empty:
                logger.warning(
                    f"[Docling] Could not extract content for table {table_ix} on page {page_no}; "
                    f"saving empty entry so it is still tagged to its section in the output JSON"
                )
                df = pd.DataFrame()

            if not df.empty and df.columns.duplicated().any():
                seen: dict = {}
                new_cols = []
                for col in df.columns:
                    col_str = str(col)
                    if col_str in seen:
                        seen[col_str] += 1
                        new_cols.append(f"{col_str}_{seen[col_str]}")
                    else:
                        seen[col_str] = 0
                        new_cols.append(col_str)
                df.columns = new_cols

            table_name = f"{self.document_name}_page_{page_no}_table_{table_ix}.csv"
            table_path = self.tables_output_path / table_name
            df.to_csv(table_path, index=False)

            tables_data.append({
                "table_index": table_ix,
                "page_no": page_no,
                "csv_path": str(table_path),
                "data": df.to_dict(orient="records"),
                "bbox": bbox_to_dict(bbox),
            })

            logger.debug(f"[Docling] Extracted table {table_ix} from page {page_no}")

        return tables_data

    def _extract_images(self, doc) -> List[Dict]:
        """Extract all images with deduplication. Includes 'hash' field per entry."""
        from docling_core.types.doc import PictureItem

        images_data = []
        picture_counter = 0
        self.images_output_path.mkdir(parents=True, exist_ok=True)

        all_items = list(doc.document.iterate_items())
        for element, _level in _tqdm(all_items, desc="[Docling] Extracting images", unit="item",
                                    leave=False, disable=not _IS_TTY):
            if isinstance(element, PictureItem):
                page_no = (
                    getattr(element.prov[0], "page_no", "unknown")
                    if getattr(element, "prov", None) else "unknown"
                )
                picture_counter += 1

                # Perceptual de-duplication (only when the optional imagehash
                # dependency is installed; otherwise every image is kept).
                image_hash = self._get_image_hash(element, doc)
                if image_hash is not None:
                    if image_hash in self.image_hashes:
                        logger.debug(f"[Docling] Skipping duplicate image from page {page_no}")
                        continue
                    self.image_hashes.add(image_hash)

                image_name = f"{self.document_name}_page_{page_no}_picture_{picture_counter}.png"
                image_path = self.images_output_path / image_name
                with image_path.open("wb") as fp:
                    element.get_image(doc).save(fp, "PNG")

                bbox = (
                    getattr(element.prov[0], "bbox", None)
                    if getattr(element, "prov", None) else None
                )

                images_data.append({
                    "index": picture_counter,
                    "page_no": page_no,
                    "image_path": str(image_path),
                    "bbox": bbox_to_dict(bbox),
                    "hash": image_hash,
                })

                logger.debug(f"[Docling] Extracted image {picture_counter} from page {page_no}")

        return images_data

    def _extract_headings(self, doc) -> List[Dict]:
        """Extract section headings detected by Docling's layout analysis."""
        headings = []

        for element, level in doc.document.iterate_items():
            if hasattr(element, 'label') and element.label in ['section_header', 'title', 'subtitle']:
                page_no = (
                    getattr(element.prov[0], "page_no", 1)
                    if getattr(element, "prov", None) else 1
                )
                bbox = (
                    getattr(element.prov[0], "bbox", None)
                    if getattr(element, "prov", None) else None
                )

                headings.append({
                    "level": level + 1,
                    "title": element.text if hasattr(element, 'text') else str(element),
                    "page": page_no,
                    "bbox": bbox_to_dict(bbox),
                })

                logger.debug(f"[Docling] Detected heading (level {level}): '{element.text}' on page {page_no}")

        return headings

    def _get_image_hash(self, picture: PictureItem, doc) -> Optional[str]:
        """Generate a perceptual hash to detect duplicate images.

        Returns ``None`` when the optional ``imagehash`` dependency is not
        installed, which disables de-duplication (every image is kept).
        """
        if imagehash is None:
            return None
        try:
            pil_image = picture.get_image(doc)
            return str(imagehash.dhash(pil_image, hash_size=8))
        except Exception:
            return None

