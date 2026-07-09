
import logging
import math
import os
import shutil
import sys
import tempfile
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path
from typing import Dict, List, Tuple,Optional

import fitz
# import imagehash
import pandas as pd
from tqdm import tqdm

# Docling is heavy to import; its symbols are imported lazily inside the
# methods that use them (see DoclingExtractor.__init__ and _extract_images).

logger = logging.getLogger(__name__)


_IS_TTY = sys.stdout.isatty()


def build_doc_converter(enable_ocr: bool = True, image_resolution_scale: float = 1.0):
    """
    Build a Docling ``DocumentConverter`` with the project's pipeline options.

    This is the expensive step (it wires up the layout/OCR model pipelines), so
    for batch processing the resulting converter should be built once and reused
    across many PDFs via :func:`get_shared_doc_converter`.
    """
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.base_models import InputFormat

    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_ocr = enable_ocr
    pipeline_options.images_scale = image_resolution_scale
    pipeline_options.generate_page_images = False
    pipeline_options.generate_picture_images = True

    # Pick the compute device by availability: CUDA when present, else Apple
    # MPS, else CPU. AUTO is Docling's default, but set it explicitly and log
    # the decision so runs on different machines show what they used.
    try:
        from docling.datamodel.pipeline_options import AcceleratorOptions, AcceleratorDevice
        from docling.utils.accelerator_utils import decide_device
        pipeline_options.accelerator_options = AcceleratorOptions(device=AcceleratorDevice.AUTO)
        device = decide_device(AcceleratorDevice.AUTO.value)
        logger.info(f"Docling accelerator device: {device}")
        print(f"ℹ Info: Docling running on device: {device}")
    except Exception as e:  # older docling without accelerator options
        logger.warning(f"Could not configure accelerator device (using docling defaults): {e}")

    # Use RapidOCR (ONNX Runtime backend) as the OCR engine. Unlike the EasyOCR
    # default it is lightweight and bundles cleanly into a standalone build; its
    # models are fetched on first use. Fall back to the default engine if
    # RapidOCR is unavailable.
    if enable_ocr:
        try:
            from docling.datamodel.pipeline_options import RapidOcrOptions
            pipeline_options.ocr_options = RapidOcrOptions(lang=["english"])
        except Exception as e:
            logger.warning(f"RapidOCR unavailable, using default OCR engine: {e}")

    return DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
        }
    )


# Process-wide cache of a shared converter, keyed by its pipeline options so a
# change of options rebuilds it. Reused across a whole batch run.
_shared_converter = None
_shared_converter_key = None


def get_shared_doc_converter(enable_ocr: bool = True, image_resolution_scale: float = 1.0):
    """
    Return a cached ``DocumentConverter``, building it on first use.

    Use this for batch analysis so the Docling model pipeline is loaded only once
    and reused for every PDF, instead of being re-initialised per file.
    """
    global _shared_converter, _shared_converter_key
    key = (bool(enable_ocr), float(image_resolution_scale))
    if _shared_converter is None or _shared_converter_key != key:
        logger.info("Building shared Docling converter (loaded once, reused for the batch).")
        _shared_converter = build_doc_converter(enable_ocr, image_resolution_scale)
        _shared_converter_key = key
    return _shared_converter

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
        doc_converter=None,
    ):
        self.input_path = Path(input_path)
        self.tables_output_path = Path(tables_output_path)
        self.images_output_path = Path(images_output_path)
        self.document_name = self.input_path.stem
        self.image_hashes = set()

        _ocr = enable_ocr if enable_ocr is not None else True
        _scale = image_resolution_scale if image_resolution_scale is not None else 1.0

        # Reuse a pre-built (shared) converter when provided; otherwise build a
        # dedicated one for this extractor (single-file / standalone use).
        self.doc_converter = doc_converter if doc_converter is not None else build_doc_converter(_ocr, _scale)

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
        tables_data = []
        self.tables_output_path.mkdir(parents=True, exist_ok=True)

        tables = doc.document.tables
        for table_ix, table in enumerate(tqdm(tables, desc="[Docling] Extracting tables",
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

            # 1. Update the extension to .xlsx
            table_name = f"_page_{page_no}_table_{table_ix}.xlsx"

            # 2. Point to the 'excel_files' subfolder inside your output path
            excel_output_dir = self.tables_output_path / "excel_files"

            # 3. Create the subfolder if it doesn't exist yet
            excel_output_dir.mkdir(parents=True, exist_ok=True)

            # 4. Define the final file path and save
            table_path = excel_output_dir / table_name
            df.to_excel(table_path, index=False)

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
        for element, _level in tqdm(all_items, desc="[Docling] Extracting images", unit="item",
                                    leave=False, disable=not _IS_TTY):
            if isinstance(element, PictureItem):
                page_no = (
                    getattr(element.prov[0], "page_no", "unknown")
                    if getattr(element, "prov", None) else "unknown"
                )
                picture_counter += 1

                # image_hash = self._get_image_hash(element, doc)

                # if image_hash in self.image_hashes:
                #     logger.debug(f"[Docling] Skipping duplicate image from page {page_no}")
                #     continue
                # self.image_hashes.add(image_hash)

                image_name = f"_page_{page_no}_picture_{picture_counter}.png"
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
                    # "hash": image_hash,
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

    # def _get_image_hash(self, picture: PictureItem, doc) -> str:
    #     """Generate perceptual hash to detect duplicate images."""
    #     pil_image = picture.get_image(doc)
    #     return str(imagehash.dhash(pil_image, hash_size=8))

