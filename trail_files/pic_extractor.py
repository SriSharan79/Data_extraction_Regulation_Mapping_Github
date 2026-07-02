import os
import re

import openpyxl  # Ensure you run 'pip3 install openpyxl'
# Global Namespace Dictionary mapping EASA custom layouts and MS Word OpenXML schemas
NAMESPACES = {
    'pkg': 'http://schemas.microsoft.com/office/2006/xmlPackage',
    'er': 'http://www.easa.europa.eu/erules-export',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml'
}

def parse_document_body_stream(root, flat_nodes, rels_map, media_parts, images_dir, tables_dir):
    """Parses paragraphs, tables, and binary shapes out of the layout content stream."""
    doc_part = root.find(".//pkg:part[@pkg:name='/word/document.xml']", NAMESPACES)
    if doc_part is None:
        print("CRITICAL: Master OpenXML document text engine target component missing.")
        return

    # Generates a parent tracking map for secure upward structural context lookups
    print("Indexing layout ancestry trees for deep structural asset tracking...")
    parent_map = {c: p for p in doc_part.findall('.//*') for c in p}

    def find_ancestor_sdt_id(element):
        """Walks up the parent tree to resolve the nearest active section container."""
        current = element
        while current in parent_map:
            current = parent_map[current]
            tag_local = current.tag.split('}')[-1] if '}' in current.tag else current.tag
            if tag_local == 'sdt':
                sdt_id_node = current.find('.//w:sdtPr/w:id', NAMESPACES)
                if sdt_id_node is not None:
                    found_id = sdt_id_node.attrib.get(f"{{{NAMESPACES['w']}}}val")
                    if found_id in flat_nodes:
                        return found_id
        return None
    
    def find_subsequent_figure_title(element):
        """Looks ahead in the linear XML block context to harvest a real figure caption."""
        current = element
        # Navigate up to the paragraph or container block to scan adjacent twins
        while current in parent_map and not current.tag.endswith('p'):
            current = parent_map[current]
            
        if current in parent_map:
            parent_block = parent_map[current]
            siblings = list(parent_block)
            try:
                start_idx = siblings.index(current)
                # Check the current and next 3 layout block elements for "Figure X" captions
                for i in range(start_idx, min(start_idx + 4, len(siblings))):
                    text_pieces = [t.text for t in siblings[i].findall('.//w:t', NAMESPACES) if t.text]
                    full_text = " ".join(text_pieces).strip()
                    if full_text.lower().startswith("figure"):
                        # Sanitize filename characters but preserve whitespace/dashes cleanly
                        keep_chars = [c for c in full_text if c.isalnum() or c in (' ', '-', '_')]
                        sanitized_title = "".join(keep_chars).strip().replace(" ", "_")
                        sanitized_title = sanitized_title[:10]
                        if sanitized_title:
                            return sanitized_title
            except ValueError:
                pass
        return None
    
    def _process_and_save_image(element, embed_id, rels_map, media_parts, current_store, images_dir, topic_base_name, fallback_type):
        """
        Processes OpenXML package drawing components, extracts base64 content streams, 
        and saves images to disk using adaptive structural naming policies.
        """
        # Create a dynamic local prefix map so it never breaks on missing global namespaces
        LOCAL_NAMESPACES = {}
        if 'NAMESPACES' in globals():
            LOCAL_NAMESPACES.update(globals()['NAMESPACES'])
            
        # Explicitly guarantee core drawing/layout schemas are registered
        drawing_defaults = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a14': 'http://schemas.microsoft.com/office/drawing/2010/main'
        }
        for prefix, uri in drawing_defaults.items():
            if prefix not in LOCAL_NAMESPACES:
                LOCAL_NAMESPACES[prefix] = uri

        # Extract target image path string out of document relationships (e.g., 'media/image5.jpeg')
        rel_target = rels_map.get(embed_id, "")
        
        # 1. Standardize Lookup Key Variations
        media_keys = []
        if rel_target:
            clean_target = rel_target.lstrip('/')
            media_keys = [
                f"/word/{clean_target}" if not clean_target.startswith("word/") else f"/{clean_target}",
                f"/{clean_target}",
                clean_target,
                clean_target.split('/')[-1]
            ]

        # 2. Extract Fallback Information out of Descriptive XML Attributes if relationship strings are broken
        xml_alt_name = None
        doc_pr = element.find('.//wp:docPr', LOCAL_NAMESPACES)
        nv_pic_pr = element.find('.//pic:cNvPr', LOCAL_NAMESPACES)
        
        if nv_pic_pr is not None and nv_pic_pr.attrib.get('name'):
            xml_alt_name = nv_pic_pr.attrib.get('name')
        elif doc_pr is not None and doc_pr.attrib.get('name'):
            xml_alt_name = doc_pr.attrib.get('name')

        # 3. Primary Payload Extraction
        image_bytes = None
        matched_key = None
        
        for key in media_keys:
            if key in media_parts and media_parts[key]:
                image_bytes = media_parts[key]
                matched_key = key
                break

        # 4. Flat Package Pattern-Matching Rescue Sequence (Runs if direct key matches fail)
        if not image_bytes and xml_alt_name:
            numeric_ids = re.findall(r'\d+', xml_alt_name)
            if numeric_ids:
                target_id = numeric_ids[0]
                # Match specific flat sequences like image5.jpeg or image5.png exactly
                target_pattern = re.compile(r'image' + target_id + r'\.(png|jpe?g|gif|emf|wmf)$', re.IGNORECASE)
                
                for real_key in media_parts.keys():
                    if target_pattern.search(real_key) or (f"media/image{target_id}." in real_key.lower()):
                        image_bytes = media_parts[real_key]
                        matched_key = real_key
                        break

        # Exit gracefully if the element references no physical payload assets
        if not image_bytes:
            return

        # Increment runtime workspace counters securely
        current_store["image_counter"] = current_store.get("image_counter", 0) + 1
        
        # Discover original file format extension safely
        _, native_extension = os.path.splitext(matched_key.lower())
        if not native_extension or len(native_extension) > 5:
            native_extension = ".jpg" if fallback_type != "imagedata" else ".png"

        # Determine structural image titles from adjacent sibling blocks ("Figure X...")
        discovered_figure_title = find_subsequent_figure_title(element)
        
        if discovered_figure_title:
            img_filename = f"{discovered_figure_title}{native_extension}"
        elif xml_alt_name and "picture" not in xml_alt_name.lower():
            clean_alt = "".join([c if c.isalnum() else "_" for c in os.path.splitext(xml_alt_name)[0]]).strip("_")
            img_filename = f"{clean_alt}{native_extension}" if clean_alt else f"img_{current_store['id']}_{embed_id}{native_extension}"
        else:
            img_filename = f"{topic_base_name}_{fallback_type}_{current_store['image_counter']}{native_extension}"

        img_path = os.path.join(images_dir, img_filename)
        
        # Write payload file to disk workspace directory
        try:
            with open(img_path, 'wb') as img_f:
                img_f.write(image_bytes)
            
            # Save relative file path back to JSON hierarchy nodes
            if img_path not in current_store["extracted_images"]:
                current_store["extracted_images"].append(img_path)
        except Exception as e:
            print(f"[ERROR] Failed writing payload byte stream to disc path: {img_path}. Details: {e}") 
    
 
    active_node_id = None
    table_counter_global = 0
    
    # Process elements sequentially in their native layout order
    for element in doc_part.findall('.//*'):
        tag_local = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        
        # Track active container section context changes linearly
        if tag_local == 'sdt':
            sdt_id_node = element.find('.//w:sdtPr/w:id', NAMESPACES)
            if sdt_id_node is not None:
                found_id = sdt_id_node.attrib.get(f"{{{NAMESPACES['w']}}}val")
                if found_id in flat_nodes:
                    active_node_id = found_id

        # Use parent-map ancestry if nesting structures bypass linear contexts
        current_context_id = active_node_id if active_node_id else find_ancestor_sdt_id(element)

        if not current_context_id:
            continue
            
        current_store = flat_nodes[current_context_id]
        topic_base_name_raw = current_store.get("attributes", {}).get("title", "document")
        if isinstance(topic_base_name_raw, list):
            topic_base_name_raw = topic_base_name_raw[0]
        topic_base_name = "".join([c if c.isalnum() else "_" for c in topic_base_name_raw]).strip("_")[:10]

        # Parse inline text characters
        if tag_local == 't' and element.text and element.text.strip():
            text_val = element.text.strip()
            if text_val not in current_store["text_lines"]:
                current_store["text_lines"].append(text_val)

        # Parse hyperlinks
        elif tag_local == 'hyperlink':
            r_id = element.attrib.get(f"{{{NAMESPACES['r']}}}id")
            link_text = "".join([t.text for t in element.findall('.//w:t', NAMESPACES) if t.text]).strip()
            if link_text:
                target_url = rels_map.get(r_id, "Internal Jump Link")
                current_store["hyperlinks"].append({"text": link_text, "target": target_url})

        # Parse structural standard drawing elements (Standard OpenXML Pictures)
        elif tag_local == 'drawing':
            pic_nodes = element.findall('.//pic:pic', NAMESPACES)
            
            if pic_nodes:
                for pic in pic_nodes:
                    blip = pic.find('.//pic:blipFill/a:blip', NAMESPACES)
                    if blip is None:
                        continue
                    embed_id = blip.attrib.get(f"{{{NAMESPACES['r']}}}embed")
                    if not embed_id:
                        continue
                    
                    print(f"[DEBUG Image] type 1-fig)")
                    _process_and_save_image(
                        element=element, 
                        embed_id=embed_id, 
                        rels_map=rels_map, 
                        media_parts=media_parts, 
                        current_store=current_store, 
                        images_dir=images_dir, 
                        topic_base_name=topic_base_name, 
                        fallback_type="fig"
                    )
            else:
                # Fallback handler for grouped or alternative inline drawing shapes
                embed_ids = []
                for blip in element.findall('.//a:blip', NAMESPACES):
                    eid = blip.attrib.get(f"{{{NAMESPACES['r']}}}embed")
                    if eid: 
                        embed_ids.append(eid)
                    
                for embed_id in embed_ids:
                    print(f"[DEBUG Image] type 2-drawing)")
                    _process_and_save_image(
                        element=element, 
                        embed_id=embed_id, 
                        rels_map=rels_map, 
                        media_parts=media_parts, 
                        current_store=current_store, 
                        images_dir=images_dir, 
                        topic_base_name=topic_base_name, 
                        fallback_type="drawing"
                    )

        # Parse legacy VML canvas shapes
        elif tag_local == 'imagedata':
            r_id = element.attrib.get(f"{{{NAMESPACES['r']}}}id")
            if r_id:
                print(f"[DEBUG Image] type 3-fallback)")
                _process_and_save_image(
                    element=element, 
                    embed_id=r_id, 
                    rels_map=rels_map, 
                    media_parts=media_parts, 
                    current_store=current_store, 
                    images_dir=images_dir, 
                    topic_base_name=topic_base_name, 
                    fallback_type="fallback"
                )

        # Parse structured layout tabular content
        elif tag_local == 'tbl':
            if "table_counter" not in current_store:
                current_store["table_counter"] = 0
            current_store["table_counter"] += 1
            
            table_filename = f"{topic_base_name}_{current_store['table_counter']}.xlsx"
            single_table_path = os.path.join(tables_dir, table_filename)
            
            single_wb = openpyxl.Workbook()
            ws = single_wb.active
            ws.title = "Extracted Grid Data"
            
            has_data = False
            for row in element.findall('.//w:tr', NAMESPACES):
                row_cells = []
                for cell in row.findall('.//w:tc', NAMESPACES):
                    cell_str = " ".join([t.text for t in cell.findall('.//w:t', NAMESPACES) if t.text]).strip()
                    row_cells.append(cell_str)
                if any(row_cells):
                    ws.append(row_cells)
                    has_data = True
            
            if has_data:
                single_wb.save(single_table_path)
                if table_filename not in current_store["extracted_tables"]:
                    current_store["extracted_tables"].append(table_filename)
            else:
                try: 
                    single_wb.close()
                except Exception: 
                    pass