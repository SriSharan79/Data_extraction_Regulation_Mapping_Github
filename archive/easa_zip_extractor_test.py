import hashlib
import os
import json
import base64

import re
import zipfile
import io
import xml.etree.ElementTree as ET
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

def sanitize_filename(name):
    """Converts a topic title into a safe, clean string usable as an OS filename."""
    if not name:
        return "Unknown_Topic"
    clean = re.sub(r'[^a-zA-Z0-9\s\-\.]', '_', name)
    clean = re.sub(r'[\s_]+', '_', clean)
    return clean.strip('_')[:120]

def resolve_paths(storage_path, zip_file_path):
    """Generates deterministic target folder paths for the data assets."""
    doc_name = os.path.splitext(os.path.basename(zip_file_path))[0]
    doc_name_2 = doc_name.rsplit('_', 1)[-1]
    target_root = os.path.join(storage_path, doc_name_2)
    Com_root = os.path.join(storage_path,"All_Combined")
    os.makedirs(Com_root, exist_ok=True)
    base_hash = hashlib.md5(doc_name_2.encode()).hexdigest()[:8]
    
    return {
        "images_dir": os.path.join(target_root, "images"),
        "tables_dir": os.path.join(target_root, "tables"),
        "output_json": os.path.join(target_root, f"{base_hash}_Extraction_Json.json"),
        "output_json_com": os.path.join(Com_root, f"{doc_name_2}.json")
    }

def load_xml_from_zip(zip_path):
    """Unpacks the main embedded XML package tree directly into memory."""
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            xml_target_name = [f for f in z.namelist() if f.endswith('.xml')][0]
            print(f"Streaming package target data stream: {xml_target_name}")
            return ET.parse(io.BytesIO(z.read(xml_target_name))).getroot()
    except Exception as e:
        print(f"CRITICAL: Failed to process ZIP archive file: {e}")
        return None

def extract_relationships_and_media(root):
    """Maps system document layout hyperlinks and extracts raw image buffers."""
    rels_map = {}
    media_parts = {}
    
    for part in root.findall('.//pkg:part', NAMESPACES):
        part_name = part.attrib.get(f"{{{NAMESPACES['pkg']}}}name", "")
        
        # Parse document explicit relationship records
        if "document.xml.rels" in part_name or part_name.endswith(".rels"):
            for rel in part.findall('.//pkg:xmlData//Relationship', NAMESPACES):
                rels_map[rel.attrib.get('Id')] = rel.attrib.get('Target')
                
        # Capture and decode embedded base64 media resources
        if "/word/media/" in part_name:
            string_data = part.find('.//pkg:stringData', NAMESPACES)
            if string_data is not None and string_data.text:
                try:
                    media_parts[part_name] = base64.b64decode(string_data.text.strip())
                except Exception:
                    pass

    # FIXED: Safely duplicate items into standard short filename indices without list.keys().items() crashing
    for part_name, binary_payload in list(media_parts.items()):
        filename = part_name.split('/')[-1]
        media_parts[filename] = binary_payload
        
    return rels_map, media_parts

def build_structural_index(er_doc_node):
    """Builds a flat catalog map of structural nodes based on EASA metadata."""
    flat_nodes = {}
    easa_attributes = [
    'sdt-id',
    'source-title',
    'ERulesId',
    'Domain',
    'ActivityType',
    'AircraftUse',
    'AircraftCategory',
    'AmendedBy',
    'ApplicabilityDate',
    'EntryIntoForceDate',
    'EquivalentForeignRegulation',
    'ICAOReference',
    'Keywords',
    'RegistryState',
    'RegulatedEntity',
    'RegulatorySource',
    'RegulatorySubject',
    'TechnicalSubjectMatter',
    'TypeOfContent',
    'ParentIR',
    'EASACategory',
    'title'
]

    for element in er_doc_node.findall('.//*[@sdt-id]'):
        sdt_id = element.attrib.get('sdt-id')
        tag_name = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        
        attribs = {}
        for attr in easa_attributes:
            if attr in element.attrib and element.attrib[attr].strip():
                attribs[attr] = element.attrib[attr].strip()
                
        raw_name = attribs.get('source-title') or attribs.get('title') or f"{tag_name}_{sdt_id}"
        sanitized_name = sanitize_filename(raw_name)
        
        flat_nodes[sdt_id] = {
            "id": sdt_id,
            "element_type": tag_name,
            "attributes": attribs,
            "sanitized_name": sanitized_name,
            "text_lines": [],
            "hyperlinks": [],
            "extracted_images": [],
            "extracted_tables": [],
            "children_ids": [],
            "image_counter": 0,
            "table_counter": 0
        }

    def map_dependencies(element):
        sdt_id = element.attrib.get('sdt-id')
        for child in element:
            child_sdt_id = child.attrib.get('sdt-id')
            if sdt_id and child_sdt_id and child_sdt_id in flat_nodes:
                # FIX: Mutate global dictionary map instance explicitly
                if child_sdt_id not in flat_nodes[sdt_id]["children_ids"]:
                    flat_nodes[sdt_id]["children_ids"].append(child_sdt_id)
            map_dependencies(child)

    map_dependencies(er_doc_node)
    return flat_nodes

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
        LOCAL_NAMESPACES = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a14': 'http://schemas.microsoft.com/office/drawing/2010/main'
        }

        # Extract target image path string out of document relationships (e.g., 'media/image5.jpeg')
        rel_target = rels_map.get(embed_id, "")
        
        # 1. Standardize Lookup Key Variations
        media_keys = []
        if rel_target:
            # Standardize structural target formats to align cleanly with package paths
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

        # Define unique save path using the active node block identity
        img_filename = f"img_{current_store['id']}_{embed_id}{native_extension}"
        
        # Try using descriptive metadata layout strings as file names if available
        if xml_alt_name and "picture" not in xml_alt_name.lower():
            clean_alt = "".join([c if c.isalnum() else "_" for c in os.path.splitext(xml_alt_name)[0]]).strip("_")
            if clean_alt:
                img_filename = f"{clean_alt}{native_extension}"

        img_path = os.path.join(images_dir, img_filename)
        
        # Write payload file to disk workspace directory
        try:
            with open(img_path, 'wb') as img_f:
                img_f.write(image_bytes)
            
            # Save relative file path back to JSON hierarchy nodes
            if img_path not in current_store["extracted_images"]:
                current_store["extracted_images"].append(img_path)
        except Exception as e:
            print(f"[ERROR] Failed writing payload byte steam to disc path: {img_path}. Details: {e}")
    
    def _process_and_save_image(element, media_keys, current_store, topic_base_name, fallback_type, xml_alt_name=None):
        """
        Modular helper to find binary payloads, execute naming strategies,
        and save image elements with fallback ID matching for unmatched relationship schemas.
        """
        # Safeguard namespace dictionary ensuring DrawingML prefixes resolve safely
        LOCAL_NAMESPACES = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a14': 'http://schemas.microsoft.com/office/drawing/2010/main'
        }
        if 'NAMESPACES' in globals():
            LOCAL_NAMESPACES.update(globals()['NAMESPACES'])

        # 1. Check if the incoming media keys are broken or empty
        is_key_broken = not media_keys or any(key in ('/word/', '') for key in media_keys)
        
        if is_key_broken:
            print(f"[DEBUG Image] Empty target detected. Attempting to harvest embedded filename properties...")
            
            doc_pr = element.find('.//wp:docPr', LOCAL_NAMESPACES)
            nv_pic_pr = element.find('.//pic:cNvPr', LOCAL_NAMESPACES)
            
            discovered_xml_name = None
            if nv_pic_pr is not None and nv_pic_pr.attrib.get('name'):
                discovered_xml_name = nv_pic_pr.attrib.get('name')
            elif doc_pr is not None and doc_pr.attrib.get('name'):
                discovered_xml_name = doc_pr.attrib.get('name')
                
            if discovered_xml_name:
                print(f"[DEBUG Image] Rescued raw layout title: '{discovered_xml_name}'")
                media_keys = [
                    f"/word/media/{discovered_xml_name}",
                    f"media/{discovered_xml_name}",
                    discovered_xml_name
                ]
                if not xml_alt_name:
                    xml_alt_name = discovered_xml_name

        # 2. Main Payload Search Attempt
        image_bytes = None
        matched_key = None
        
        for key in media_keys:
            if key in media_parts and media_parts[key]:
                image_bytes = media_parts[key]
                matched_key = key
                break

        # 3. Flat XML Package Lookup Strategy (Resolving layout titles to underlying media items)
        if not image_bytes and xml_alt_name:
            print(f"[DEBUG Image] Direct map failed. Attempting numeric ID isolation pattern matching for '{xml_alt_name}'...")
            
            # Pull any sequence of integers out of the string (e.g., "Picture 1" -> "1")
            numeric_ids = re.findall(r'\d+', xml_alt_name)
            if numeric_ids:
                target_id = numeric_ids[0]
                print(f"[DEBUG Image] Extracted target identity index code: '{target_id}'")
                
                # Step 3A: Search available media keys using absolute exact boundaries
                # Looks for "image1.png", "/word/media/image1.jpeg", etc.
                target_pattern = re.compile(r'image' + target_id + r'\.(png|jpe?g|gif|emf|wmf)$', re.IGNORECASE)
                
                for real_key in media_parts.keys():
                    if target_pattern.search(real_key):
                        print(f"[DEBUG Image] Flat XML Package Success! Mapping '{xml_alt_name}' to asset package route: '{real_key}'")
                        image_bytes = media_parts[real_key]
                        matched_key = real_key
                        break
                
                # Step 3B: Relaxed search fallback (If strict match fails, look for the digits anywhere in the image name)
                if not image_bytes:
                    for real_key in media_parts.keys():
                        if 'media' in real_key.lower() and re.search(r'(?<!\d)' + target_id + r'(?!\d)', real_key):
                            print(f"[DEBUG Image] Relaxed Pattern Match Success! Mapping to route: '{real_key}'")
                            image_bytes = media_parts[real_key]
                            matched_key = real_key
                            break

        # Exit cleanly if completely unresolvable
        if not image_bytes:
            print(f"[DEBUG Image] WARNING: No binary asset payload found matching keys: {media_keys} or structural fallbacks.")
            return

        current_store["image_counter"] += 1
        
        # Discover original file format extension safely
        _, native_extension = os.path.splitext(matched_key.lower())
        if not native_extension or len(native_extension) > 5:
            native_extension = ".jpg" if fallback_type != "imagedata" else ".png"
            
        # Discover adjacent structural textual captions ("Figure X...")
        discovered_figure_title = find_subsequent_figure_title(element)
        print(f"Picture identified: {discovered_figure_title}")
        
        if discovered_figure_title:
            img_filename = f"{discovered_figure_title}{native_extension}"
        elif xml_alt_name:
            clean_name = os.path.splitext(xml_alt_name)[0].replace(" ", "_")
            img_filename = f"{clean_name}{native_extension}"
            print(f"[DEBUG Image] Sibling text missing. Using embedded XML tag title: {img_filename}")
        else:
            img_filename = f"{topic_base_name}_{fallback_type}_{current_store['image_counter']}{native_extension}"
            print(f"[DEBUG Image] Sibling text & XML tags missing. Using structure fallback title: {img_filename}")
            
        img_path = os.path.join(images_dir, img_filename)
        print(f"[DEBUG Image] SUCCESS: Saving image to target route -> {img_path} ({len(image_bytes)} bytes)")
        
        try:
            with open(img_path, 'wb') as img_f:
                img_f.write(image_bytes)
            if img_filename not in current_store["extracted_images"]:
                current_store["extracted_images"].append(img_filename)
        except Exception as e:
            print(f"[DEBUG Image] CRITICAL: Failed writing payload to disc. Error: {e}")
 
    active_node_id = None
    
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
        topic_base_name_raw = current_store["sanitized_name"]
        topic_base_name = topic_base_name_raw[:10]

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
                    
                    rel_target = rels_map.get(embed_id, "")
                    media_keys = [
                        f"/word/{rel_target}" if not rel_target.startswith("/") else rel_target,
                        rel_target.split('/')[-1] if rel_target else ""
                    ]
                    
                    nv_pr = pic.find('.//pic:nvPicPr/pic:cNvPr', NAMESPACES)
                    xml_attr_name = nv_pr.attrib.get('name') if nv_pr is not None else None
                    print(f"[DEBUG Image] type 1-fig)")
                    
                    # Call modular processor
                    _process_and_save_image(element, media_keys, current_store, topic_base_name, "fig", xml_attr_name)
            else:
                # Fallback handler for grouped or alternative inline drawing shapes
                embed_ids = []
                for blip in element.findall('.//a:blip', NAMESPACES):
                    eid = blip.attrib.get(f"{{{NAMESPACES['r']}}}embed")
                    if eid: 
                        embed_ids.append(eid)
                    
                for embed_id in embed_ids:
                    rel_target = rels_map.get(embed_id, "")
                    media_keys = [
                        f"/word/{rel_target}" if not rel_target.startswith("/") else rel_target,
                        rel_target.split('/')[-1] if rel_target else ""
                    ]
                    
                    print(f"[DEBUG Image] type 2-drawing)")
                    _process_and_save_image(element, media_keys, current_store, topic_base_name, "drawing")

        # Parse legacy VML canvas shapes
        elif tag_local == 'imagedata':
            r_id = element.attrib.get(f"{{{NAMESPACES['r']}}}id")
            if r_id:
                rel_target = rels_map.get(r_id, "")
                target_filename = rel_target.split('/')[-1] if rel_target else ""
                media_keys = [f"/word/{rel_target}", target_filename]
                
                print(f"[DEBUG Image] type 3-fallback)")
                
                _process_and_save_image(element, media_keys, current_store, topic_base_name, "fallback")

        # Parse structured layout tabular content
        elif tag_local == 'tbl':
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

# def parse_document_body_stream2(root, flat_nodes, rels_map, media_parts, images_dir, tables_dir):
#     """Parses paragraphs, tables, and binary shapes out of the layout content stream."""
#     doc_part = root.find(".//pkg:part[@pkg:name='/word/document.xml']", NAMESPACES)
#     if doc_part is None:
#         print("CRITICAL: Master OpenXML document text engine target component missing.")
#         return

#     # FIX: Generate a parent mapping table to support upward traversal without full XPath specifiers
#     print("Indexing layout ancestry trees for deep structural asset tracking...")
#     parent_map = {c: p for p in doc_part.findall('.//*') for c in p}

#     def find_ancestor_sdt_id(element):
#         """Walks up the parent tree to resolve the nearest active section container."""
#         current = element
#         while current in parent_map:
#             current = parent_map[current]
#             tag_local = current.tag.split('}')[-1] if '}' in current.tag else current.tag
#             if tag_local == 'sdt':
#                 sdt_id_node = current.find('.//w:sdtPr/w:id', NAMESPACES)
#                 if sdt_id_node is not None:
#                     found_id = sdt_id_node.attrib.get(f"{{{NAMESPACES['w']}}}val")
#                     if found_id in flat_nodes:
#                         return found_id
#         return None

#     active_node_id = None
    
#     # Process elements sequentially in their native layout order
#     for element in doc_part.findall('.//*'):
#         tag_local = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        
#         # Track active container section context changes linearly
#         if tag_local == 'sdt':
#             sdt_id_node = element.find('.//w:sdtPr/w:id', NAMESPACES)
#             if sdt_id_node is not None:
#                 found_id = sdt_id_node.attrib.get(f"{{{NAMESPACES['w']}}}val")
#                 if found_id in flat_nodes:
#                     active_node_id = found_id

#         # FIX: Leverage ancestor mapping if nested elements or tables bypass linear context tracking
#         current_context_id = active_node_id
#         if not current_context_id:
#             current_context_id = find_ancestor_sdt_id(element)

#         if not current_context_id:
#             continue
            
#         current_store = flat_nodes[current_context_id]
#         topic_base_name = current_store["sanitized_name"]

#         # Parse inline characters
#         if tag_local == 't' and element.text and element.text.strip():
#             text_val = element.text.strip()
#             if text_val not in current_store["text_lines"]:
#                 current_store["text_lines"].append(text_val)

#         # Parse links
#         elif tag_local == 'hyperlink':
#             r_id = element.attrib.get(f"{{{NAMESPACES['r']}}}id")
#             link_text = "".join([t.text for t in element.findall('.//w:t', NAMESPACES) if t.text]).strip()
#             if link_text:
#                 target_url = rels_map.get(r_id, "Internal Jump Link")
#                 current_store["hyperlinks"].append({"text": link_text, "target": target_url})

#         # Parse structural standard images
#         elif tag_local == 'drawing':
#             embed_ids = []
#             for blip in element.findall('.//a:blip', NAMESPACES):
#                 eid = blip.attrib.get(f"{{{NAMESPACES['r']}}}embed")
#                 if eid: embed_ids.append(eid)
#             for pic in element.findall('.//pic:blipFill/a:blip', NAMESPACES):
#                 eid = pic.attrib.get(f"{{{NAMESPACES['r']}}}embed")
#                 if eid and eid not in embed_ids: embed_ids.append(eid)

#             for embed_id in embed_ids:
#                 rel_target = rels_map.get(embed_id, "")
#                 media_keys = [
#                     f"/word/{rel_target}" if not rel_target.startswith("/") else rel_target,
#                     rel_target.split('/')[-1] if rel_target else ""
#                 ]
                
#                 image_bytes = None
#                 for key in media_keys:
#                     if key in media_parts and media_parts[key]:
#                         image_bytes = media_parts[key]
#                         break
                        
#                 if image_bytes:
#                     current_store["image_counter"] += 1
#                     img_filename = f"{topic_base_name[:10]}_{current_store['image_counter']}.png"
#                     img_path = os.path.join(images_dir, img_filename)
                    
#                     with open(img_path, 'wb') as img_f:
#                         img_f.write(image_bytes)
#                     if img_filename not in current_store["extracted_images"]:
#                         current_store["extracted_images"].append(img_filename)

#         # Parse legacy canvas image paths
#         elif tag_local == 'imagedata':
#             r_id = element.attrib.get(f"{{{NAMESPACES['r']}}}id")
#             if r_id:
#                 rel_target = rels_map.get(r_id, "")
#                 target_filename = rel_target.split('/')[-1] if rel_target else ""
#                 image_bytes = media_parts.get(f"/word/{rel_target}") or media_parts.get(target_filename)
                
#                 if image_bytes:
#                     current_store["image_counter"] += 1
#                     img_filename = f"{topic_base_name[:10]}_fallback_{current_store['image_counter']}.png"
#                     img_path = os.path.join(images_dir, img_filename)
                    
#                     with open(img_path, 'wb') as img_f:
#                         img_f.write(image_bytes)
#                     if img_filename not in current_store["extracted_images"]:
#                         current_store["extracted_images"].append(img_filename)

#         # Parse tables
#         elif tag_local == 'tbl':
#             current_store["table_counter"] += 1
#             table_filename = f"{topic_base_name[:10]}_{current_store['table_counter']}.xlsx"
#             single_table_path = os.path.join(tables_dir, table_filename)
            
#             single_wb = openpyxl.Workbook()
#             ws = single_wb.active
#             ws.title = "Extracted Grid Data"
            
#             has_data = False
#             for row in element.findall('.//w:tr', NAMESPACES):
#                 row_cells = []
#                 for cell in row.findall('.//w:tc', NAMESPACES):
#                     cell_str = " ".join([t.text for t in cell.findall('.//w:t', NAMESPACES) if t.text]).strip()
#                     row_cells.append(cell_str)
#                 if any(row_cells):
#                     ws.append(row_cells)
#                     has_data = True
            
#             if has_data:
#                 single_wb.save(single_table_path)
#                 if table_filename not in current_store["extracted_tables"]:
#                     current_store["extracted_tables"].append(table_filename)
#             else:
#                 try: single_wb.close()
#                 except Exception: pass


def compile_hierarchy_tree(flat_nodes):
    """Maps individual items into the final recursive JSON tree format."""
    def assemble_nested_tree(node_id):
        raw_data = flat_nodes[node_id]
        
        # FIXED: Ensure all structural fields transfer cleanly to the output JSON blocks
        node_json = {
            "element_type": raw_data["element_type"],
            "attributes": raw_data["attributes"],
            "extracted_images": raw_data["extracted_images"],
            "extracted_tables": raw_data["extracted_tables"],
            "children_ids": raw_data["children_ids"],
            "image_counter": raw_data["image_counter"],
            "table_counter": raw_data["table_counter"]
        }
        
        text_content = "\n".join(raw_data["text_lines"]).strip()
        if text_content:
            node_json["text_content"] = text_content
        if raw_data["hyperlinks"]:
            node_json["hyperlinks"] = raw_data["hyperlinks"]
            
        if raw_data["children_ids"]:
            node_json["children"] = [assemble_nested_tree(c_id) for c_id in raw_data["children_ids"]]
            
        return node_json

    all_children_ids = set()
    for n in flat_nodes.values():
        all_children_ids.update(n["children_ids"])
        
    root_nodes = [n_id for n_id in flat_nodes if n_id not in all_children_ids]
    return [assemble_nested_tree(r_id) for r_id in root_nodes]

def extract_easa_from_zip_v2(zip_path, storage_path):
    """Executes the pipeline."""
    print(f"Beginning Processing Pipeline: {zip_path}")
    
    paths = resolve_paths(storage_path, zip_path)
    os.makedirs(paths["images_dir"], exist_ok=True)
    os.makedirs(paths["tables_dir"], exist_ok=True)

    root = load_xml_from_zip(zip_path)
    if root is None:
        return

    # Phase 1: Relationship mapping and image parsing
    rels_map, media_parts = extract_relationships_and_media(root)

    # Phase 2: Building structural elements mapping
    er_doc_node = root.find('.//er:document', NAMESPACES)
    if er_doc_node is None:
        print("CRITICAL: Root <er:document> metadata node missing.")
        return

    document_metadata = {
        "guid": er_doc_node.attrib.get('guid', ''),
        "pub-time": er_doc_node.attrib.get('pub-time', ''),
        "source-title": er_doc_node.attrib.get('source-title', ''),
        "Domain": er_doc_node.attrib.get('Domain', '')
    }

    flat_nodes = build_structural_index(er_doc_node)

    # Phase 3: Text narrative stream component loop tracking
    parse_document_body_stream(root, flat_nodes, rels_map, media_parts, paths["images_dir"], paths["tables_dir"])

    # Phase 4: Constructing JSON output object references
    hierarchy_tree = compile_hierarchy_tree(flat_nodes)

    final_output = {
        "document_metadata": document_metadata,
        "rules_hierarchy": hierarchy_tree
    }

    print(f"Saving compiled extraction data to file path: {paths['output_json']}")
    print(f"Saving compiled extraction data to file path: {paths['output_json_com']}")
    with open(paths["output_json"], 'w', encoding='utf-8') as json_f:
        json.dump(final_output, json_f, indent=4, ensure_ascii=False)
    with open(paths["output_json_com"], 'w', encoding='utf-8') as json_f:
        json.dump(final_output, json_f, indent=4, ensure_ascii=False)
        
    print("Execution completed successfully!")


if __name__ == "__main__":
    zip_filename = r"C:\Users\kata_du\Downloads\313A4D_2025-11-27_11.38.35_EAR-for-Initial-Airworthiness-and-Environmental-Protection-Regulation-EU-No-748-2012 (1).zip"
    workspace_dir = r'C:\Users\kata_du\Documents\Literature\EASA\XML _Data_extractions'
    
    extract_easa_from_zip_v2(zip_filename, workspace_dir)