import hashlib
import os
import json
import base64
import re
import zipfile
import io
import xml.etree.ElementTree as ET
import openpyxl  # Ensure you run 'pip3 install openpyxl'

def sanitize_filename(name):
    """Converts a topic title into a safe, clean string usable as a filename."""
    if not name:
        return "Unknown_Topic"
    # Replace non-alphanumeric characters with underscores
    clean = re.sub(r'[^a-zA-Z0-9\s\-\.]', '_', name)
    # Collapse multiple spaces or underscores
    clean = re.sub(r'[\s_]+', '_', clean)
    return clean.strip('_')[:120]  # Truncate to prevent reaching OS file path limits

def resolve_paths(storage_path, zip_file_path):
    """Generates the structured workspace file layout paths."""
    doc_name = os.path.splitext(os.path.basename(zip_file_path))[0]
    target_root = os.path.join(storage_path, doc_name)
    base_hash = hashlib.md5(doc_name.encode()).hexdigest()[:8]
    
    return {
        "images_dir": os.path.join(target_root, "images"),
        "tables_dir": os.path.join(target_root, "tables"),
        "output_json": os.path.join(target_root, f"{base_hash}_Extraction_Json.json")
    }

def extract_easa_from_zip_v2(zip_path, storage_path):
    print(f"Opening ZIP Archive: {zip_path}")
    
    namespaces = {
        'pkg': 'http://schemas.microsoft.com/office/2006/xmlPackage',
        'er': 'http://www.easa.europa.eu/erules-export',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'v': 'urn:schemas-microsoft-com:vml'
    }
    
    paths = resolve_paths(storage_path, zip_path)
    images_dir = paths["images_dir"]
    tables_dir = paths["tables_dir"]
    output_json = paths["output_json"]
    
    os.makedirs(images_dir, exist_ok=True)
    os.makedirs(tables_dir, exist_ok=True)

    # 1. Extract and read the raw XML payload from the ZIP file
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            xml_target_name = [f for f in z.namelist() if f.endswith('.xml')][0]
            print(f"Decompressing payload: {xml_target_name}")
            xml_bytes = z.read(xml_target_name)
    except Exception as e:
        print(f"CRITICAL: Failed to unpack ZIP container: {e}")
        return

    print("Parsing XML package tree structure...")
    root = ET.parse(io.BytesIO(xml_bytes)).getroot()

    # --- PHASE 1: Capture Relationships & Decoded Media Binaries ---
    print("Mapping resource hyperlinks and base64 media blocks...")
    rels_map = {}
    media_parts = {}
    
    for part in root.findall('.//pkg:part', namespaces):
        part_name = part.attrib.get(f"{{{namespaces['pkg']}}}name", "")
        
        # Capture Resource Relations Map
        if "document.xml.rels" in part_name or part_name.endswith(".rels"):
            for rel in part.findall('.//pkg:xmlData//Relationship', namespaces):
                rels_map[rel.attrib.get('Id')] = rel.attrib.get('Target')
                
        # Decode base64 Binary Assets
        if "/word/media/" in part_name:
            string_data = part.find('.//pkg:stringData', namespaces)
            if string_data is not None and string_data.text:
                try:
                    media_parts[part_name] = base64.b64decode(string_data.text.strip())
                except Exception:
                    pass

    # Map flat filenames as lookups if fallback schemes are used
    for part_name in list(media_parts.keys()):
        filename = part_name.split('/')[-1]
        media_parts[filename] = media_parts[part_name]

    # --- PHASE 2: Index Structural Constraints & Mapping References ---
    er_doc_node = root.find('.//er:document', namespaces)
    if er_doc_node is None:
        print("CRITICAL: Root <er:document> metadata node missing.")
        return

    document_metadata = {
        "guid": er_doc_node.attrib.get('guid', ''),
        "pub-time": er_doc_node.attrib.get('pub-time', ''),
        "source-title": er_doc_node.attrib.get('source-title', ''),
        "Domain": er_doc_node.attrib.get('Domain', '')
    }

    easa_attributes = [
        'sdt-id', 'source-title', 'ERulesId', 'Domain', 'ActivityType', 
        'AircraftCategory', 'AircraftUse', 'AmendedBy', 'ApplicabilityDate', 
        'EntryIntoForceDate', 'TypeOfContent', 'ParentIR', 'EASACategory', 'title'
    ]

    flat_nodes = {}
    
    # Pre-populate all metadata configurations from the TOC tree map
    for element in er_doc_node.findall('.//*[@sdt-id]'):
        sdt_id = element.attrib.get('sdt-id')
        tag_name = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        
        attribs = {}
        for attr in easa_attributes:
            if attr in element.attrib and element.attrib[attr].strip():
                attribs[attr] = element.attrib[attr].strip()
                
        # Resolve a descriptive base name for assets belonging to this node
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

    # Map tree hierarchy dependency links
    def map_tree_dependencies(element):
        sdt_id = element.attrib.get('sdt-id')
        for child in element:
            child_sdt_id = child.attrib.get('sdt-id')
            if sdt_id and child_sdt_id and child_sdt_id in flat_nodes:
                if child_sdt_id not in flat_nodes[sdt_id]["children_ids"]:
                    flat_nodes[sdt_id]["children_ids"].append(child_sdt_id)
            map_tree_dependencies(child)

    map_tree_dependencies(er_doc_node)

    # --- PHASE 3: Physical Document Body Stream Parsing ---
    print("Parsing linear document body stream matching content components...")
    doc_part = root.find(".//pkg:part[@pkg:name='/word/document.xml']", namespaces)
    if doc_part is None:
        print("CRITICAL: Failed to locate central text stream body part.")
        return

    active_node_id = None
    
    for element in doc_part.findall('.//*'):
        tag_local = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        
        if tag_local == 'sdt':
            sdt_id_node = element.find('.//w:sdtPr/w:id', namespaces)
            if sdt_id_node is not None:
                found_id = sdt_id_node.attrib.get(f"{{{namespaces['w']}}}val")
                if found_id in flat_nodes:
                    active_node_id = found_id

        if not active_node_id:
            continue
            
        current_store = flat_nodes[active_node_id]
        topic_base_name = current_store["sanitized_name"]

        # Parse Text Elements
        if tag_local == 't' and element.text and element.text.strip():
            text_val = element.text.strip()
            if text_val not in current_store["text_lines"]:
                current_store["text_lines"].append(text_val)

        # Parse Hyperlinks
        elif tag_local == 'hyperlink':
            r_id = element.attrib.get(f"{{{namespaces['r']}}}id")
            link_text = "".join([t.text for t in element.findall('.//w:t', namespaces) if t.text]).strip()
            if link_text:
                target_url = rels_map.get(r_id, "Internal Reference")
                current_store["hyperlinks"].append({"text": link_text, "target": target_url})

        # Parse Images (Drawings)
        elif tag_local == 'drawing':
            embed_ids = []
            for blip in element.findall('.//a:blip', namespaces):
                eid = blip.attrib.get(f"{{{namespaces['r']}}}embed")
                if eid: embed_ids.append(eid)
            for pic in element.findall('.//pic:blipFill/a:blip', namespaces):
                eid = pic.attrib.get(f"{{{namespaces['r']}}}embed")
                if eid and eid not in embed_ids: embed_ids.append(eid)

            for embed_id in embed_ids:
                rel_target = rels_map.get(embed_id, "")
                media_keys = [
                    f"/word/{rel_target}" if not rel_target.startswith("/") else rel_target,
                    rel_target.split('/')[-1] if rel_target else ""
                ]
                
                image_bytes = None
                for key in media_keys:
                    if key in media_parts and media_parts[key]:
                        image_bytes = media_parts[key]
                        break
                        
                if image_bytes:
                    current_store["image_counter"] += 1
                    # Append sequential index if multiple elements exist
                    img_filename = f"{topic_base_name}_{current_store['image_counter']}.png"
                    img_path = os.path.join(images_dir, img_filename)
                    
                    with open(img_path, 'wb') as img_f:
                        img_f.write(image_bytes)
                    if img_path not in current_store["extracted_images"]:
                        current_store["extracted_images"].append(img_path)

        # Parse Fallback Legacy Imagery Elements
        elif tag_local == 'imagedata':
            r_id = element.attrib.get(f"{{{namespaces['r']}}}id")
            if r_id:
                rel_target = rels_map.get(r_id, "")
                target_filename = rel_target.split('/')[-1] if rel_target else ""
                image_bytes = media_parts.get(f"/word/{rel_target}") or media_parts.get(target_filename)
                
                if image_bytes:
                    current_store["image_counter"] += 1
                    img_filename = f"{topic_base_name[:10]}_fallback_{current_store['image_counter']}.png"
                    img_path = os.path.join(images_dir, img_filename)
                    
                    with open(img_path, 'wb') as img_f:
                        img_f.write(image_bytes)
                    if img_path not in current_store["extracted_images"]:
                        current_store["extracted_images"].append(img_path)

        # Parse Grid Tables
        elif tag_local == 'tbl':
            current_store["table_counter"] += 1
            table_filename = f"{topic_base_name[:10]}_{current_store['table_counter']}.xlsx"
            single_table_path = os.path.join(tables_dir, table_filename)
            
            # Create a separate fresh workbook for this table
            single_wb = openpyxl.Workbook()
            ws = single_wb.active
            ws.title = "Extracted Data"
            
            has_data = False
            for row in element.findall('.//w:tr', namespaces):
                row_cells = []
                for cell in row.findall('.//w:tc', namespaces):
                    cell_str = " ".join([t.text for t in cell.findall('.//w:t', namespaces) if t.text]).strip()
                    row_cells.append(cell_str)
                if any(row_cells):
                    ws.append(row_cells)
                    has_data = True
            
            if has_data:
                single_wb.save(single_table_path)
                tbl_info = {"excel_file": single_table_path, "sheet_name": "Extracted Data"}
                if tbl_info not in current_store["extracted_tables"]:
                    current_store["extracted_tables"].append(tbl_info)

    # --- PHASE 4: Build Nested Hierarchical JSON Tree Array ---
    print("Assembling structured tree output format...")
    
    def assemble_nested_tree(node_id):
        raw_data = flat_nodes[node_id]
        
        node_json = {
            "element_type": raw_data["element_type"],
            "attributes": raw_data["attributes"]
        }
        
        text_content = "\n".join(raw_data["text_lines"]).strip()
        if text_content:
            node_json["text_content"] = text_content
        if raw_data["hyperlinks"]:
            node_json["hyperlinks"] = raw_data["hyperlinks"]
        if raw_data["extracted_images"]:
            node_json["extracted_images"] = raw_data["extracted_images"]
        if raw_data["extracted_tables"]:
            node_json["extracted_tables"] = raw_data["extracted_tables"]
            
        if raw_data["children_ids"]:
            node_json["children"] = [assemble_nested_tree(c_id) for c_id in raw_data["children_ids"]]
            
        return node_json

    all_children_ids = set()
    for n in flat_nodes.values():
        all_children_ids.update(n["children_ids"])
        
    root_nodes = [n_id for n_id in flat_nodes if n_id not in all_children_ids]
    hierarchy_tree = [assemble_nested_tree(r_id) for r_id in root_nodes]

    final_output = {
        "document_metadata": document_metadata,
        "rules_hierarchy": hierarchy_tree
    }

    print(f"Writing fully aligned JSON tree compilation to: {output_json}")
    with open(output_json, 'w', encoding='utf-8') as json_f:
        json.dump(final_output, json_f, indent=4, ensure_ascii=False)
        
    print("Conversion completed successfully!")

if __name__ == "__main__":
    zip_filename = r"C:\Users\kata_du\Downloads\313A4D_2025-11-27_11.38.35_EAR-for-Initial-Airworthiness-and-Environmental-Protection-Regulation-EU-No-748-2012 (1).zip"
    workspace_dir = r'C:\Users\kata_du\Documents\Literature\EASA\XML _Data_extractions'
    
    extract_easa_from_zip_v2(zip_filename, workspace_dir)