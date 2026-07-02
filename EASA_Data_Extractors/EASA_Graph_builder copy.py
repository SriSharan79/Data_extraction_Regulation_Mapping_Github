import json
import csv
import openpyxl
import os
import re

def save_to_excel(headers, rows, output_path):
    
    """Utility helper to write data matrix arrays to an Excel Workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cosmograph Export"
    
    ws.append(headers)
    for row in rows:
        ws.append(row)
        
    # Auto-fit columns layout formatting
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 50)
        
    wb.save(output_path)
    print(f"[SUCCESS] Generated Excel file at: {output_path}")

def export_to_cosmograph_data(json_path):
    
    
    storage_path=os.path.dirname(json_path)
    nodes_csv_path=os.path.join(storage_path,"cosmograph_nodes.csv")
    edges_csv_path=os.path.join(storage_path,"cosmograph_edges.csv")
    nodes_xlsx_path=os.path.join(storage_path,"cosmograph_nodes.xlsx")
    edges_xlsx_path=os.path.join(storage_path,"cosmograph_edges.xlsx")
    print(f"[INFO] Ingesting targeted JSON data matrix from: {json_path}")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    flat_topics = {}
    
    # 1. Recursive parsing to gather explicit internal document topics
    def gather_topics(nodes_list):
        for node in nodes_list:
            if node.get("element_type") == "topic":
                attribs = node.get("attributes", {})
                source_title = attribs.get("source-title") or attribs.get("title")
                sdt_id = attribs.get("sdt-id") or node.get("id")
                
                if source_title and source_title.strip() and sdt_id:
                    t_title = source_title.strip()
                    flat_topics[sdt_id] = {
                        "id": sdt_id,
                        "title": t_title,
                        "domain": attribs.get("Domain", "Unknown").strip("; "),
                        "content_type": attribs.get("TypeOfContent", "General").strip("; "),
                        "hyperlinks": node.get("hyperlinks", []),
                        "text": node.get("text_content", "")
                    }
            
            if "sub_nodes" in node and node["sub_nodes"]:
                gather_topics(node["sub_nodes"])
            elif "children" in node and node["children"]:
                gather_topics(node["children"])
                
    gather_topics(data.get("rules_hierarchy", []))
    print(f"[INFO] Cataloged {len(flat_topics)} canonical EASA topics.")

    # Sort titles by length descending to ensure specific matches resolve first
    all_internal_titles = sorted([t["title"] for t in flat_topics.values()], key=len, reverse=True)

    # Containers for compiled Node rows and Edge rows
    node_rows = []
    edge_rows = []
    
    # Tracking sets to avoid writing duplicate network edges or nodes
    edges_set = set()
    external_discovered_nodes = set()

    # Populate explicit internal node entries
    for t_data in flat_topics.values():
        node_rows.append([
            t_data["title"],         # id
            t_data["title"],         # label
            t_data["domain"],        # Domain
            t_data["content_type"],  # TypeOfContent
            t_data["text"]       # TextSizeValue
        ])

    # 2. Process connections and isolate external references
    print("[INFO] Processing links, evaluating matches and mapping dynamic relation statuses...")
    for t_data in flat_topics.values():
        source_node = t_data["title"]
        
        for link in t_data["hyperlinks"]:
            link_text = link.get("text", "").strip()
            link_target = link.get("target", "Unknown Target").strip()
            if not link_text:
                continue
                
            matched_target_node = None
            
            # Try to match the link text to a known internal topic title
            for internal_title in all_internal_titles:
                if source_node == internal_title:
                    continue
                if internal_title.lower() in link_text.lower() or link_text.lower() in internal_title.lower():
                    matched_target_node = internal_title
                    break
            
            # If matched internally, link to that topic node
            if matched_target_node:
                edge_pair = (source_node, matched_target_node, link_target)
                if edge_pair not in edges_set:
                    edges_set.add(edge_pair)
                    edge_rows.append([source_node, matched_target_node, link_target])
            else:
                # Unmatched/External Link: Treat the unique text as its own target node
                external_node_id = link_text
                
                # Register this new placeholder node if it hasn't been discovered yet
                if external_node_id not in external_discovered_nodes and external_node_id not in all_internal_titles:
                    external_discovered_nodes.add(external_node_id)
                    node_rows.append([
                        external_node_id,     # id
                        external_node_id,     # label
                        "External Reference", # Domain categorization category
                        "Hyperlink Fragment", # TypeOfContent categorization category
                        0                     # TextSizeValue (0 text content length)
                    ])
                
                edge_pair = (source_node, external_node_id, link_target)
                if edge_pair not in edges_set:
                    edges_set.add(edge_pair)
                    edge_rows.append([source_node, external_node_id, link_target])

    # 3. Export data vectors to CSV format
    node_headers = ["id", "label", "Domain", "TypeOfContent", "TextSizeValue"]
    edge_headers = ["source", "target", "relation"]

    print("[INFO] Exporting dataset rows to standard flat CSV layouts...")
    with open(nodes_csv_path, 'w', newline='', encoding='utf-8') as n_csv:
        csv.writer(n_csv).writerows([node_headers] + node_rows)
    with open(edges_csv_path, 'w', newline='', encoding='utf-8') as e_csv:
        csv.writer(e_csv).writerows([edge_headers] + edge_rows)
    print("[SUCCESS] CSV files generated.")

    # 4. Export identical data vectors to Excel spreadsheet format
    print("[INFO] Converting and exporting identical dataset structures into Excel matrix files...")
    save_to_excel(node_headers, node_rows, nodes_xlsx_path)
    save_to_excel(edge_headers, edge_rows, edges_xlsx_path)
    
    print(f"\n[SUMMARY] Graph construction completed successfully!")
    print(f" -> Total indexed nodes: {len(node_rows)} ({len(flat_topics)} internal, {len(external_discovered_nodes)} external hyperlinks)")
    print(f" -> Total parsed network connections: {len(edge_rows)}")

# Execution Trigger
if __name__ == "__main__":
    json_path = r"C:\Users\kata_du\Documents\Literature\EASA\XML _Data_extractions\Initial-Airworthiness-and-Environmental-Protection-Regulation-EU-No-748-2012\dbc96709_Extraction_Json.json"
    export_to_cosmograph_data(json_path)
