import json
import csv
import os
import re

def export_to_cosmograph_csv(json_path,):
    storage_path=os.path.dirname(json_path)
    nodes_csv_path=os.path.join(storage_path,"cosmograph_nodes.csv")
    edges_csv_path=os.path.join(storage_path,"cosmograph_edges.csv")
    print(f"[INFO] Ingesting targeted JSON data matrix from: {json_path}")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    flat_topics = {}
    
    # Recursive worker to extract topics based exactly on your schema keys
    def gather_topics(nodes_list):
        for node in nodes_list:
            if node.get("element_type") == "topic":
                attribs = node.get("attributes", {})
                
                # Fetching source-title or title from attributes dictionary mapping
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
                        # Dynamically calculating text string lengths from your unified text_content field
                        "text_len": len(node.get("text_content", ""))
                    }
            
            # Unrolling structural hierarchy child loops securely
            if "sub_nodes" in node and node["sub_nodes"]:
                gather_topics(node["sub_nodes"])
            elif "children" in node and node["children"]:
                gather_topics(node["children"])
                
    gather_topics(data.get("rules_hierarchy", []))
    print(f"[INFO] Cataloged {len(flat_topics)} canonical EASA topics from hierarchy.")

    # Sorting canonical titles by length (longest first) to prevent sub-string collision mismatches
    all_titles = sorted([t["title"] for t in flat_topics.values()], key=len, reverse=True)

    # 1. Generate the NODES Table File
    print(f"[INFO] Compiling and writing Node rows to: {nodes_csv_path}")
    with open(nodes_csv_path, 'w', newline='', encoding='utf-8') as n_file:
        writer = csv.writer(n_file)
        writer.writerow(["id", "label", "Domain", "TypeOfContent", "TextSizeValue"])
        
        for t_data in flat_topics.values():
            writer.writerow([
                t_data["title"],         # unique ID reference used by cosmograph
                t_data["title"],         # display label string
                t_data["domain"],        # grouping color category A
                t_data["content_type"],  # grouping color category B
                t_data["text_len"]       # numerical node sizing metric value
            ])

    # 2. Compute and Generate the EDGES Table File via String Mapping
    print(f"[INFO] Evaluating reference matching links and writing Edge paths to: {edges_csv_path}")
    edges_set = set()
    
    with open(edges_csv_path, 'w', newline='', encoding='utf-8') as e_file:
        writer = csv.writer(e_file)
        writer.writerow(["source", "target", "relation"])
        
        for t_data in flat_topics.values():
            source_node = t_data["title"]
            
            for link in t_data["hyperlinks"]:
                link_text = link.get("text", "").strip()
                if not link_text:
                    continue
                    
                for target_node in all_titles:
                    # Avoid establishing reflexive loops back to the same parent node
                    if source_node == target_node:
                        continue
                        
                    # Connectivity confirmation rule: matching containment constraints case-insensitively
                    if target_node.lower() in link_text.lower() or link_text.lower() in target_node.lower():
                        edge_pair = (source_node, target_node)
                        if edge_pair not in edges_set:
                            edges_set.add(edge_pair)
                            writer.writerow([source_node, target_node, "hyperlink_reference"])
                        break # Break out early once match resolved to preserve link specificity

    print(f"[SUCCESS] Task execution completed. Discovered {len(edges_set)} unique structural cross-connections.")

# Execute transformation mapping
if __name__ == "__main__":
    json_path = r"C:\Users\kata_du\Documents\Literature\EASA\XML _Data_extractions\Bulk_Extraction\313A4D_2025-11-27_11.38.35_EAR-for-Initial-Airworthiness-and-Environmental-Protection-Regulation-EU-No-748-2012\dbc96709_Extraction_Json.json"
    export_to_cosmograph_csv(json_path)
