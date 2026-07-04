import json
import csv
import os

import openpyxl


def save_to_excel(headers, rows, output_path):
    """Utility helper to write a data matrix array to an Excel Workbook."""
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


def export_to_cosmograph_csv(json_path):
    """Build Cosmograph node/edge tables from an EASA extraction JSON.

    Emits both CSV and Excel copies of the nodes and edges. Internal topics
    become sized nodes; unmatched hyperlink targets are added as lightweight
    "External Reference" nodes so cross-document links are still visible.
    """
    storage_path = os.path.dirname(json_path)
    nodes_csv_path = os.path.join(storage_path, "cosmograph_nodes.csv")
    edges_csv_path = os.path.join(storage_path, "cosmograph_edges.csv")
    nodes_xlsx_path = os.path.join(storage_path, "cosmograph_nodes.xlsx")
    edges_xlsx_path = os.path.join(storage_path, "cosmograph_edges.xlsx")

    print(f"[INFO] Ingesting targeted JSON data matrix from: {json_path}")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    flat_topics = {}

    # Recursive worker to extract topics based exactly on the extraction schema.
    def gather_topics(nodes_list):
        for node in nodes_list:
            if node.get("element_type") == "topic":
                attribs = node.get("attributes", {})

                # Fetching source-title or title from the attributes mapping
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
                        # Numerical node-sizing metric: length of the topic text
                        "text_len": len(node.get("text_content", "")),
                    }

            if node.get("children"):
                gather_topics(node["children"])

    gather_topics(data.get("rules_hierarchy", []))
    print(f"[INFO] Cataloged {len(flat_topics)} canonical EASA topics from hierarchy.")

    # Sort titles by length (longest first) to prevent sub-string collisions.
    all_internal_titles = sorted(
        [t["title"] for t in flat_topics.values()], key=len, reverse=True
    )

    node_rows = []
    edge_rows = []
    edges_set = set()
    external_discovered_nodes = set()

    # 1. Internal topic nodes
    for t_data in flat_topics.values():
        node_rows.append([
            t_data["title"],          # id
            t_data["title"],          # label
            t_data["domain"],         # Domain (color category A)
            t_data["content_type"],   # TypeOfContent (color category B)
            t_data["text_len"],       # TextSizeValue (numeric node size)
        ])

    # 2. Edges: match hyperlinks to internal topics, else register as external.
    print("[INFO] Evaluating reference matching links and mapping relations...")
    for t_data in flat_topics.values():
        source_node = t_data["title"]

        for link in t_data["hyperlinks"]:
            link_text = link.get("text", "").strip()
            link_target = link.get("target", "Unknown Target").strip()
            if not link_text:
                continue

            matched_target_node = None
            for internal_title in all_internal_titles:
                if source_node == internal_title:
                    continue
                if (internal_title.lower() in link_text.lower()
                        or link_text.lower() in internal_title.lower()):
                    matched_target_node = internal_title
                    break

            if matched_target_node:
                edge_pair = (source_node, matched_target_node, link_target)
                if edge_pair not in edges_set:
                    edges_set.add(edge_pair)
                    edge_rows.append([source_node, matched_target_node, link_target])
            else:
                # Unmatched/external link: treat the link text as its own node.
                external_node_id = link_text
                if (external_node_id not in external_discovered_nodes
                        and external_node_id not in all_internal_titles):
                    external_discovered_nodes.add(external_node_id)
                    node_rows.append([
                        external_node_id,      # id
                        external_node_id,      # label
                        "External Reference",  # Domain category
                        "Hyperlink Fragment",  # TypeOfContent category
                        0,                     # TextSizeValue
                    ])

                edge_pair = (source_node, external_node_id, link_target)
                if edge_pair not in edges_set:
                    edges_set.add(edge_pair)
                    edge_rows.append([source_node, external_node_id, link_target])

    # 3. Export CSV
    node_headers = ["id", "label", "Domain", "TypeOfContent", "TextSizeValue"]
    edge_headers = ["source", "target", "relation"]

    print("[INFO] Exporting dataset rows to standard flat CSV layouts...")
    with open(nodes_csv_path, 'w', newline='', encoding='utf-8') as n_csv:
        csv.writer(n_csv).writerows([node_headers] + node_rows)
    with open(edges_csv_path, 'w', newline='', encoding='utf-8') as e_csv:
        csv.writer(e_csv).writerows([edge_headers] + edge_rows)
    print("[SUCCESS] CSV files generated.")

    # 4. Export identical data to Excel
    print("[INFO] Exporting identical dataset structures into Excel matrix files...")
    save_to_excel(node_headers, node_rows, nodes_xlsx_path)
    save_to_excel(edge_headers, edge_rows, edges_xlsx_path)

    print("\n[SUMMARY] Graph construction completed successfully!")
    print(
        f" -> Total indexed nodes: {len(node_rows)} "
        f"({len(flat_topics)} internal, {len(external_discovered_nodes)} external)"
    )
    print(f" -> Total parsed network connections: {len(edge_rows)}")


# Execute transformation mapping
if __name__ == "__main__":
    json_path = r"C:\Users\kata_du\Documents\Literature\EASA\XML _Data_extractions\Bulk_Extraction\313A4D_2025-11-27_11.38.35_EAR-for-Initial-Airworthiness-and-Environmental-Protection-Regulation-EU-No-748-2012\dbc96709_Extraction_Json.json"
    export_to_cosmograph_csv(json_path)
