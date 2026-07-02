import json
import xmltodict

# Replace with your actual XML filename if it's slightly different
xml_filename = r"C:\Users\kata_du\Downloads\313A4D_2025-11-27_11.38.35_EAR-for-Initial-Airworthiness-and-Environmental-Protection-Regulation-EU-No-748-2012\Easy Access Rules for Initial Airworthiness and Environmental Protection (Regulation (EU) No 7482012) - xml.xml"
json_filename = r'XML _Data_extractions'

print("Reading XML file... This might take a few seconds for a large file.")

with open(xml_filename, 'r', encoding='utf-8') as xml_file:
    xml_content = xml_file.read()
    
    # Convert XML to a Python dictionary
    # xml_attribs=True ensures attributes like "guid" and "pub-time" are kept
    data_dict = xmltodict.parse(xml_content, xml_attribs=True)

print("Writing JSON file...")
with open(json_filename, 'w', encoding='utf-8') as json_file:
    # indent=4 makes the resulting JSON beautifully formatted and readable
    json.dump(data_dict, json_file, indent=4, ensure_ascii=False)

print(f"Success! Converted JSON saved as: {json_filename}")