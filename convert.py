import os
import json
import pandas as pd
import re

# ✅ Folder containing all JSON files
input_folder = r"C:\JSON_Files"
output_file = r"C:\JSON_Files\FinalOutput1.xlsx"

# ✅ Define the required fields
selected_fields = [
    "manufacturer_contact_state",
    "manufacturer_g1_city",
    "manufacturer_contact_address_1",
    "manufacturer_contact_pcity",
    "event_type",
    "report_number",
    "type_of_report",
    "product_problem_flag",
    "date_received",
    "manufacturer_address_2",
    "pma_pmn_number",
    "date_of_event",
    "manufacturer_contact_zip_code",
    "source_type",
    "brand_name",
    "generic_name",
    "manufacturer_d_name",
    "manufacturer_d_address_1",
    "manufacturer_d_address_2",
    "manufacturer_d_city",
    "manufacturer_d_state",
    "manufacturer_d_zip_code",
    "manufacturer_d_country",
    "manufacturer_d_postal_code",
    "device_operator",
    "model_number",
    "catalog_number",
    "lot_number",
    "device_class",
    "product_problems",
    "patient_problems",
    "date_changed",
    "initial_report_to_fda",
    "mdr_text_key",
    "text_type_code",
    "patient_sequence_number",
    "text"
]

# ✅ Function to remove illegal characters
def clean_text(value):
    if isinstance(value, str):
        value = re.sub(r"[^\x20-\x7E]", "", value)  # Remove non-printable characters
        return value.replace("♥", "").replace("¿", "").replace("�", "")
    return value

# ✅ List to store all extracted data
all_data = []

# ✅ Function to process each JSON file
def process_json_file(json_file):
    input_file_path = os.path.join(input_folder, json_file)

    print(f"Processing: {json_file}...")

    try:
        # Load JSON file
        with open(input_file_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Extract the "results" section
        results = data.get("results", [])

        # Extract and flatten data
        for entry in results:
            extracted_data = {field: clean_text(entry.get(field, "")) for field in selected_fields}

            # Extract device-related fields (handling nested structure)
            if "device" in entry and isinstance(entry["device"], list):
                for device in entry["device"]:
                    extracted_data.update({
                        "brand_name": clean_text(device.get("brand_name", "")),
                        "generic_name": clean_text(device.get("generic_name", "")),
                        "manufacturer_d_name": clean_text(device.get("manufacturer_d_name", "")),
                        "manufacturer_d_address_1": clean_text(device.get("manufacturer_d_address_1", "")),
                        "manufacturer_d_address_2": clean_text(device.get("manufacturer_d_address_2", "")),
                        "manufacturer_d_city": clean_text(device.get("manufacturer_d_city", "")),
                        "manufacturer_d_state": clean_text(device.get("manufacturer_d_state", "")),
                        "manufacturer_d_zip_code": clean_text(device.get("manufacturer_d_zip_code", "")),
                        "manufacturer_d_country": clean_text(device.get("manufacturer_d_country", "")),
                        "manufacturer_d_postal_code": clean_text(device.get("manufacturer_d_postal_code", "")),
                        "device_operator": clean_text(device.get("device_operator", "")),
                        "model_number": clean_text(device.get("model_number", "")),
                        "catalog_number": clean_text(device.get("catalog_number", "")),
                        "lot_number": clean_text(device.get("lot_number", "")),
                        "device_class": clean_text(device.get("openfda", {}).get("device_class", "")),
                    })
                    all_data.append(extracted_data.copy())  # Store each device entry separately

            # Extract product problems
            extracted_data["product_problems"] = clean_text(", ".join(entry.get("product_problems", [])))

            # Extract patient problems
            if "patient" in entry and isinstance(entry["patient"], list):
                patient_problems = []
                for patient in entry["patient"]:
                    patient_problems.extend(patient.get("patient_problems", []))
                extracted_data["patient_problems"] = clean_text(", ".join(patient_problems))

            # Extract text fields from mdr_text
            if "mdr_text" in entry and isinstance(entry["mdr_text"], list):
                for text_entry in entry["mdr_text"]:
                    extracted_data.update({
                        "mdr_text_key": clean_text(text_entry.get("mdr_text_key", "")),
                        "text_type_code": clean_text(text_entry.get("text_type_code", "")),
                        "patient_sequence_number": clean_text(text_entry.get("patient_sequence_number", "")),
                        "text": clean_text(text_entry.get("text", "")),
                    })
                    all_data.append(extracted_data.copy())

    except Exception as e:
        print(f"❌ Error processing {json_file}: {e}")

# ✅ Loop through all JSON files in the folder
for file in os.listdir(input_folder):
    if file.endswith(".json"):
        process_json_file(file)

# ✅ Convert to Pandas DataFrame
df = pd.DataFrame(all_data)

# ✅ Check total rows and split if necessary
max_rows_per_sheet = 1000000  # Safe limit for Excel
num_sheets = (len(df) // max_rows_per_sheet) + 1

# ✅ Save to multiple sheets if too large
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for i in range(num_sheets):
        start_row = i * max_rows_per_sheet
        end_row = (i + 1) * max_rows_per_sheet
        df[start_row:end_row].to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

print(f"🎉 Final Excel file created successfully: {output_file} (Split into {num_sheets} sheets)")
