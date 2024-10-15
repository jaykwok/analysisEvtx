import sys
import os
import glob
import json
import logging
from concurrent.futures import ProcessPoolExecutor, as_completed
from evtx import PyEvtxParser
import pandas as pd

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def flatten_dict(d, parent_key="", sep="_"):
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)


def process_evtx_file(file_path, output_dir):
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(output_dir, f"{file_name}.xlsx")

    parser = PyEvtxParser(file_path)
    records = []
    all_fields = set()

    for record in parser.records_json():
        try:
            event_data = json.loads(record["data"])
            flat_record = flatten_dict(event_data)

            # Add standard fields
            flat_record["Event_Record_ID"] = record["event_record_id"]
            flat_record["Event_Timestamp"] = record["timestamp"]

            records.append(flat_record)
            all_fields.update(flat_record.keys())
        except Exception as e:
            logging.error(f"Error processing record in {file_path}: {str(e)}")
            logging.debug(f"Problematic record: {record}")

    if records:
        # Ensure all records have all fields
        for record in records:
            for field in all_fields:
                if field not in record:
                    record[field] = None

        df = pd.DataFrame(records)

        # 按照Event_System_EventRecordID排序
        df = df.sort_values(by="Event_System_EventRecordID")

        # 删除所有行为空白的列
        df = df.dropna(axis=1, how="all")

        # 删除只有一个唯一值的列
        df = df.loc[:, df.nunique() != 1]

        df.to_excel(output_file, index=False, engine="openpyxl")
        logging.info(f"Processed {file_path} and saved to {output_file}")
        logging.info(f"Total fields extracted: {len(df.columns)}")
    else:
        logging.warning(f"No valid records found in {file_path}")


def process_folder(folder_path):
    output_dir = os.path.join(folder_path, "output")
    os.makedirs(output_dir, exist_ok=True)

    evtx_files = glob.glob(os.path.join(folder_path, "*.evtx"))

    with ProcessPoolExecutor() as executor:
        futures = [
            executor.submit(process_evtx_file, evtx_file, output_dir)
            for evtx_file in evtx_files
        ]
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as ex:
                logging.error(f"An error occurred while processing a file: {str(ex)}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <folder_path>")
        sys.exit(1)

    folder_path = sys.argv[1]
    if not os.path.isdir(folder_path):
        print(f"Error: {folder_path} is not a valid directory")
        sys.exit(1)

    process_folder(folder_path)
