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

# 事件ID和描述的映射
event_descriptions = {
    1102: "日志清除记录",
    5156: "出入站连接记录",
    5158: "出入站连接记录",
    4720: "账户管理（创建/删除）记录",
    4726: "账户管理（创建/删除）记录",
    4732: "安全组管理（添加/移除）记录",
    4733: "安全组管理（添加/移除）记录",
    4624: "账户行为（登录/注销/失败）记录",
    4634: "账户行为（登录/注销/失败）记录",
    4625: "账户行为（登录/注销/失败）记录",
    4776: "凭证验证及特殊登录",
    4672: "凭证验证及特殊登录",
    4698: "计划任务事件（创建/删除/已启用/已停用/已更新）",
    4699: "计划任务事件（创建/删除/已启用/已停用/已更新）",
    4700: "计划任务事件（创建/删除/已启用/已停用/已更新）",
    4701: "计划任务事件（创建/删除/已启用/已停用/已更新）",
    4702: "计划任务事件（创建/删除/已启用/已停用/已更新）",
    4688: "进程（创建/终止）记录",
    4689: "进程（创建/终止）记录",
    4657: "注册表修改",
}


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

    for record in parser.records_json():
        try:
            event_data = json.loads(record["data"])
            flat_record = flatten_dict(event_data)

            # 添加标准字段
            flat_record["Event_Record_ID"] = record["event_record_id"]
            flat_record["Event_Timestamp"] = record["timestamp"]

            # 添加事件ID和事件描述
            event_id = int(flat_record.get("Event_System_EventID", 0))
            flat_record["Event_ID"] = event_id
            flat_record["Event_Description"] = event_descriptions.get(event_id, "")

            # 只保留指定事件ID的记录
            if event_id in event_descriptions:
                records.append(flat_record)
        except Exception as e:
            logging.error(f"Error processing record in {file_path}: {str(e)}")
            logging.debug(f"Problematic record: {record}")

    if records:
        df = pd.DataFrame(records)

        # 按Event_System_EventRecordID排序
        df = df.sort_values(by="Event_System_EventRecordID")

        # 删除所有行为空白的列
        df = df.dropna(axis=1, how="all")

        # 删除只有一个唯一值的列
        df = df.loc[:, df.nunique() != 1]

        # 重新排列列，使Event_ID和Event_Description在前面
        columns = ["Event_ID", "Event_Description"] + [
            col for col in df.columns if col not in ["Event_ID", "Event_Description"]
        ]
        df = df[columns]

        df.to_excel(output_file, index=False, engine="openpyxl")
        logging.info(f"Processed {file_path} and saved to {output_file}")
        logging.info(f"Total records extracted: {len(df)}")
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
