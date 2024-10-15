# analysisEvtx

## 更新日志

### 2024/10/15

1、修改读取文件方式，改为读取文件夹，文件夹中自动筛选evtx文件，并输出对应文件至目标文件夹/output/

2、修改读取方式，由xml格式读取改为json格式展开，加快读取速度

3、筛选事件ID并添加事件描述，输出事件中的全量字段，并删除空白列和列值仅有一种的列。筛选的事件ID包含如下部分：

| 事件描述                                       | 事件ID                           |
| ---------------------------------------------- | -------------------------------- |
| 日志清除记录                                   | 1102                             |
| 出入站连接记录                                 | 5156 、5158                      |
| 账户管理（创建/删除）记录                      | 4720 、4726                      |
| 安全组管理（添加/移除）记录                    | 4732 、4733                      |
| 账户行为（登录/注销/失败）记录                 | 4624 、4634 、4625               |
| 凭证验证及特殊登录                             | 4776 、 4672                     |
| 计划任务事件（创建/删除/已启用/已停用/已更新） | 4698 、4699 、4700 、4701 、4702 |
| 进程（创建/终止）记录                          | 4688 、4689                      |
| 注册表修改                                     | 4657                             |



## 项目描述

分析windows日志文件（.evtx），通过写定的Json tag值，将对应的日志内容转换为excel文件，方便进行数据筛选.   

开发语言：python   

主要使用python库：evtx    

如需修改则修改python文件中的process_evtx_file函数

### 文件描述

output：根据给定的事件ID筛选并导出excel文件。

fully_output：导出所有信息至excel文件，删除空白列和列仅有一种值的列（即方差为0）。



## 用法

python output.py <folder_path>

python fully_output.py <folder_path>



## Changelog

### 2024/10/15

1. Modified the file reading method to read from a folder. The program automatically filters EVTX files in the folder and outputs the corresponding files to the target folder `/output/`.
2. Changed the reading method from XML format to JSON format for faster processing.
3. Filtered event IDs and added event descriptions. Outputs all fields from the events, removing empty columns and columns with only one unique value. The filtered event IDs include the following:

| Event Description                                           | Event ID                     |
| ----------------------------------------------------------- | ---------------------------- |
| Log clearing record                                         | 1102                         |
| Inbound/outbound connection record                          | 5156, 5158                   |
| Account management (creation/deletion) record               | 4720, 4726                   |
| Security group management (addition/removal)                | 4732, 4733                   |
| Account activity (login/logout/failure) record              | 4624, 4634, 4625             |
| Credential validation and special login                     | 4776, 4672                   |
| Scheduled task events (create/delete/enable/disable/update) | 4698, 4699, 4700, 4701, 4702 |
| Process (create/terminate) record                           | 4688, 4689                   |
| Registry modification                                       | 4657                         |

## Project Description

This project analyzes Windows event log files (.evtx). Using predefined JSON tag values, it converts the corresponding log contents into an Excel file for easier data filtering.

Programming language: Python

Primary Python library: `evtx`

To make modifications, update the `process_evtx_file` function in the Python file.

### File Description

- `output`: Exports filtered Excel files based on the given event IDs.
- `fully_output`: Exports all information to an Excel file, removing empty columns and columns with only one unique value (i.e., columns with zero variance).

## Usage

```
python output.py <folder_path>

python fully_output.py <folder_path>
```
