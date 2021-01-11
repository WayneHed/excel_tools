# excel_tools

## 1 Introduction

Frequently used tools for Excel files with .xlsx type support, e.g., converting excel to Json, combining multiple Excel files.

## 2 Prerequisites

- Python 3.x

- Install required packages using pip

  ```bash
  pip install -r requirements.txt
  ```

## 3 Usages 

```bash
python excel_tools.py -h
#
# Frequently used tools for Excel files with .xlsx type support, e.g., converting excel to Json, combining multiple Excel files.
# 
# positional arguments:
#  {2json,combine}       available functions
#
# optional arguments:
#   -h, --help            show this help message and exit
#   -if INPUT_FILE, --input_file INPUT_FILE
#                         path to the input Excel file
#   --dumps               dumps loaded Json to file in the same path with input Excel file
#
# Enjoy this program! :)
```

### 3.1 Converting to Json Usage
Test data can be found in this path: `./data/test-data.xlsx`
```bash
python excel_tools.py -if ./data/test-data.xlsx --dumps
```
After executing above script, a `test-data.json` file can be found in the same path, which reads as followings:
```json
{
  "name": "test-data.xlsx",
  "path": "./data/test-data.xlsx",
  "sheets": [
    {
      "name": "Sheet1",
      "headers": [
        "xlsx-表1-t1",
        "xlsx-表1-t2",
        "xlsx-表1-t3",
        "xlsx-表1-t4",
        "xlsx-表1-t5"
      ],
      "contents": [
        {
          "xlsx-表1-t1": "xlsx-表1-内容21",
          "xlsx-表1-t2": "xlsx-表1-内容22",
          "xlsx-表1-t3": "xlsx-表1-内容23",
          "xlsx-表1-t4": "xlsx-表1-内容24",
          "xlsx-表1-t5": "xlsx-表1-内容25"
        },
        {
          "xlsx-表1-t1": "xlsx-表1-内容31",
          "xlsx-表1-t2": "xlsx-表1-内容32",
          "xlsx-表1-t3": "xlsx-表1-内容33",
          "xlsx-表1-t4": "xlsx-表1-内容34",
          "xlsx-表1-t5": "xlsx-表1-内容35"
        },
        {
          "xlsx-表1-t1": "xlsx-表1-内容41",
          "xlsx-表1-t2": "xlsx-表1-内容42",
          "xlsx-表1-t3": "xlsx-表1-内容43",
          "xlsx-表1-t4": "xlsx-表1-内容44",
          "xlsx-表1-t5": "xlsx-表1-内容45"
        },
        {
          "xlsx-表1-t1": "xlsx-表1-内容51",
          "xlsx-表1-t2": "xlsx-表1-内容52",
          "xlsx-表1-t3": "xlsx-表1-内容53",
          "xlsx-表1-t4": "xlsx-表1-内容54",
          "xlsx-表1-t5": "xlsx-表1-内容55"
        }
      ]
    },
    {
      "name": "Sheet2",
      "headers": [
        "xlsx-表2-t1",
        "xlsx-表2-t2",
        "xlsx-表2-t3",
        "xlsx-表2-t4",
        "xlsx-表2-t5"
      ],
      "contents": [
        {
          "xlsx-表2-t1": "xlsx-表2-内容21",
          "xlsx-表2-t2": "xlsx-表2-内容22",
          "xlsx-表2-t3": "xlsx-表2-内容23",
          "xlsx-表2-t4": "xlsx-表2-内容24",
          "xlsx-表2-t5": "xlsx-表2-内容25"
        },
        {
          "xlsx-表2-t1": "xlsx-表2-内容31",
          "xlsx-表2-t2": "xlsx-表2-内容32",
          "xlsx-表2-t3": "xlsx-表2-内容33",
          "xlsx-表2-t4": "xlsx-表2-内容34",
          "xlsx-表2-t5": "xlsx-表2-内容35"
        },
        {
          "xlsx-表2-t1": "xlsx-表2-内容41",
          "xlsx-表2-t2": "xlsx-表2-内容42",
          "xlsx-表2-t3": "xlsx-表2-内容43",
          "xlsx-表2-t4": "xlsx-表2-内容44",
          "xlsx-表2-t5": "xlsx-表2-内容45"
        },
        {
          "xlsx-表2-t1": "xlsx-表2-内容51",
          "xlsx-表2-t2": "xlsx-表2-内容52",
          "xlsx-表2-t3": "xlsx-表2-内容53",
          "xlsx-表2-t4": "xlsx-表2-内容54",
          "xlsx-表2-t5": "xlsx-表2-内容55"
        }
      ]
    }
  ]
}
```

