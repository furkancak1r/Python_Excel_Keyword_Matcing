# Python Excel Keyword Matcing
 
This script automates the process of matching keywords in an Excel file and updating the values in a specified column. It is particularly useful for categorizing data based on predefined keywords.

## Prerequisites
- Python installed
- Required Python packages installed. Install them using the following command:
  ```bash
  pip install pywin32

## Usage
Clone the repository to your local machine.
```
git clone https://github.com/furkancak1r/Python_Excel_Keyword_Matcing

```
Navigate to the project directory.

```
cd your-repository
```

Ensure that your Excel file is in the specified path (path) and contains a sheet named "Sayfa1".

Create a keywords.json file in the project directory with the following format:
```
{
  "keywords": [
    "keyword1",
    "keyword2",
    "keyword3",
    ...
  ]
}

```
Run the script.

```
python excel_data_matching.py
```