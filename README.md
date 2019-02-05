# Format Excel
This script is used to format our weekly SSL report

# Sample Data
`sample.xlsx` is a raw ssl file exported from Digicert domain services with sensitive information stripped out

# Dependency
I use the popular python library [openpyxl](https://openpyxl.readthedocs.io/en/stable/) to parse the excel file and manipulate data.

```bash
pip install openpyxl # Python 2
pip3 install openpyxl # Python 3
```

# Usage
Run the python script with the argument of the sample excel file. The output will be an excel file with the name `ssl_datetimestamp.xlsx`
```bash
git clone https://github.com/zaidongy/python-excel.git
cd python-excel
python3 script.py sample.xlsx
```

# Author
  [Chris Yang](https://chrisyang.io) (EIS Systems Analyst, Cedars-Sinai)
