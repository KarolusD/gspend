# CSV transactions into Excel file with charts
Simple python script which takes your CSV transactions and writes them into Excel


### Requirements
- Python3 (version > 3.6)
```bash
python --version
# OR
python3 -V
```

- Openpyxl [A python library to write/read to Excel file](https://openpyxl.readthedocs.io/en/stable/)
```bash
pip install openpyxl
# OR if you prefer to use your fav package manager (i.e homebrew)
brew install openpyxl 

# you can also SKIP this step and create virtualenv using pipenv
# assuming you have pipenv installed use:
pipenv shell
pipenv install
python3 gspend.py
```

- Transactions history CSV file (implementation assumes it finds certain values: look closely to gspend.py)
```python
WHERE_TRANSACTIONS_START = "#Data operacji"
WHERE_TRANSACTIONS_END = "___End of file___"
# AND
date = datetime.strptime(line[0], "%Y-%m-%d").strftime("%m/%d/%Y")
category = line[3]
amount = float(
    line[4].replace(" PLN", "").replace(",", ".").replace(" ", "")
)
```


