# VendorInvoiceMessageParser
This script is used to extract the Vendor Invoice information from SAP Staging database for further analysis.

## Prerequisite
Before you start the engine, you'd need to intall some Python packages.

**PyODBC**

A Python DB API 2 module for ODBC. This project provides an up-to-date, convenient interface to ODBC using native data types like datetime and decimal.

You can install the package with PIP, or download the package from https://pypi.org/project/pyodbc/

**Beautiful Soup**

Beautiful Soup is a library that makes it easy to scrape information from web pages. It sits atop an HTML or XML parser, providing Pythonic idioms for iterating, searching, and modifying the parse tree.

You can install the package with PIP, or download the package from https://pypi.org/project/beautifulsoup4/

**openpyxl**

openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.
It was born from lack of existing library to read/write natively from Python the Office Open XML format.
All kudos to the PHPExcel team as openpyxl was initially based on PHPExcel.

You can install the package with PIP, or download the package from https://pypi.org/project/openpyxl/

## Usage
You'd need to specify the database connection information in *AppSettings.xml*, and update the validation rule for TypeOfInvoice and DocumentType is necessary.

You may need to change the query criteria specified by SQL script listed underneath:
```sql
checkSql = '''select top 3 * from
(select 
s.FileID,
s.ProcessDate,
s.ProcessResult,
s.MessageID,
s.FileLength,
s.Data.value('declare namespace n0="http://finance.group.volvo.com/vendorinvoicedetails/1_0";
(/n0:SyncVendorInvoiceDetails/VendorInvoiceDetails/VendorInvoiceHeader/CompanyCode)[1]', 'nvarchar(20)') AS CompanyCode,
s.Data
from satbSourceData s
where s.FileType = 'VendorInvoiceDetails') as t
where ProcessResult = 'Error'
    and MessageID <> '414D512046524F444F5F423120202020AA90F15B24C61FEF'
order by T.ProcessDate desc'''
```
Considering the extraction could be time consuming, it may not be a good idea to take too many records involved.
