## AliExpresPaymentsParser
Simple tool to parse some imformation from aliExpress payments pdf
Input: directory filled with pdfs with name format 999999999999_payment.pdf
Output: excel file with colums: Date, without taxes, taxes, and total
Be carefull while using if your pdf is formatted differently it will not work.
### Dependecies:
- Python (tested with 3.10 should probably work for all versions >= 3.7)
- [openpyxl](https://pypi.org/project/openpyxl/)
- [pdfreader](https://pypi.org/project/pdfreader/)
