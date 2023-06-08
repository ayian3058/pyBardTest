# app.py
from flask import Flask, render_template, request

import openpyxl

app = Flask(__name__)

@app.route("/")
def index():
    # Get the dropdown selections from the user
    drop1 = request.args.get("drop1")
    drop2 = request.args.get("drop2")
    drop3 = request.args.get("drop3")
    drop4 = request.args.get("drop4")

    # Create an Excel document with the user's selections
    excel_document = create_excel_document(drop1, drop2, drop3, drop4)

    # Return the Excel document to the user
    return excel_document

def create_excel_document(drop1, drop2, drop3, drop4):
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Create a new worksheet
    worksheet = workbook.create_sheet("Sheet1")

    # Write the user's selections to the worksheet
    worksheet.write(0, 0, drop1)
    worksheet.write(1, 0, drop2)
    worksheet.write(2, 0, drop3)
    worksheet.write(3, 0, drop4)

    # Save the Excel workbook
    workbook.save("excel_document.xlsx")

    # Return the Excel workbook to the user
    return excel_document

if __name__ == "__main__":
    app.run(debug=True)