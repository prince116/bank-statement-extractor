################################################################
### Bank Statement Extractor (Standard Chartered - Credit Card)
################################################################
### Author: Prince Wong
### Created at: 26 Oct 2022
### Version: 1.0.0
################################################################

import os
import camelot
import pandas as pd
from decimal import Decimal

def processFile(filePath: str):
    rootDir = os.path.dirname(__file__)
    outputFileName = os.path.splitext(os.path.basename(filePath))[0] + ".xlsx"
    outputPath = os.path.join(rootDir, "output", outputFileName)

    # extract all the tables in the PDF file
    tables = camelot.read_pdf(filePath, "all")

    totalAmount: Decimal = 0
    payBalance: Decimal = 0
    transactionDetails: list = []

    for table in tables:

        # Credit Card Summary
        if table.shape[0] == 4 and table.shape[1] == 4:
            payBalance = Decimal(table.data[1][1].replace(",", ""))

        # Transaction details
        if table.shape[0] == 1 and table.shape[1] == 5:
            transactionList = table.data[0][2].splitlines()
            descriptionList = table.data[0][1].splitlines()
            hkdAmountList = [s.replace(",", "").strip() for s in table.data[0][4].splitlines()]

            transactionList = transactionList[1:]
            if table.page == 1:
                descriptionList = descriptionList[2:]
                hkdAmountList = hkdAmountList[2:]
            elif table.page == 2:
                descriptionList = descriptionList[1:]
                hkdAmountList = hkdAmountList[1:]

            for i in range(0, len(transactionList)):
                if transactionList[i] is not None and "CR" not in hkdAmountList[i]:
                    totalAmount = totalAmount + Decimal(hkdAmountList[i])
                    transactionDetails.append([transactionList[i], descriptionList[i], hkdAmountList[i]])

    if payBalance == totalAmount:
        print("Pay Balance: {}".format(payBalance))

        transactionDetails = sorted(transactionDetails, key=lambda x: x[1])

        excelDf = pd.DataFrame(transactionDetails)

        header: list = [
            "Transaction Reference",
            "Description",
            "HKD Amount"
        ]

        with pd.ExcelWriter(outputPath) as writer:
            print("Writing data to the Excel file ...")
            excelDf.to_excel(writer, sheet_name="Sheet1", index=False, header=header)
            print("Successful!")

    else:
        print("Calculation Error")

    exit()

def main():
    fileName: str = input("Enter File Name: ")
    filePath: str = os.path.join(os.path.abspath(os.getcwd()), fileName)

    if not os.path.exists(filePath):
        print("File dost not exist.")
        return

    processFile(filePath)

if __name__ == "__main__":
    main()