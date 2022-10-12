from webbrowser import get
import PyPDF2
import xlsxwriter
from dateutil.parser import parse


def main_function():

    filePathScotia = 'Scotia_September_Statement.pdf'
    parse_Scotia_Pdf(filePathScotia)


def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try:
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False


def parse_Scotia_Pdf(filePath):
    transactionDict = {}
    pdfFileObj = open(filePath, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    for pageNo in range(0, pdfReader.numPages):
        lineCounter = 0

        if pageNo != 1:
            pageObj = pdfReader.getPage(pageNo)

            statementPageLines = pageObj.extract_text().split('\n')
            for item in statementPageLines:
                if pageNo == 0:

                    if "Statement Period" in item and pageNo == 0 and statementPageLines[lineCounter-8] == "SCENE":
                        # We can get statement period here
                        print(item)
                        print(statementPageLines[lineCounter+1])
                        print(statementPageLines[lineCounter+3])

                    if "Continued on page" in item:
                        pass

                    if "AMOUNT($)" in item:

                        for firstTrans in range(lineCounter, len(statementPageLines)):

                            if statementPageLines[firstTrans] == "001" and is_date(statementPageLines[firstTrans+1]):

                                statementLineCounter = 0
                                for i in range(firstTrans, len(statementPageLines)):

                                    if statementPageLines[i].isnumeric() and int(statementPageLines[firstTrans]) + (statementLineCounter) == int(statementPageLines[i]):

                                        transactionElements = []
                                        for j in range(0, 9):

                                            if statementPageLines[i+j].isnumeric() and int(statementPageLines[firstTrans])+statementLineCounter + 1 == int(statementPageLines[i+j]):

                                                break
                                            if statementPageLines[i+j] == "If you have any questions about this":
                                                break
                                            if j != 0:

                                                transactionElements.append(
                                                    statementPageLines[i+j])

                                        transactionDict.update(
                                            {int(statementPageLines[i]): transactionElements})
                                        tmp = statementPageLines[i]

                                        statementLineCounter += 1
                        lastTrans = tmp

                    lineCounter += 1
                else:

                    if "Transactions - continued" in item:

                        for firstTrans in range(lineCounter, len(statementPageLines)):

                            if statementPageLines[firstTrans].isnumeric() and int(statementPageLines[firstTrans]) == int(lastTrans)+1 and is_date(statementPageLines[firstTrans+1]):

                                statementLineCounter = 0
                                for i in range(firstTrans, len(statementPageLines)):

                                    if statementPageLines[i].isnumeric() and int(statementPageLines[firstTrans]) + (statementLineCounter) == int(statementPageLines[i]):

                                        transactionElements = []
                                        for j in range(0, 9):

                                            if statementPageLines[i+j].isnumeric() and int(statementPageLines[firstTrans])+statementLineCounter + 1 == int(statementPageLines[i+j]):
                                                break
                                            if statementPageLines[i+j] == "If you have any questions about this":
                                                break
                                            if j != 0:

                                                transactionElements.append(
                                                    statementPageLines[i+j])

                                        transactionDict.update(
                                            {int(statementPageLines[i]): transactionElements})
                                        tmp = statementPageLines[i]
                                        statementLineCounter += 1
                        lastTrans = tmp

                    lineCounter += 1
    print("\n")
    pdfFileObj.close()
    xlsxGenerator(transactionDict, "ScotiabankTransactions.xlsx")


def xlsxGenerator(transactionDict, filename):
    # converting the dictionary file to the excel file
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "Reference Number")
    worksheet.write(0, 1, "Transaction Date")
    worksheet.write(0, 2, "Transaction Post")
    worksheet.write(0, 3, "Description")
    worksheet.write(0, 4, "Amount(CAD)")

    row = 1
    for key in transactionDict.keys():
        worksheet.write(row, 0, key)
        worksheet.write(row, 1, transactionDict[key][0])
        worksheet.write(row, 2, transactionDict[key][1])
        if transactionDict[key][-1] == "-":
            ptr = 1
        else:
            ptr = 0
        description = ""
        for i in range(2, len(transactionDict[key])-ptr-1):
            description = description + " " + transactionDict[key][i]
        worksheet.write(row, 3, description)
        if ptr == 1:
            amount = transactionDict[key][-1] + transactionDict[key][-2]
        else:
            amount = transactionDict[key][-1]
        worksheet.write(row, 4, float(amount))
        row += 1

    workbook.close()
    print('''"''' + filename + '''"''' + " has been generated successfully!")


def parse_Neo_Pdf(filePath):
    pdfFileObj = open(filePath, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    for pageNo in range(0, pdfReader.numPages):
        if pageNo != 1:
            pageObj = pdfReader.getPage(pageNo)

    pdfFileObj.close()


if __name__ == '__main__':

    main_function()
