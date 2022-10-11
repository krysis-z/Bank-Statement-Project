from webbrowser import get
import PyPDF2
import xlsxwriter


def main_function():

    filePathScotia = 'statement-34.pdf'
    parse_Scotia_Pdf(filePathScotia)


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
                    # print(item)
                    if "Statement Period" in item and pageNo == 0 and statementPageLines[lineCounter-8] == "SCENE":
                        # We can get statement period here
                        print(item)
                        print(statementPageLines[lineCounter+1])
                        print(statementPageLines[lineCounter+3])
                        # print("\n")
                    if "Continued on page" in item:
                        # We can directly jump to page number
                        # print(item[-1])
                        pass
                    if "AMOUNT($)" in item:
                        statementNum = statementPageLines[lineCounter + 1]
                        # Found
                        # print("Transaction Number:", statementNum)
                        statementLineCounter = 0
                        for i in range(lineCounter + 1, len(statementPageLines)):
                            if statementPageLines[i].isnumeric():
                                if int(statementPageLines[lineCounter + 1])+ (statementLineCounter) == int(statementPageLines[i]):
                                    # print("\n")
                                    # print(statementLineCounter)
                                    transactionElements = []
                                    for j in range(0, 9):
                                        if statementPageLines[i+j].isnumeric():
                                            if int(statementPageLines[lineCounter + 1])+statementLineCounter + 1 == int(statementPageLines[i+j]):
                                                # print(int(statementPageLines[lineCounter+12])+statementLineCounter +1)
                                                break
                                        if statementPageLines[i+j] == "If you have any questions about this":
                                            break
                                        if j != 0:
                                            transactionElements.append(
                                                statementPageLines[i+j])
                                        
                                        # print(statementPageLines[i+j])
                                    # print(int(statementPageLines[i]))
                                    
                                    cost = [transactionElements[len(transactionElements) - 1]]
                                     # print cost
                                    Discription = ""
                                    

                                    for x in range(2 , (len(transactionElements)-1)):
                                        Discription =  Discription + " " + transactionElements[x]
                                                                       
                                    transactionElements = (transactionElements[0:2]) + [Discription] + cost

                                    transactionDict.update(
                                        {int(statementPageLines[i]): transactionElements})
                                    # print(transactionDict)
                                    statementLineCounter += 1

                        # print(int(statementNum)+1)

                    lineCounter += 1
                else:
                    # print(item)
                    if "Transactions - continued" in item:
                        # print(statementNum)
                        statementLineCounter = 0
                        for i in range(lineCounter+8, len(statementPageLines)):
                            if statementPageLines[i].isnumeric():
                                if int(statementPageLines[lineCounter+8])+statementLineCounter == int(statementPageLines[i]):
                                    # print("\n")
                                    # print(statementPageLines[i])
                                    transactionElements = []
                                    for j in range(0, 9):
                                        if statementPageLines[i+j].isnumeric():
                                            if int(statementPageLines[lineCounter+8])+statementLineCounter + 1 == int(statementPageLines[i+j]):
                                                # print(int(statementPageLines[lineCounter+12])+statementLineCounter +1)
                                                break
                                        if statementPageLines[i+j] == "SUB-TOTAL CREDITS":
                                            break
                                        if j != 0:
                                            transactionElements.append(
                                                statementPageLines[i+j])

                                    cost = [transactionElements[len(transactionElements) - 1]]
                                     # print cost
                                
                                    transactionElements = (transactionElements[0:3]) + cost
                                        # print(statementPageLines[i+j])
                                    # print(int(statementPageLines[i]))
                                    transactionDict.update(
                                        {int(statementPageLines[i]): transactionElements})
                                    # print(transactionDict)
                                    statementLineCounter += 1
                    lineCounter += 1

    print("\n")
    pdfFileObj.close()
    print(transactionDict)

    #converting the dictionary file to the excel file 
    workbook = xlsxwriter.Workbook('transactions.xlsx')
    worksheet = workbook.add_worksheet()
    
    worksheet.write(0 , 0 , "Reference Number")
    worksheet.write(0 , 1 , "Transaction Date")
    worksheet.write(0 , 2 , "Transaction Post")
    worksheet.write(0 , 3 , "Details")
    worksheet.write(0 , 4 , "Amount(CAD)")
    
    row = 1
    for key in transactionDict.keys():
        worksheet.write(row, 0, key)
        worksheet.write_row(row, 1, transactionDict[key])
        row += 1

    workbook.close()


def parse_Neo_Pdf(filePath):
    pdfFileObj = open(filePath, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    for pageNo in range(0, pdfReader.numPages):
        if pageNo != 1:
            pageObj = pdfReader.getPage(pageNo)

    pdfFileObj.close()


if __name__ == '__main__':

    main_function()
