import xlrd  # pip install xlrd
import xlwt  # pip install xlwt
import requests  # pip install requests

startFromWhichRow = 1   # in case the first row has no rules but only column titles put 1 otherwise 0
sheetNumber = 0  # the index of the excel sheet you want to process, if doOnlyOneSheet is False it will do them all starting from this index
doOnlyOneSheet = False  # if for any reason you need to do one file at a time put this to True (capital T)

pathFileToRead = 'W:\\Example\\Local\\Folder\\myExcelFileWithRules.xlsx'    # path of the file to read
pathExcelToWrite = 'W:\\Example\\Local\\Folder\\result.xlsx'  # path of the file created with results
excel_file_path = xlrd.open_workbook(pathFileToRead)
excel_sheet = excel_file_path.sheet_by_index(sheetNumber)

#pathFileToWrite = open("W:\\Example\\Local\\Folder\\rewrited-rules.txt", "w")  # path of the file to write to

headers = headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko)'}

#excel to write
exel_file_to_write = xlwt.Workbook()

index = 0

for worksheets in excel_file_path.sheet_names():
    excel_sheet = excel_file_path.sheet_by_index(sheetNumber)
    sheetToWrite = exel_file_to_write.add_sheet(excel_file_path.sheet_names().__getitem__(index))

    for row in range(startFromWhichRow, excel_sheet.nrows):

        # row/column of the old URL
        urlFrom = excel_sheet.cell_value(row, 0)
        urlFrom = urlFrom.strip()

        # re-write the start url in the first cell
        sheetToWrite.write(row, 0, urlFrom)

        # here I get the history of the old file with every redirect (if there are any)
        fromUrlResponse = requests.get(urlFrom, headers=headers)
        status = fromUrlResponse.status_code #if needed, 200 or 404 or something else
        fromUrl_history = fromUrlResponse.history
        fromNew_url = fromUrlResponse.url

        if fromUrl_history.__len__() > 0 & fromUrl_history.__len__() < 2:
            # if history size > 0
            redirectHistory = "Redirect status: " + str(fromUrlResponse.history.__getitem__(0).status_code)
            sheetToWrite.write(row, 2, redirectHistory)
        else:
            redirectHistory = "No redirects"
            sheetToWrite.write(row, 2, redirectHistory)

        # row/column of the new URL
        urlTo = excel_sheet.cell_value(row, 1)
        urlTo = urlTo.strip()

        # re-write the end url in the second cell
        sheetToWrite.write(row, 1, urlTo)

        if urlTo == fromNew_url and status != 404:
            redirectStatus = "The redirect is correct - Status code: " + str(status)
        elif urlTo != fromNew_url:
            redirectStatus = "The redirect it\'s not correct. Check it!"
        elif status == str(404):
            redirectStatus = "The redirect of the new page is 404."

        sheetToWrite.write(row, 3, redirectStatus)

    # if you want to do only one sheet at a time interrupts the "for" here
    if doOnlyOneSheet:
        break

    index = index + 1

exel_file_to_write.save(pathExcelToWrite)