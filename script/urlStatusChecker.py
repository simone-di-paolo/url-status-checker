import xlrd  # pip install xlrd
import xlwt  # pip install xlwt
import requests  # pip install requests
import tkinter as tk


# Funzione per aggiungere un messaggio di errore al log_text
def add_error_message(log_text, error_message):
    log_text.insert(tk.END, f"ERROR: {error_message}\n")
    log_text.see(tk.END)


index = 0


def execute(start_from_row, sheet_number, do_only_one_sheet, file_to_read, excel_path_to_write, log_text):
    global index
    try:
        excel_file_path = xlrd.open_workbook(file_to_read)
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko)'}

        # excel to write
        exel_file_to_write = xlwt.Workbook()

        # Aggiungi l'output di log al widget di testo
        log_text.delete("1.0", tk.END)
        log_text.configure(fg="white")
        log_text.insert(tk.END, "### Inizio script ###\n")

        for worksheets in excel_file_path.sheet_names():
            excel_sheet = excel_file_path.sheet_by_index(sheet_number)
            sheetToWrite = exel_file_to_write.add_sheet(excel_file_path.sheet_names().__getitem__(index))

            for row in range(start_from_row, excel_sheet.nrows):

                # row/column of the old URL
                urlFrom = excel_sheet.cell_value(row, 0)
                urlFrom = urlFrom.strip()

                # re-write the start url in the first cell
                sheetToWrite.write(row, 0, urlFrom)

                # here I get the history of the old file with every redirect (if there are any)
                # verify=False so you don't have to worry about the SSL: CERTIFICATE_VERITY_FAILED error
                fromUrlResponse = requests.get(urlFrom, headers=headers, verify=False)
                # I'm disabling the warnings
                requests.packages.urllib3.disable_warnings()
                status = fromUrlResponse.status_code  # if needed, 200 or 404 or something else
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
                else:
                    redirectStatus = "ERROR: Something strange is happened"

                sheetToWrite.write(row, 3, redirectStatus)

                # Aggiorna il log con lo stato di completamento
                completion_percentage = int((row - start_from_row + 1) / (excel_sheet.nrows - start_from_row) * 100)
                if completion_percentage % 10 == 0:
                    log_text.insert(tk.END, f"\nPercentuale di completamento: {completion_percentage}%")

            # if you want to do only one sheet at a time interrupts the "for" here
            if do_only_one_sheet:
                break

            index = index + 1

        exel_file_to_write.save(excel_path_to_write)

        # Aggiorna il log con l'indicazione di completamento
        log_text.insert(tk.END, "\n### Completato ###\n")
    except FileNotFoundError:
        add_error_message(log_text, "File not found. \nPlease select a valid file.")
    except requests.exceptions.MissingSchema:
        add_error_message(log_text, "Invalid URL in row number " + str(index + 1)
                          + ". \nPlease, check the source excel file.")
