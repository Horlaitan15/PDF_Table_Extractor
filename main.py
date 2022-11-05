from colorama import Fore
import xlsxwriter
import camelot as cm
import csv
import os
from colorama import Fore
import pyfiglet



font = pyfiglet.figlet_format('PDF Table Extractor.', font = 'doh', width=200)
print(font)
class PDFTable:
    def __init__(self):
        self.table_data = []
        self.input_pdf = cm.read_pdf("LIST-OF-SUCCESSFUL-CANDIDATES-FOR-UPLOADING-1.pdf", flavor="lattice", pages='all', line_scale=50) # stream or lattice
        self.tables_num = len(list(self.input_pdf))

    def extract_pdf(self):
        print(Fore.YELLOW + "Extracting Table from PDF...\n")
        self.input_pdf.export("new.csv", f='csv', compress=False)

        for i in range(len(list(self.input_pdf))):
            print(self.input_pdf[i].df)


    def merge_csv(self):
        print(Fore.BLUE + "Creating a CSV File from all generated tables...")
        table = 1
        for i in range(1, self.tables_num + 2):
            try:
                with open(f'new-page-{i}-table-{table}.csv', 'r') as f:
                    reader = csv.reader(f)
                    self.table_data.extend(list(reader))
            except FileNotFoundError:
                i -= 1
                table += 1
                
            
        with open("successes.csv", 'w') as w:
            writer = csv.writer(w)
            for data in self.table_data:
                writer.writerow(data)

    def write_excel(self):
        print(Fore.BLUE + "Creating an Excel File from all generated tables...")
        workbook = xlsxwriter.Workbook('Successes.xlsx')
        worksheet = workbook.add_worksheet()

        for col, col_data in enumerate(list(zip(self.table_data))):
            for row, row_data in enumerate(list(col_data[0])):
                worksheet.write(col, row, row_data)
        workbook.close()


    def delete(self):
        table = 1
        for i in range(1, self.tables_num + 1):
            try:
                os.remove(f'new-page-{i}-table-{table}.csv', dir_fd=None)
            except:
                self.tables_num -= 1
                table += 1
            continue



if __name__ == "__main__":
    run = PDFTable()
    run.extract_pdf()
    run.merge_csv()
    run.write_excel()
    run.delete()
    print(Fore.GREEN + "JOB DONE!".center(150))
    print(pyfiglet.figlet_format("Thanks,\nAjani", font='doh', width=400))