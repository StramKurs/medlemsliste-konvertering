import os
import argparse
import csv
import openpyxl

class convert(object):
    _import_file = ''
    _export_file = ''
    _export_format = 'xlsx'
    _parse_args: object

    def __init__(self):
        if __name__ == "__main__": self._parse_args()

    def _parse_args(self, args=None):
        parser = argparse.ArgumentParser(prog="Medlemsliste konverting", description='Konvertere cvs til xlsx')
        parser.add_argument('--import-file', type=str, help='importer fil')
        parser.add_argument('--import-directory', type=str, help='importer all filer')
        parser.add_argument('--export-file', type=str, help='eksporter filnavn uden filtypenavnet')
        parser.add_argument('--eksport-excel-2007', action='store_true', help='eksport format')
        self._parse_args = parser.parse_args(args=args)

    def set_import_file(self, file):
        self._import_file = file

    def set_export_file(self, file):
        self._export_file = file

    def set_export_formet(self, file_format):
        self._eksport_format = file_format

    def convert_files_in_folder(self, directory, extension='csv'):
        for f in os.listdir(directory):

            if f.endswith('.' + extension):
                self.convert_csv_to_xlsx(directory + '/' + f)

    def convert_csv_to_xlsx(self, file):
            wb = openpyxl.Workbook()
            ws = wb.active

            with open(file) as f:
                reader = csv.reader(f, delimiter='\t')
                for row in reader:
                    ws.append(row)

            ws.auto_filter.ref = "A1:K1"

            for cell in ws["1"]:
                cell.font = openpyxl.styles.Font(bold=True,)

            dims = {}
            for row in ws.rows:
                for cell in row:
                    if cell.value:
                        dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
            for col, value in dims.items():
                ws.column_dimensions[col].width = value

            export_file = os.path.splitext(os.path.basename(file))[0] if self._export_file == '' else self._export_file

            wb.save(export_file + '.' + self._export_format)

            self.convert_csv_to_vcf(file)

    def convert_csv_to_vcf(self, file):
        export_file = os.path.splitext(os.path.basename(file))[0] if self._export_file == '' else self._export_file
        allvcf = open(export_file + '.vcf', 'w')
        with open(file) as f:
            i = 0
            count = 0
            reader = csv.reader(f, delimiter='\t')
            for row in reader:
                if i == 100:
                    count += 1
                    allvcf = open(export_file + str(count) + '.vcf', 'w')
                    i = 0

                allvcf.write( 'BEGIN:VCARD' + "\n")
                allvcf.write( 'VERSION:2.1' + "\n")
                allvcf.write( 'N:' + row[1] + ';' + row[2] + "\n")
                allvcf.write( 'FN:' + row[0] + ' ' + row[2] + ' ' + row[1] + "\n") #remember that lastname first
                allvcf.write( 'ORG:' + 'Stram Kurs' + "\n")
                allvcf.write( 'TEL;CELL:' + row[5] + "\n")
                allvcf.write( 'EMAIL:' + row[4] + "\n")
                allvcf.write( 'END:VCARD' + "\n")
                allvcf.write( "\n")
                i += 1

    def run(self):
        if self._parse_args.import_file:
            self.set_import_file(self._parse_args.import_file)

        if self._parse_args.export_file:
            self.set_import_file(self._parse_args.eksport_file)

        if self._import_file != '' and self._export_file == '':
            self._export_file = self._import_file

        if self._export_format == 'xlsx':
            if self._parse_args.import_directory:
                self.convert_files_in_folder(self._parse_args.import_directory)
            else:
                self.convert_csv_to_xlsx(self._import_file)

if __name__ == "__main__":
    app = convert()
    app.run()
