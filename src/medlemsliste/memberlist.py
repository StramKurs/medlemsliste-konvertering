import argparse
import csv
import openpyxl

class convertCvs(object):
    _import_file = ''
    _eksport_file = 'medlemsliste'
    _eksport_format = 'xlsx'
    _parse_args: object

    def __init__(self):
        if __name__ == "__main__":
            self._parse_args()

    def _parse_args(self, args=None):
        parser = argparse.ArgumentParser(prog="Medlemsliste konverting", description='Konvertere cvs til xlts')
        parser.add_argument('--import-file', type=str, help='importer fil')
        parser.add_argument('--eksport-excel-2007', action='store_false', help='eksport format')
        self._parse_args = parser.parse_args(args=args)

    def set_import_file(self, file):
        self._import_file = file

    def set_eksport_file(self, file):
        self._eksport_file = file

    def set_export_formet(self, file_format):
        self._eksport_format = file_format

    def run(self):
        if self._parse_args.import_file:
            self.set_import_file(self._parse_args.import_file)

        if self._eksport_format == 'xlsx':
            wb = openpyxl.Workbook()
            ws = wb.active

            with open(self._import_file) as f:
                reader = csv.reader(f, delimiter='\t')
                for row in reader:
                    ws.append(row)

            ws.auto_filter.ref = "A1:H1"
            for cell in ws["1"]:
                cell.font = openpyxl.styles.Font(bold=True,)

            dims = {}
            for row in ws.rows:
                for cell in row:
                    if cell.value:
                        dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
            for col, value in dims.items():
                ws.column_dimensions[col].width = value

            wb.save(self._eksport_file + '.' + self._eksport_format)

if __name__ == "__main__":
    app = convertCvs()
    app.run()
