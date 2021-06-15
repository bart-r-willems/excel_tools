import win32com.client
from pathlib import Path


def main():
    '''Testig the module functions'''
    #test_file = Path(r'r:\test_data\custperf_ml.xlsb')
    #xl = ExcelFile(test_file)
    #xl.ExportPivot('ffe-customer')
    for direction in ['Export', 'Export']:
        filename = Path(f'R:\\test_data\\3Yr Trend {direction}.xlsb')
        xl = ExcelFile(filename)
        xl.ExportPivot('PIVOT')


class ExcelFile:
    # excel constants
    XLORIENTATIONHIDDEN = 0
    XLCSVUTF8 = 62

    def __init__(self, name):
        self.name = str(name)
        self.app = win32com.client.Dispatch('Excel.Application')
        self.app.visible = 1
        self.wb = self.app.workbooks.open(self.name, 0, 1)

    def __repr__(self):
        return f"ExcelFile('{self.name}')"

    def __del__(self):
        try:
            self.wb.saved = True
            self.app.quit()
        except Exception as e:
            pass

    def ExportPivot(self, sheet_name, filename=None, index=1):
        '''
        Export the first pivot table on a worksheet to a file

        Arguments:
            sheet_name: name of the worksheet containing the pivot table
            filename: name of the output file; if ommited, same as 
                      the input filename
            index: if there are multiple pivot tables, the index of the table
        '''
        ws = self.wb.worksheets[sheet_name]
        pvt = ws.pivottables(index)
        self._remove_labels(pvt)
        ws_data = self._drill_down(pvt)
        if not filename:
            filename = Path(self.name)
            clean_sheet_name = ws.name.replace(" ", '_').replace('\\', '_')
            new_name = f'{filename.name}_{clean_sheet_name}'
            filename = filename.with_name(new_name).with_suffix('.csv')
        self.ExportWorksheet(ws_data.name, filename)

    def ExportWorksheet(self, sheet_name, filename):
        path = Path(filename)
        ws = self.wb.worksheets[sheet_name]
        if path.exists():
            path.unlink()
        if path.suffix.lower() == '.csv':
            ws.saveas(str(path), self.XLCSVUTF8)

    def _remove_labels(self, pvt):
        '''Remove row and column labels from pivot table'''
        row_labels = [fld.name for fld in pvt.rowfields]
        col_labels = [fld.name for fld in pvt.columnfields]
        for name in [name
                     for name in row_labels + col_labels
                     if name != 'Values']:
            pvt.pivotfields(name).orientation = self.XLORIENTATIONHIDDEN

    def _drill_down(self, pvt):
        '''
        Returns the drill down worksheet of a pivot table

        Arguments:
            pvt: the pivot table to be drilled down on
        '''
        before_names = {ws.name for ws in self.wb.worksheets}
        tbl_rng = pvt.tablerange1
        # the showdetail property forces a drill down on a new worksheet
        tbl_rng.cells(tbl_rng.rows.count, tbl_rng.columns.count).showdetail = 1
        # excel doesn't tell you what the new worksheet is though
        after_names = {ws.name for ws in self.wb.worksheets}
        data_name = (after_names - before_names).pop()
        return self.wb.worksheets[data_name]
        





    

if __name__ == '__main__':
    main()