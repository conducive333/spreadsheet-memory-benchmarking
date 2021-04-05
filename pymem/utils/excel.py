import win32.win32process as win32process
import win32com.client

class Excel:

    def __init__(self, start=False):
        self.wb = None
        if start:
            self.start()
        else:
            self.excel, self.pid = None, None

    def start(self):
        # A full list of Excel's properties may be found here: 
        # https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)
        self.excel = win32com.client.DispatchEx('Excel.Application')
        _,self.pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
        self.excel.ScreenUpdating   = False
        self.excel.DisplayAlerts    = True
        self.excel.Visible          = False

    def open_wb(self, path):
        if self.wb is None:
            self.wb = self.excel.Workbooks.Open(path)

    def close_wb(self):
        if self.wb is not None:
            self.wb.Close()
            self.wb = None

    def end(self):
        self.excel.Application.Quit()