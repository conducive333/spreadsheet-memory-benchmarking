import win32process
import win32api
import win32con
import psutil

class ExcelMemDataCollector:

    PEAK_WSS = "Peak WSS"
    WSS = "WSS"
    RSS = "RSS"
    USS = "USS"

    def __init__(self, excel):
        self.excel          = excel
        _,self.pid          = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
        self.handl          = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, self.pid)
        self.count          = 0
        self.tot_peak_wss   = 0
        self.tot_wss        = 0
        self.tot_rss        = 0
        self.tot_uss        = 0
        self.max_peak_wss   = float('-inf')
        self.max_wss        = float('-inf')
        self.max_rss        = float('-inf')
        self.max_uss        = float('-inf')
        self.min_peak_wss   = float('inf')
        self.min_wss        = float('inf')
        self.min_rss        = float('inf')
        self.min_uss        = float('inf')

    def measure(self, normalizer=1):
        
        # See: https://docs.microsoft.com/en-us/windows/win32/api/psapi/ns-psapi-process_memory_counters
        win32c_meminfo = win32process.GetProcessMemoryInfo(self.handl)

        # See: https://psutil.readthedocs.io/en/latest/#psutil.Process.memory_full_info
        psutil_meminfo = psutil.Process(self.pid).memory_full_info()

        # Collect metrics of interest (default unit is bytes)
        peak_wss = win32c_meminfo['PeakWorkingSetSize'] / normalizer
        wss = win32c_meminfo['WorkingSetSize'] / normalizer
        rss = psutil_meminfo.rss / normalizer
        uss = psutil_meminfo.uss / normalizer

        # Update metrics of interest
        self.tot_peak_wss   += peak_wss
        self.tot_wss        += wss
        self.tot_rss        += rss
        self.tot_uss        += uss
        self.max_peak_wss   = max(self.max_peak_wss, peak_wss)
        self.max_wss        = max(self.max_wss, wss)
        self.max_rss        = max(self.max_rss, rss)
        self.max_uss        = max(self.max_uss, uss)
        self.min_peak_wss   = min(self.min_peak_wss, peak_wss)
        self.min_wss        = min(self.min_wss, wss)
        self.min_rss        = min(self.min_rss, rss)
        self.min_uss        = min(self.min_uss, uss)

        # Clean up
        self.count += 1

    def report(self, smooth=True, prefix="", suffix=""):
        if smooth and self.count >= 3:
            return {
                prefix + ExcelMemDataCollector.PEAK_WSS + suffix: (self.tot_peak_wss - self.min_peak_wss - self.max_peak_wss) / (self.count - 2),
                prefix + ExcelMemDataCollector.WSS + suffix: (self.tot_wss - self.min_wss - self.max_wss) / (self.count - 2),
                prefix + ExcelMemDataCollector.RSS + suffix: (self.tot_rss - self.min_rss - self.max_rss) / (self.count - 2),
                prefix + ExcelMemDataCollector.USS + suffix: (self.tot_uss - self.min_uss - self.max_uss) / (self.count - 2)
            }
        else:
            if self.count != 0:
                return {
                    prefix + ExcelMemDataCollector.PEAK_WSS + suffix: self.tot_peak_wss / self.count,
                    prefix + ExcelMemDataCollector.WSS + suffix: self.tot_wss / self.count,
                    prefix + ExcelMemDataCollector.RSS + suffix: self.tot_rss / self.count,
                    prefix + ExcelMemDataCollector.USS + suffix: self.tot_uss / self.count
                }
            else:
                return {
                    prefix + ExcelMemDataCollector.PEAK_WSS + suffix: None,
                    prefix + ExcelMemDataCollector.WSS + suffix: None,
                    prefix + ExcelMemDataCollector.RSS + suffix: None,
                    prefix + ExcelMemDataCollector.USS + suffix: None
                }
    
    def close_handle(self):
        self.handl.close()