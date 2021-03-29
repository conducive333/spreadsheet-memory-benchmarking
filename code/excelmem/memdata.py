import win32.win32process as win32process
import win32.win32api as win32api
import win32con
import psutil
import numpy

class ExcelMemDataCollector:

    TITLES = [
        'nonpaged_pool',
        'num_page_faults',
        'paged_pool',
        'pagefile',
        'peak_nonpaged_pool',
        'peak_paged_pool',
        'peak_pagefile',
        'peak_wset',
        'private',
        'rss',
        'uss',
        'vms',
        'wset'
    ]

    def __init__(self, excel):
        self.count          = 0
        self.excel          = excel
        _,self.pid          = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
        self.handl          = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, self.pid)
        self.tot_mem        = numpy.array([ 0              for _ in ExcelMemDataCollector.TITLES ])
        self.max_mem        = numpy.array([ float('-inf')  for _ in ExcelMemDataCollector.TITLES ])
        self.min_mem        = numpy.array([ float('inf')   for _ in ExcelMemDataCollector.TITLES ])

    def measure(self):
        
        # See: https://psutil.readthedocs.io/en/latest/#psutil.Process.memory_full_info
        meminfo = psutil.Process(self.pid).memory_full_info()
        meminfo = numpy.array([getattr(meminfo, t) for t in ExcelMemDataCollector.TITLES])

        # Collect metrics of interest (default unit is bytes)
        self.tot_mem = numpy.add(self.tot_mem, meminfo)
        self.max_mem = numpy.max(self.max_mem, meminfo)
        self.min_mem = numpy.min(self.min_mem, meminfo)

        # Clean up
        self.count += 1

    def report(self, normalizer=1, smooth=True, prefix="", suffix=""):
        if smooth and self.count >= 3:
            aggregated = ((self.tot_mem - self.min_mem - self.max_mem) / (self.count - 2)) / normalizer
            return { prefix + t + suffix : v for t, v in zip(ExcelMemDataCollector.TITLES, aggregated) }
        else:
            if self.count != 0:
                aggregated = (self.tot_mem / self.count) / normalizer
                return { prefix + t + suffix : v for t, v in zip(ExcelMemDataCollector.TITLES, self.tot_mem) }
            else:
                return { prefix + t + suffix : None for t, v in zip(ExcelMemDataCollector.TITLES, self.tot_mem) }
    
    def close_handle(self):
        self.handl.close()