import psutil
import numpy

class MemDataCollector:

    # Raise an error if overflow occurs
    numpy.seterr(over='raise')

    TITLES = [
        'peak_nonpaged_pool',
        'peak_paged_pool',
        'peak_pagefile',
        'nonpaged_pool',
        'paged_pool',
        'peak_wset',
        'pagefile',
        'private',
        'wset',
        'rss',
        'uss',
        'vms'
    ]

    def __init__(self):
        self.counter = 0
        self.tot_mem = numpy.array([ 0                               ] * len(MemDataCollector.TITLES), dtype=numpy.ulonglong)
        self.max_mem = numpy.array([ numpy.iinfo(numpy.longlong).min ] * len(MemDataCollector.TITLES), dtype=numpy.longlong)
        self.min_mem = numpy.array([ numpy.iinfo(numpy.longlong).max ] * len(MemDataCollector.TITLES), dtype=numpy.longlong)

    def measure(self, pid):
        
        # See: https://psutil.readthedocs.io/en/latest/#psutil.Process.memory_full_info
        meminfo = psutil.Process(pid).memory_full_info()
        meminfo = numpy.array([getattr(meminfo, t) for t in MemDataCollector.TITLES])

        # Collect metrics of interest (default unit is bytes)
        # An unsigned long long should provide more than enough space here
        self.tot_mem = numpy.add(self.tot_mem, meminfo)
        self.max_mem = numpy.maximum(self.max_mem, meminfo)
        self.min_mem = numpy.minimum(self.min_mem, meminfo)

        # This is for averaging the number of measurements
        self.counter += 1

    def report(self, smooth=True, prefix="", suffix="", normalizer=1):
        if smooth and self.counter >= 3:
            aggregated = ((self.tot_mem - self.min_mem - self.max_mem) / (self.counter - 2)) / normalizer
            return { prefix + t + suffix : v for t, v in zip(MemDataCollector.TITLES, aggregated) }
        else:
            if self.counter != 0:
                aggregated = (self.tot_mem / self.counter) / normalizer
                return { prefix + t + suffix : v for t, v in zip(MemDataCollector.TITLES, aggregated) }
            else:
                return { prefix + t + suffix : None for t in MemDataCollector.TITLES }
