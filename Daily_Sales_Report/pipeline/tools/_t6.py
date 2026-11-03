# -*- coding: utf-8 -*-
import os, sys, time
import openpyxl
from openpyxl import load_workbook

def rss_mb():
    try:
        import ctypes
        from ctypes import wintypes
        class PMC(ctypes.Structure):
            _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD),
                        ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t),
                        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t), ("QuotaPagedPoolUsage", ctypes.c_size_t),
                        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t), ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                        ("PagefileUsage", ctypes.c_size_t), ("PeakPagefileUsage", ctypes.c_size_t)]
        pmc = PMC(); pmc.cb = ctypes.sizeof(PMC)
        ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.windll.kernel32.GetCurrentProcess(), ctypes.byref(pmc), pmc.cb)
        return pmc.WorkingSetSize / 1048576
    except Exception as e:
        return -1

p = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/1.每日销售数据表/欧盟英国库存销售表.xlsx"
t = time.time()
print("RSS start %.0f MB" % rss_mb(), flush=True)
wb = load_workbook(p)
print("RSS after load %.0f MB, t=%.1fs" % (rss_mb(), time.time() - t), flush=True)
ws = wb["英国-每日库存分析表"]
print("access ok", ws.max_row, ws.max_column, flush=True)
wb.save(r"C:/Users/dell/Desktop/work-main/sale_file_gen/_archive/_t6_out.xlsx")
print("saved t=%.1fs RSS %.0f MB" % (time.time() - t, rss_mb()), flush=True)
