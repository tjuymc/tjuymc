import sys
sys.path.append(r'C:\my job\ICBC\ZOS\performance\mylib')
import cf
import cpu
import csusage
import vstor
import wlm


cf.cf()
cpu.cpu()
csusage.csusage()
vstor.vstor()
wlm.wlm()