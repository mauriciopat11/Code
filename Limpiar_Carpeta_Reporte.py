import os
import glob

files1 = glob.glob('/Users/mpatinob/Dropbox/Reportes_LP/Reportes Procesados/*')
files2 = glob.glob('/Users/mpatinob/Dropbox/Reportes_LP/Reportes SAP/*')

for f in files1:
    os.remove(f)

for f in files2:
    os.remove(f)