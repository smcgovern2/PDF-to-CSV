import sys
import subprocess as sp  
# implement pip as a subprocess:
sp.check_call([sys.executable, '-m', 'pip', 'install', 'tabula-py'])
sp.check_call([sys.executable, '-m', 'pip', 'install', 'pandas'])
sp.check_call([sys.executable, '-m', 'pip', 'install', 'numpy'])
sp.check_call([sys.executable, '-m', 'pip', 'install', 'PySimpleGui'])

#check for valid java
try:
  print(sp.check_output("java -version", stderr=sp.STDOUT, shell=True).decode('utf-8'))
except OSError:
  print("java not found on path")