import os
import subprocess
import time
def printing(path):
    printer = subprocess.Popen(["powershell", "Start-Process", "-FilePath", path, "-verb", "print"])
    time.sleep(5)
    #os.startfile(path,"print")

if  os.path.isdir("PrintData"):
    os.chdir("PrintData")
    files = os.listdir(os.getcwd())
    rootPath = os.getcwd()
    for file in files:
        path = os.path.join(rootPath, file)
        printing(path)
