import os
import subprocess

def printing(path):
    printer = subprocess.Popen(["powershell", "Start-Process", "-FilePath", path, "-verb", "print"])
    printer.wait()
    #os.startfile(path,"print")
    #time.sleep(5)

os.chdir("PrintData")
files = os.listdir(os.getcwd())
rootPath = os.getcwd()
os.chdir("..")
for file in files:
        path = os.path.join(rootPath, file)
        printing(path)
