#!/usr/bin/env python3
#coding=utf-8

'''

Reads the directory and outputs sub-dir or full-dir within 
Output into .txt file in a clean style
Determines Root Folder Size if needed

Dependencies:
pip install pywin

Run:
Copy Script in Folder and run from Terminal within Folder (else dependency pywin will fail)

'''

import os
import sys
sys.path.append("C:\\_path_to_virtual_environment\\Lib\\site-packages\\")
import win32com.client as com

absolute_path = os.path.dirname(__file__)

#Re-Define Walk to select between full depth search or only first depth
def walk(top, maxdepth):
    dirs, nondirs = [], []
    for name in os.listdir(top):
        (dirs if os.path.isdir(os.path.join(top, name)) else nondirs).append(name)
    yield top, dirs, nondirs
    if maxdepth > 1:
        for name in dirs:
            for x in walk(os.path.join(top, name), maxdepth-1):
                yield x

#Full Depth or only first level
depth = input("Full Directory Depth: (Y/N) ").upper()

if depth == "Y":
    with open("output.txt", "w", newline='', encoding="utf-8") as a:
        for path, subdirs, files in os.walk(".", 2):
            a.write('.\\' + os.path.relpath(path + os.linesep, start=absolute_path)) 
            for filename in files:
                a.write('\t%s\n' % filename)
else:
    with open("output.txt", "w", newline='', encoding="utf-8") as a:
        for x in walk(".", 1):
            for dir in x:
                for item in dir:
                    a.write('.\\' + os.path.relpath(item + os.linesep, start=absolute_path))

folder_size = input("Output System Folder total Size: (Y/N) ").upper()

#You want to show the size of the root directory in GB
if folder_size == "Y":  
    
    fso = com.Dispatch("Scripting.FileSystemObject")
    folder = fso.GetFolder(".")
    GB = 1024 * 1024.0 * 1024
    print("%.2f GB" % (folder.Size / GB))

print("Done")