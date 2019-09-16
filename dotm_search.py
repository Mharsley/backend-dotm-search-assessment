#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "@Mharsley2"


import zipfile
import argparse
import sys
import os

def get_word_xml(docm_filename, searchtext):
    """ open zipfile and searches for substring"""
    if not zipfile.is_zipfile(docm_filename):
        print("error: not a zipfile")
        return 0
    with zipfile.ZipFile(docm_filename) as zip:
        # print(zip.namelist())
        with zip.open('word/document.xml') as doc:
            for line in doc:
                line = str(line)
                index = line.find(searchtext)
                if index >= 0:
                    print("Match found in file {}".format(docm_filename))
                    print("   ..." + line[index-40:index+40]+"...")
                    return True
            return False

def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", help="dotm folder to search")
    parser.add_argument("text", help="text to search for")
    return parser
# zipfile.is_zipfile(filename)
# ZipFile.namelist()
# ZipFile.open(name[, mode[, pwd]])


def main():
    """ searches a folder for dotm files"""
    parser = create_parser()
    namespace = parser.parse_args()
    searched = 0
    matched = 0
    print("Searching directory " + namespace.folder + " for " + namespace.text)
    for f in os.listdir(namespace.folder):
        if not f.endswith(".dotm"):
            print("skip file {}".format(f))
            continue
        searched += 1
        full_path = os.path.join(namespace.folder, f)
        if get_word_xml(full_path, namespace.text):
            matched += 1
    print("files searched = {}".format(searched))
    print("files matched = {}".format(matched))



if __name__ == '__main__':
