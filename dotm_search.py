#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "cdanjha"

import zipfile
import os
# import argparse
# parser = argparse.ArgumentParser()
# parser.parse_args()


def main(directory, text_to_search):
    file_list = os.listdir(directory)
    files_searched = 0
    files_matched = 0
    print("searching directory {} for text '{}' ...".format(directory, text_to_search))
    for file in file_list:
        files_searched += 1
        full_path = os.path.join(directory, file)
        if not full_path.endswith(".dotm"):
            print("This is not a dotm {}".format(full_path))
            continue
        if not zipfile.is_zipfile(full_path):
            print("This is not a zip file {}".format(full_path))
            continue
        with zipfile.ZipFile(full_path, "r") as zipped:
            toc = zipped.namelist()
            if "word/document.xml" in toc:
                with zipped.open("word/document.xml", "r") as doc:
                    for line in doc:
                        i = line.find(text_to_search)
                        if i >= 0:
                            files_matched += 1
                            print(" ...{}...".format(line[i - 40:i + 40]))  
    print("Files matched: {}".format(files_matched))
    print("Files searched: {}".format(files_searched))                         

                            

# def a_library():
#     if len(sys.argv) == 4:
#         return text_to_search(sys.argv[1], sys.argv[3])
#     elif len(sys.argv) == 2:
#         return text_to_search(sys.argv[1], './')
#     else: sys.exit(0)



if __name__ == '__main__':
    main("dotm_files", "$")
