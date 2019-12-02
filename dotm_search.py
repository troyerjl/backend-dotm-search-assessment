#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "troyerjl"


import zipfile
import argparse
import os
import sys


# def open_function(args):
#     full_path = args
#     with ZipFile(full_path) as z:
#         with z.open('word/document.xml') as dotm_doc:
#             print(dotm_doc.read())


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", default=".", help="read .dotm files from a directory")
    parser.add_argument("text")
    return parser


def main(args):
    parser = create_parser()
    args = parser.parse_args()
    
    file_list = os.listdir(args.dir)
    match_count = 0
    search_count = 0

    for file in file_list:
        if not file.endswith(".dotm"):
            print("Disregarding file {}".format(file))
            continue
        search_count += 1
        full_path = os.path.join(args.dir, file)

        with zipfile.ZipFile(full_path) as z:
            with z.open("word/document.xml") as doc:
                for line in doc:
                    text_position = line.find(args.text)
                    if text_position >= 0:
                        match_count += 1 
                        print("Match found in file {}".format(file))
                        print("...{}...".format(line[text_position-40:text_position+40]))
    print("Files searched: {}".format(search_count))
    print("Line Matched {}".format(match_count))

if __name__ == '__main__':
    main(sys.argv[1:])
