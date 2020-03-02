#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.

Assistance: Enrique Galindo, Jake Hershey, Derek Barnes
"""
__author__ = "Patrick Buzzo"

import os
import sys
import zipfile
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--dir', dest='direct', type=str, default='dotm_files/',
                    help='path to folder')
# parser.add_argument(
#     '--dir', '--output', help='Output file name', default='')
requiredNamed = parser.add_argument_group('required named arguments')
requiredNamed.add_argument(
    'input', help='Input file name')
args = parser.parse_args()
direct = args.direct
search_term = args.input


def main():
    matches = 0
    search_count = 0
    if not direct:
        for i in os.listdir('dotm_files'):
            full_filename = 'dotm_files/' + i
            # print(full_filename)
            if zipfile.is_zipfile(full_filename):
                z = zipfile.ZipFile(full_filename)

                for name in z.namelist():
                    match = False
                    search_count += 1
                    if "word/document.xml" in name:
                        with z.open("word/document.xml", "r") as dotm:
                            for line in dotm:
                                line = line.decode('utf-8')
                                if search_term in line:
                                    match = True
                                    start = line.index(search_term)
                                    print(full_filename + '/' + name + ': ' +
                                          line[start - 20:start + 20] + '\n')
                                    # print(line.index(search_term))
                    if match == True:
                        matches += 1
            else:
                continue
    elif direct:
        for i in os.listdir(direct):
            full_filename = direct + '/' + i
            # print(full_filename)
            if zipfile.is_zipfile(full_filename):
                z = zipfile.ZipFile(full_filename)

                for name in z.namelist():
                    match = False
                    search_count += 1
                    if "word/document.xml" in name:
                        with z.open("word/document.xml", "r") as dotm:
                            for line in dotm:
                                line = line.decode('utf-8')
                                if search_term in line:
                                    match = True
                                    start = line.index(search_term)
                                    print('Match found in file ' + full_filename +
                                          '\n' + line[start - 20:start + 20] + '\t' + '\n')
                                    # print(line.index(search_term))
                    if match == True:
                        matches += 1
            else:
                continue
        else:
            print(direct + ' --> Path does not exist!')

    print('Total Files Searched: ' + str(search_count))
    print('Total Matched Files: ' + str(matches))

    print(direct)
    print(search_term)


if __name__ == '__main__':
    main()
