#!usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
@Author: Joshua
@File: excel_comparer.py
@Time: 2020/1/8
"""
import sys
sys.path.append("E:/GitHub/gadgetdaily")
from compare2excels.utils import excel_logger
from compare2excels import excel_reader
WEIGHT = {
    "node_name": 3,
    "common service": 5,
    "label": 5,
    "storage info": 5,
    "role": 10
}

if __name__ == "__main__":
    if len(sys.argv) != 2:
        excel_logger.error("Wrong input.")
        raise Exception("Wrong input.")
    source, target = excel_reader.read_excels(sys.argv[1])
