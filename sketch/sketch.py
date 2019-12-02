#!usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
@Author: Joshua
@File: sketch.py
@Time: 2019/12/2
"""

import cv2
import sys

def main(img_source, img_destination):
    # read source image
    rgb_img = cv2.imread(img_source)
    # rgb to gray
    gray_img = cv2.cvtColor(rgb_img, cv2.COLOR_RGB2GRAY)
    # blur the gray
    gray_img_blurred = cv2.medianBlur(gray_img, 5)
    # Gaussian blur the gray
    img_Gaussian_blurred = cv2.GaussianBlur(gray_img_blurred, ksize=(21, 21), sigmaX=0, sigmaY=0)
    # merge the gray and the blurred gray
    img_edge = cv2.divide(gray_img, img_Gaussian_blurred, scale=255)
    # write the image to the img_destination
    cv2.imwrite(img_destination, img_edge)

if __name__ == '__main__':
    img_source = sys.argv[1]
    img_destination = sys.argv[2]
    main(img_source, img_destination)