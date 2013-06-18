Image to Excel
==============

Image to Excel is a <del>stupid</del> simple Java program for creating a pixel art Excel spreadsheet
from an image file.

Usage:

    java -jar image-to-excel.jar <image> <offsetX> <offsetY> <step>
    
Pixels are sampled from a grid with origin (offsetX, offsetY) and a spacing of `step`.
The output file will be `<image>.xlsx` with the image rendered at A1 of the first sheet.
