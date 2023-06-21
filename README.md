# NumberFormatted-Excel-Data-Extraction

In an Excel Workbook, number-formatted values may contain important information, such as the currency of the value. However, there is no direct way to obtain the formatted value. Therefore, this program aims to flatten the number formats into the value of the cell, with all date values converted to the format of DD/MM/YYYY. 

This program supports all valid number formats, and with the integration of Regular Expressions, it also supports additional date formats. The supported date formats are listed below:

Excel-supported formats:
~ 2023/06/21        ~ 2023-06-21
~ 21/06/2023        ~ 21/06/23
~ 21-Jun-23         ~ 21-June-23
~ 21-Jun-2023       ~ 21-June-2023
~ 21 June 2023      ~ 21 June, 2023
~ etc.

Extended supporting formats:
~ June 21, 2023     ~ June 21st, 2023
~ 23rd June 2021    ~ 21.06.2023
~ etc.

Unsupported formats:
~ 06/21/2023        ~ 6/21/23


To facilitate automation, the program has been modified to be compatible with Linux. 
