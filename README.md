# pandas-project

This is a Python script that takes a .xlsx file, currently hardcoded, but might be changed in the future, and the first thing it does is removing some unnecessary information in row 1 and column W. After that, the file will be saved under **'filename_cropped.xlsx'**, and this last one is the one where we do the **important stuff**, let's call it that. Here the program reads from **data_config_file.txt** a column and some parameters to search in that column, aswell as an output file to save the rows in which those parameters appear in that column.
