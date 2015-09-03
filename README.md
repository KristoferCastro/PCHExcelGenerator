# PCHExcelGenerator
Randomly generate a dataset with a parient-child hiearchy between employees and manager

## How to run

install XlsxWriter: http://xlsxwriter.readthedocs.org/en/latest/getting_started.html

1. cd to directory
2. execute command "python PCHExcelGenerator.py"
3. input parameters that is asked


## Warning
This program uses a database of 5000 first names and 80000 last names which it uses to generate the id.  This program does not guarantee the ID to be unique.  Need to be reworked to minimize this maybe using a random id generator as the id.

## Note
It may take a while for larger input numbers.  This program is not optimized yet to handle generating large excel files.  Need to implement multi-threading and break down into smaller task to make better memory usage.
