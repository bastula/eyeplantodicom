# eyeplantodicom
A Python script used to convert an Eyeplan Excel dose file to a DICOM RT Dose file.

Required files:
* Eyeplan dose file in Excel format
* Existing DICOM RT Dose file to model the output dose upon (new UIDs will be generated automatically)

The code has been tested on Python 2 and requires the following modules:

* [numpy](http://www.numpy.org) - used to process the data
* [scipy](http://www.scipy.org) - used to interpolate the dose grid
* [pandas](http://pandas.pydata.org) - used to read and process the Excel file
* [pydicom](http://www.pydicom.org) - used to read and write DICOM data

All script configuration options are documented by running the ```eyeplantodicom.py``` command with the argument ```--help```
