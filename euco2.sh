#!/bin/sh
#Examples to collect CRFReporter Reporting tables (Excel files)
#Collect TableA-TableF for years 1990-20xx for each all countries
python euco2.py euco2.py -d /data/d4/projects/khk/CRF_vertailu/ --all  -s 1990 -e 2019 
#Using CRFReporter Reporting tables (Excel files)
#Collect Table4.Gs1 HWP for years 1990-20xx for each all countries
python euco2hwp.py euco2.py -d /data/d4/projects/khk/CRF_vertailu/ --all  -s 1990 -e 2019 
