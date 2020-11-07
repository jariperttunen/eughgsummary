# eughgsummary
Read CRFReporter files and crete summary for EU countries.
The main file is euco2.py
Command line is (for example):
python euco2.py -d EU-MS/2017 -s 1990 -e 2015 -a 

As usual python euco2.py -h gives the help.

EU-MS/2017 is the directory is the for the downloaded excel files

To setup virtual python3 virtual environment use the requirements.txt file.
The proxy option tells the Luke proxy if needed.

python3 -m venv eughg

source eughg/bin/activate

pip --proxy <proxy server> install -r requirements.txt

