# Create GHG summaries from CRF Reporting tables.

To alleviate examination and comparison of GHG net emissions in reporting countries
eughgsummary contains three scripts to collect and organise data from annual CRF Reporting tables
(i.e. "official excel files") as country by country and year by year excel files and sheets. 

Currently it is possible to gather data concerning LULUCF land use categories,
harvested wood products and LULUCF summary. The reporting categories or subcategories used are the ones 
that are common to all parties. The collected data is provided as is, no further analysis or examination is done.

## euco2.py
The scripts collects net emissions from the Table4.A to Table4.F. This amounts to data for LULUCF
land use, land-use change and forestry. The output is a single file of all net emissions collected
and one file of net emissions collected for each of the Table4 A-F.

The command line arguments are:
+ -d: The main directory for the CRF Reporting tables. It is assumed the excel files are
      organised in this directory by countries denoted with three letter acronyms.
+ -s: Start year of the inventory.
+ -e: End year of the inventory.

To select countries one of the following must be used:
+ --eu: EU countries.
+ --euplus: EU plus GBR, ISL and NOR
+ -a: All reporting countries.
+ -c: List of country acronyms separated by spaces.

## euco2hwp.py
The script collects net emissions Harvested wood products (HWP) data for Total HWP, Total HWP Domestic
and Total HWP Exported from the Table4.Gs1. Note a country may have reported Total HWP only.
Whether Apporoach A or B is used is not explicitely mentioned. The output is a single
file with three sheets for HWP net emissions.

The command line arguments are:
+ -d: The main directory for the CRF Reporting tables. It is assumed the excel files are
      organised in this directory by countries denoted with three letter acronyms.
+ -s: Start year of the inventory.
+ -e: End year of the inventory.

To select countries one of the following must be used:
+ --eu: EU countries.
+ --euplus: EU plus GBR, ISL and NOR
+ --all: All reporting countries.
+ --countries: List of country acronyms separated by spaces.
+ --list: Use the countries in source directory pointed to with the option -d
  
Some countries report years in 1980's. These are filtered out.

## eurestoration.py
Collect row 10 from CRFReporter Excel files Table4 A,B,C and D.
The command line arguments are:
+ -d: The main directory for the CRF Reporting tables. It is assumed the excel files are
      organised in this directory by countries denoted with three letter acronyms.
+ -s: Start year of the inventory.
+ -e: End year of the inventory.

To select countries one of the following must be used:
+ --eu: EU countries.
+ --euplus: EU plus GBR, ISL and NOR
+ --all: All reporting countries.
+ --countries: List of country acronyms separated by spaces.
+ --list: Use the countries in source directory pointed to with the option -d

## eululucftotal.py
The script collects sums of CO<sub>2</sub>, CH<sub>4</sub> and N<sub>2</sub>O net emissions from Table 4 for 
Total LULUCF and categories Forest land, Cropland, Grassland, Wetlands, Settlements, Other land, HWP and
Other (rows A, B, C, D, E and F).  The unit is CO<sub>2</sub>eq, i.e. CH<sub>4</sub> and N<sub>2</sub>O
net emissions are changed to CO<sub>2</sub>eq  with their Global warming potentials (GWP).
The output is a single file containing sheets for items collected from Table 4.


The command line arguments are:
+ -d: The main directory for the CRF Reporting tables. It is assumed the excel files are
      organised in this directory by countries denoted with three letter acronyms.
+ -s: Start year of the inventory.
+ -e: End year of the inventory.

To select countries one of the following must be used:
+ --eu: EU countries.
+ --euplus: EU plus GBR, ISL and NOR
+ --all: All reporting countries.
+ --countries: List of country acronyms separated by spaces.
+ --list: Use the countries in source directory pointed to with the option -d
  
To define GWP to be used:
+ --GWP: possible values are AR4 (default, used in GHG) or AR5.

Some countries report years in 1980's. These are filtered out.

## Examples

CRFReporter excel files are in *CRF_vertailu* and collect data for EU plus GBR, ISL and NOR.

      python euco2.py -d /data/d4/projects/khk/CRF_vertailu/ -s 1990 -e 2019 --euplus 
      python euco2hwp.py -d /data/d4/projects/khk/CRF_vertailu/ -s 1990 -e 2019 --euplus 
      python eululucftotal.py -d /data/d4/projects/khk/CRF_vertailu/ -s 1990 -e 2019 --euplus

The `-h` option for python scripts gives help for command line arguments.

## Slurm 
To use Slurm workload manager see `eughgsummary.slurm` and èurestoration.slurm for details for each of the python scripts.

## Python virtual environment

The `requirements.txt`  contains information for pip to install python packages
required by eughgsummary. First, create python virtual environment, activate it and
install the packages:
+ pip  install -r requirements.txt

You may need to tell the proxy server for `pip`.
