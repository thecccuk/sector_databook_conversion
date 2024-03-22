# NZIP Sector Databook Conversion

NZIP covers the following CCC sectors:

1. Industry
2. Fuel supply

The python scripts in this folder allow you to convert the outputs from an NZIP run into a Sector Databook.

## Description of files

- `nzip_model_sector_map.csv` is a mapping from the EE subsector name (used by NZIP), to the subsector we use at the CCC.
- `nzip.py` is a python module containing the "low level" functions which handle all the calculations and conversion logic. Assuming the structure of the NZIP model outputs is consistent, you shouldn't need to make changes here.
- `nb.pynb` is a Jupyter notebook which which contains configuration and then calls the different functions in `nzip.py`. The notebook contains some additional information about each step of the conversion process.
- `requirements.txt` is a list of python packages required to run the conversion.

## Running the conversion

To run the notebook, we will use [Google Colab](https://colab.research.google.com/), which provides python notebooks via the browser for free.

You can create a new colab based on the notebook stored in this gitub repository by following this link:

- https://colab.research.google.com/github/thecccuk/sector_databook_conversion/blob/main/nzip/nb.ipynb

