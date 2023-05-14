# Project
Project for EGM722 - Microalgae cell concentration to carbon content
-----------

This Python script is designed to convert data on cell concentration of microalgae (cells per liter) to carbon content (picogram carbon) and generate the output in a Microsoft Excel Workbook. The script makes use of the HELCOM PEG Biovolume list (Olenina et al. 2006) and its annually updated annex  (version 2020), and the Species_Order database (via algaebase and WoRMS) to create dictionaries in which the biovolume of a singular cell for the input species is found. 
By using the provided cell concentration, the script calculates the total biovolume for each input species based on the dictionaries. With the calculated biovolume, the script then determines the amount of carbon content for each input species.


Set-up and installation
------------
A GitHub account and a conda environment is required to run the script. 
To install Git https://git-scm.com/downloads (and optionally GitHub Desktop https://desktop.github.com/) and Anaconda https://docs.anaconda.com/free/anaconda/install/, see provided links for detailed information.

The main dependencies needed for the conda environment are:

-	openpyxl (version 3.1.1)
-	os
-	pandas (version 2.0.0)

The environment file is provided in this repository (https://github.com/GPV1/Project/tree/main). The dependencies can be installed using the command ‘pip install’ or by adding the environment via Anaconda. 
In the folder project script the conversion script can be found (cell_concentration_to_carbon_content.py). To run the code, make sure to provide the correct paths for the databases and the input data. Additionally, fill in the necessary function arguments within the script (see method).

The HELCOM PEG Biovolume list can be retrieved from this manual (click on manual https://helcom.fi/wp-content/uploads/2020/01/HELCOM-Guidelines-for-monitoring-of-phytoplankton-species-composition-abundance-and-biomass.pdf and click on the link on page eleven (2.3.3.4.2 Biovolume calculation), the direct link https://www.ices.dk/data/Documents/ENV/PEG_BVOL.zip does unfortunately not work). Two files are downloaded, a Word document with information about the HELCOM PEG Biovolume list and the list itself (PEG_BVOL2022, Excel workbook). 

This repository includes an extract of the Species_Order database (Species_Order_GPV.xlsx) and test data (Data_2018_GPV.xlsx). Place all three files in your repository and provide the correct paths for the files in the beginning of the script.

To easily locate and organize the input files make sure the input data is placed in a subdirectory in your repository. 
