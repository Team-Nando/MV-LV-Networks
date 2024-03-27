# Australian MV-LV Networks

## Introduction

This repository provides four real large-scale three-phase Australian MV (22kV L-L) distribution networks and corresponding pseudo LV networks (European-style three-phase networks for urban and two-phase networks for rural) down to the connection point of single-phase customers (the pseudo LV networks have been created following modern Australian LV design principles [[1]](https://www.researchgate.net/publication/344346531_On_the_role_of_integrated_MV-LV_network_modelling_in_DER_studies)). These networks are run and operated by [AusNet Services](https://www.ausnetservices.com.au/), a distribution company in the State of Victoria. This repository also includes a large pool of anonymised real 30-min resolution residential demand (kW) profiles. Furthermore, this repository provides the code necessary to run time-series power flow simulations and plot the corresponding results, including visualisation of geographical data, voltage profiles, and asset utilisation.

The table below presents some of the main characteristics of the four Australian distribution networks. The "Name" corresponds to the code/ID used by AusNet Services. "C&I" stands for Commercial and Industrial. 

| Network | Type | Name | Nominal Voltages | Total MV Lines | Dist. Trafos | Residential Customers | C&I Customers |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 1 | Rural | _SMR8_ | 22kV, 0.4kV and SWER (12.7kV and 0.23kV) | 680 km | 704 | 3,608 | 61 |
| 2 | Rural | _KLO14_ | 22kV, 0.4kV and SWER (12.7kV and 0.23kV) | 329 km | 700 | 4,691 | 24 |
| 3 | Urban | _HPK11_ | 22kV and 0.4kV | 20 km | 44 | 5,274 | 1 |
| 4 | Urban | _CRE21_ | 22kV and 0.4kV | 30 km | 79 | 3,374 | 9 |

Key files:
- `Network_X_XXX_XXX.xls`. All the data for each of these networks are in the MS Excel files. Each Excel file has different sheets for different elements such as primary substation transformers, distribution transformers, MV lines, LV lines, etc.
- `Profiles.zip`. The pool of anonymised real 30-min resolution residential demand (kW) profiles is available as Numpy files that have been zipped. You need to extract it, then copy and paste it to the working directory where the MS Excel files are placed.
- `Main.ipynb`. This is the interactive code via Jupyter Notebook and Python to run power flow analysis for any of the networks.
- `backend.py`. It includes the code that runs time-series power flow simulations and plots the corresponding results.

### Pre-Requisites
- Ideally, you should have completed all parts of the [Tutorial on DER Hosting Capacity](https://github.com/Team-Nando#1-tutorial-on-der-hosting-capacity).
- Python (Anaconda) and Jupyter Notebook (comes with Anaconda). For download links and more info: https://www.anaconda.com/download. Note that you must install the Anaconda that is compatible with your operating system (e.g., Windows, Mac). Also, note that this repository is meant to be used by individuals (who can get free access to Anaconda).

## Run the Code
Make sure you have installed Anaconda on your computer. Otherwise, you will not be able to go through the repository.

To ensure that this repo runs as intended (correct versions of libraries, dependencies, etc.), we will create a new [environment](https://conda.io/projects/conda/en/latest/user-guide/tasks/manage-environments.html). Follow the steps below. Next time you run this repo, simply activate the environment, go to Step 4, and find the folder with the files.

1. Download all the files using the green **`<> Code`** button at the top right.
   - You will get a ZIP file with a folder that contains all the files.
   - Unzip the file and place the folder somewhere on your computer/laptop. The path of this folder will be your <folder_path>.
   - Also, unzip the profiles.zip folder and place the .npy files in the <folder_path> (where the MS Excel and .py files are located).
2. Create a new environment.
   - Open the Anaconda Prompt. 
   - Create the environment using the command `conda create --name myenv python=3.10` (replace 'myenv' with your desired environment name). Accept the installation of packages. It will take a few seconds.
   - Activate your environment with `conda activate myenv`. Once activated, you will see the `(myend)` on the left of the command line.
3. Install Jupyter Notebook and the required packages.
   - Using the same environment, type `pip install jupyter`. It will take a minute or so. 
   - After installing Jupyter Notebook, copy (Ctrl+C) the <folder_path>.
   - Then, type `cd <folder_path>`. Note that you are using your new environment.
   - Finally, type `pip install -r requirements.txt` to install the packages needed for the repo. It will take a minute or so.
4. To open the Jupyter Notebook file (extension **`ipynb`**) you need to:
   - In the Anaconda Prompt, type `Jupyter Notebook`. It will open in your browser. You will see all the files of this repo.
   - Open the **`ipynb`** file.

All the instructions to run the repo will be in the **`ipynb`** file.

Enjoy! 🤓

## Credits
### This Repo and Adaptations to the Original MV-LV Network Models
Eshan Karunarathne (akarunarathn@student.unimelb.edu.au)  
Orlando Pereira Guzman (opereiraguzm@student.unimelb.edu.au)  
Nando Ochoa (luis.ochoa@unimelb.edu.au ; https://sites.google.com/view/luisfochoa)

### Original Python Code
Andreas Procopiou (andreasprocopiou@ieee.org)

## Acknowledgement

We thank [AusNet Services](https://www.ausnetservices.com.au/) for providing the anonymised data of the four MV networks and the historical demand profiles used in this repository. These MV-LV networks were first created as part of the project [Advanced Planning of PV-Rich Distribution Networks](https://electrical.eng.unimelb.edu.au/power-energy/projects/pv-rich-distribution-networks). More details can be found in this report [[2]](https://www.researchgate.net/publication/334458042_Deliverable_1_HV-LV_modelling_of_selected_HV_feeders).

The content of this repository has been produced with direct and/or indirect inputs from multiple members (past and present) of Prof Nando Ochoa’s Research Team. So, special thanks to all of them (many of whom are now in different corners of the world).

* https://sites.google.com/view/luisfochoa/research/research-team
* https://sites.google.com/view/luisfochoa/research/past-team-members

## Licenses

Since this repository uses dss_python which is based on OpenDSS, both licenses have been included. This repository uses the BSD 3-Clause "New" or "Revised" license. Check all corresponding files (`LICENSE-OpenDSS`, `LICENSE-dss_python`, `LICENSE`).
