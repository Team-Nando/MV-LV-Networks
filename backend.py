import sys
import os
import glob
from colorama import Fore, Style
import numpy as np
import warnings
import pandas as pd
import geopandas as gp
from tqdm import tqdm
from tabulate import tabulate
from itertools import permutations
import re
from datetime import datetime, timedelta

import matplotlib.pyplot as plt
from shapely.geometry import Point, LineString
from matplotlib.lines import Line2D
import matplotlib.dates as mdates

try:
    import dss as dss_direct
    print(f"Your Python environment has 'dss_python: {dss_direct.__version__}'.", file=sys.stderr)
    if dss_direct.__version__ != "0.12.1":
        print(f"The dss_python version installed in your python environemnt is different from what we have used to develop this code (dss_python version 0.12.1). If you have troubles related to OpenDSS in the execution of power flow simulation, please re-install dss_python version 0.12.1 for smooth execution.")

except ModuleNotFoundError:
    raise ModuleNotFoundError(
        f"Module 'dss_python' not found. \n"
        f'  → Please install via command "pip install dss_python" in terminal.')
            
            
# user inputs
np.random.seed(100)

# %% Section 2: Classes and Definitions

class WrongNetworkNameError(Exception):
    pass


class MissingExcelSheetsError(Exception):
    pass


def check_data():
    """
    Check for the existence of a folder named 'Network_Data' in the current working directory.

    If the folder is found, adjust the working directory and print the path.
    If the folder is not found, print an error message indicating its absence.
    """
    current_directory = os.getcwd()
#     data_folder_name = "Network_Data"
#     data_folder_path = os.path.join(current_directory, data_folder_name)

    if os.path.exists(current_directory) and os.path.isdir(current_directory):
        os.chdir(current_directory)
        networks = [n for n in os.listdir(current_directory) if
                    n.lower().endswith(".xlsx") and os.path.isfile(os.path.join(current_directory, n))]
        if {'Network_3_Urban_HPK11.xlsx', 'Network_4_Urban_CRE21.xlsx', 'Network_1_Rural_SMR8.xlsx', 'Network_2_Rural_KLO14.xlsx'} != set(networks):
            raise MissingExcelSheetsError(f"Excel sheets are missing in this folder. Please Check")
        else:
            print("\033[1;92mAll the required excel files including network data for running this code is available.\033[0m")
    else:
        os.chdir(current_directory)
        print("\033[1;91mThere is some issue in the current directory.\033[0m")
    
    
    if os.path.exists(current_directory) and os.path.isdir(current_directory):
        os.chdir(current_directory)
        profiles = [n for n in os.listdir(current_directory) if
                    n.lower().endswith(".npy") and os.path.isfile(os.path.join(current_directory, n))]
        if {'Com_load_data_30min_res.npy', 'Res_load_data_30min_res.npy'} != set(profiles):
            raise MissingExcelSheetsError(f"Profiles data are missing in this folder. Please Check")
        else:
            print("\033[1;92mAll the required profile data for running this code is available.\033[0m")
    else:
        os.chdir(current_directory)
        print("\033[1;91mThere is some issue in the current directory.\033[0m")
        

#     return data_folder_path


def identify_network():
    """
    Prompt the user to enter a network name, convert it to uppercase, and print the identification process.
    """
    
    while True:
        try:
            user_input = input("Enter the number (1, 2, 3 or 4) associated to the network to study \n*Available Network names are \n    1) Network_1_Rural_SMR8, \n    2) Network_2_Rural_KLO14, \n    3) Network_3_Urban_HPK11, \n    4) Network_4_Urban_CRE21 : ")
            user_input = user_input.upper()
            checkmark = '\u2713'.encode('utf-8').decode('cp1252')
            options = ["1", "2", "3", "4"]
            net_names = ["Network_1", "Network_2", "Network_3", "Network_4"]
            
            if user_input.upper() in options:
                print("\033[1;92mUser requested network is: ", net_names[options.index(user_input)], "\033[0m")
                current_directory = os.getcwd()
                network_prefix = net_names[options.index(user_input)].split("_")
                
                # Search for Excel files matching the network name
                pattern = f"{current_directory}/{network_prefix[0]}_{network_prefix[1]}_*.xlsx"
                
                matched_files = glob.glob(pattern)
                
                if matched_files:
                    network_name = os.path.basename(matched_files[0])  # Assuming the first match is the correct file
                    data_folder_path = os.path.join(current_directory, network_name)
                    print("\033[1;92mThe circuit's information file is': ", network_name, "\033[0m")
                    all_data = NetworkData(name=user_input, data_location=data_folder_path)
                    data = all_data.get_network_data()
                    network_gp = all_data.network_plotting()
                    break
                else:
                    raise MissingExcelSheetsError(f"No Excel files found matching the network: {network_prefix}")
                
            else:
                raise WrongNetworkNameError(
                    f"There is no network associated to the number '{user_input}' in this repository."
                    f"The available options are 1, 2, 3 or 4. Please re enter a correct option."
                )

        except WrongNetworkNameError as e:
            print(f"User Input Error: {e}")
        except MissingExcelSheetsError as e:
            print(f"Network Data Error: {e}")
    
    return data, network_gp

class NetworkData:
    """
    tis is
    """

    def __init__(self, name, data_location):
        options = ["1", "2", "3", "4"]
        net_names = ["Rural_1", "Rural_2", "Urban_1", "Urban_2"]
        
        self.network_name: str = net_names[options.index(name.upper())]
        self.data_location: dict = data_location
        self.data: dict = {}
        self.get_network_data()
        self.tabulate_data()
        # self.network_plotting()
        

    def get_network_data(self):
        sheets = self.get_excel_sheet_names()
        for sheet_name in tqdm(sheets, desc=f"Extracting data: ", unit="sheet"):
            self.data[sheet_name] = pd.read_excel(self.data_location, sheet_name=sheet_name)
        
        return self.data

    def get_excel_sheet_names(self):
        try:
            xl = pd.ExcelFile(r""+self.data_location)
            sheet_names = xl.sheet_names
            
            return sheet_names
        
        except FileNotFoundError:
            print("File not found or incorrect path.")
            
            return None

    def tabulate_data(self):
        # for keys in self.data_locations.items():
        
        # Getting data from the dataframes:
            
        # Number of residential substations:
        n_res_txs = self.data['lvtx']["Type"].value_counts().get('RES', 0)
        # Number of residential customers:
        n_res_cus = self.data['lv_loads']["phases"].value_counts().get(1, 0)
        # Number of non-residential substations:
        n_com_txs = self.data['lvtx']["Type"].value_counts().get('COM', 0)
        # Number of non-residential customers:
        n_com_cus = self.data['lv_loads']["phases"].value_counts().get(3, 0)
        # Number of MV/MV transformers
        try:
            n_mvmv_txs = (~self.data['mvtx']["Substation_ID"].str.contains('_REG')).sum()
        except:
            n_mvmv_txs = 0
        # Number of Regulators
        try:
            n_reg_txs = self.data['mvtx']["Substation_ID"].str.contains('_REG').sum()
        except:
            n_reg_txs = 0
        # Number of capacitors:
        try:
            n_caps = self.data['mvcaps']["phases"].value_counts().get(3, 0)
        except:
            n_caps = 0
        # Number of SWER transformers
        n_swer_txs = len(self.data['lvtx'][self.data['lvtx']["Conn_Type"].astype(str).str.len() != 3])
        # Length of MV conductors
        len_mv_lines = self.data["lines"]["Length"].sum()
        # Length of MV SWER conductors
        len_mv_swer_lines = self.data["lines"].loc[self.data["lines"]["Phases"] == 1]["Length"].sum()
        # Length of LV conductors
        len_lv_lines = self.data["lv_lines"]["length"].sum()
        # Length of LV SWER conductors
        len_lv_swer_lines = self.data["lv_lines"].loc[self.data["lv_lines"]["phases"] == 1]["length"].sum()
            
    
        # Table creation
        data_tuples = [
            ('# of LV Residential Substations', n_res_txs),
            ('# of LV Residential Customers', n_res_cus),
            ('# of LV Non-Residential Substations', n_com_txs),
            ('# of LV Non-Residential Customers', n_com_cus),
            ('# of SWER MV Transformers', n_mvmv_txs),
            ('# of Voltage Regulators', n_reg_txs),
            ('# of MV Capacitors', n_caps),
            ('# of SWER LV Transformers', n_swer_txs),
            ('# of MV conductor length', f"{np.round(len_mv_lines,2)} km"),
            ('# of MV SWER conductor length', f"{np.round(len_mv_swer_lines,2)} km"),
            ('# of LV conductor length', f"{np.round(len_lv_lines/1000,2)} km"),
            ('# of LV SWER conductor length', f"{np.round(len_lv_swer_lines/1000,2)} km"),
        ]
    
        # Filter out rows where the second element (quantity) is zero
        filtered_data = [item for item in data_tuples if item[1] != 0 and item[1] != '0.0 km' and item[1] != '0.00 km']
    
        # Print a nicely formatted table
        print("\033[1;92mNetwork Data\033[0m")
        print(tabulate(filtered_data, headers=['Parameter', 'Quantity'], tablefmt="github"))
        
    def network_plotting(self):
        
        " This function creates the plot of the circuit by using geopandas"
        
        # Coordinate system to use (for Australian cases)
        desired_crs = 'EPSG:4462'
        
        ####### MV Nodes setup:
            
        # Calling the file with the node information
        mvbuses_layer = self.data["buscoords"]

        # Create Point geometries from X and Y coordinates
        geometry = [Point(xy) for xy in zip(mvbuses_layer['NodeStartX'], mvbuses_layer['NodeStartY'])]

        # Create a GeoDataFrame
        mvbuses_layer_gp = gp.GeoDataFrame(mvbuses_layer, geometry=geometry, crs=desired_crs)  
        # Set the geometry column name
        mvbuses_layer_gp = mvbuses_layer_gp.set_geometry('geometry')
        # Set the GeoDataframe index
        mvbuses_layer_gp.index = mvbuses_layer_gp["Node_ID"]
        
        ####### MV transformer:

        # Calling the file with the node information
        mvtx_layer = self.data["mv_net_txs"]

        # Create geometries using nodes' geometries from 'mv_buses_layer_gp'
        mvtx_geometries = []
        for idx, row in mvtx_layer.iterrows():
            start_node = row['Bus2']
            
            mvtx_geometry = mvbuses_layer_gp.loc[start_node, 'geometry']  # Get start node geometry
            mvtx_geometries.append(mvtx_geometry)

        self.mvtxs_layer_gp = gp.GeoDataFrame(mvtx_layer, geometry=mvtx_geometries, crs=desired_crs)
        
        ####### MV lines setup:
        
        # Calling the file with the node information
        mvlines_layer = self.data["lines"].loc[self.data["lines"]["Element_Name"].str.lower()!="delete"]

        # Create LineString geometries using nodes' geometries from 'mv_buses_layer_gp'
        line_geometries = []
        for idx, row in mvlines_layer.iterrows():
            start_node = row['Start_Node']
            end_node = row['End_Node']
            
            start_geometry = mvbuses_layer_gp.loc[start_node, 'geometry']  # Get start node geometry
            end_geometry = mvbuses_layer_gp.loc[end_node, 'geometry']      # Get end node geometry
            
            line = LineString([start_geometry, end_geometry])  # Create LineString using start and end geometries
            line_geometries.append(line)

        self.mvlines_layer_gp = gp.GeoDataFrame(mvlines_layer, geometry=line_geometries, crs=desired_crs)
        
        ###### MV/LVTransformers setup

        # Calling the file with the node information
        txs_layer = self.data['lvtx']

        # Create geometries using nodes' geometries from 'mv_buses_layer_gp'
        tx_geometries = []
        for idx, row in txs_layer.iterrows():
            start_node = row['Bus1']
            
            tx_geometry = mvbuses_layer_gp.loc[start_node, 'geometry']  # Get start node geometry
            tx_geometries.append(tx_geometry)

        self.txs_layer_gp = gp.GeoDataFrame(txs_layer, geometry=tx_geometries, crs=desired_crs)
        
        ###### Capacitors setup
        cap_flag = False
        try:
            caps_layer = self.data['mvcaps']
        
            # Create geometries using nodes' geometries from 'mv_buses_layer_gp'
            caps_geometries = []
            for idx, row in caps_layer.iterrows():
                start_node = row['Bus1']
                
                cap_geometry = mvbuses_layer_gp.loc[start_node, 'geometry']  # Get start node geometry
                caps_geometries.append(cap_geometry)
    
            self.caps_layer_gp = gp.GeoDataFrame(caps_layer, geometry=caps_geometries, crs=desired_crs)
            cap_flag = True
        except:
            pass
        
        
        ###### Saving the GeoDataFrames
        
        self.network_gp = {}
        self.network_gp["MV_tx"] = self.mvtxs_layer_gp
        self.network_gp["MV_lines"] = self.mvlines_layer_gp
        self.network_gp["MVLV_txs"] = self.txs_layer_gp
        if cap_flag:
            self.network_gp["caps"] = self.caps_layer_gp
        
        ###### Plotting
        if "urban" in self.network_name.lower():
    
            # Create a figure and axis
            fig, ax = plt.subplots(figsize=(8, 8), dpi=300)
    
            # Plot the GeoDataFrames
            self.mvtxs_layer_gp.plot(ax=ax, color="black", alpha=0.7, marker="^", markersize=264, label='MV tx', zorder=1)
            self.mvlines_layer_gp.plot(ax=ax, color="grey", alpha=0.7, label='MV Lines', zorder=2, linewidth=2.5)
            self.txs_layer_gp.loc[self.txs_layer_gp["Type"] == "COM"].plot(ax=ax, color="blue", alpha=1, marker="o", markersize=64, label='LV txs_com', zorder=3)
            self.txs_layer_gp.loc[self.txs_layer_gp["Type"] == "RES"].plot(ax=ax, color="black", alpha=1, marker="o", markersize=64, label='LV txs_res', zorder=3)
            if cap_flag:
                self.caps_layer_gp.plot(ax=ax, color="orange", alpha=1, marker="s", markersize=64, label='LV txs_res', zorder=3)
    
    
            ax.set_xticks([])
            ax.set_yticks([])
            if cap_flag:
                legend_elements = [
                    # Line2D([0], [0], color='grey', lw=2, label='MV Lines'),
                    Line2D([0], [0], marker='^', color='black', markersize=10, label='Main MV tx', linestyle='None'),
                    Line2D([0], [0], marker='o', color='black', markersize=5, label='LV txs, RES', linestyle='None'),
                    Line2D([0], [0], marker='o', color='blue', markersize=5, label='LV txs, COM', linestyle='None'),
                    Line2D([0], [0], marker='s', color='orange', markersize=5, label='MV Caps', linestyle='None')
                ]
            else:
                legend_elements = [
                    # Line2D([0], [0], color='grey', lw=2, label='MV Lines'),
                    Line2D([0], [0], marker='^', color='black', markersize=10, label='Main MV tx', linestyle='None'),
                    Line2D([0], [0], marker='o', color='black', markersize=5, label='LV txs, RES', linestyle='None'),
                    Line2D([0], [0], marker='o', color='blue', markersize=5, label='LV txs, COM', linestyle='None')
                ]
    
            ax.legend(handles=legend_elements, loc='lower right')
            ax.set_title("Network Topology", fontsize=10)
            plt.tight_layout()
            plt.show()
        
        else:
            # Create a figure and axis
            fig, ax = plt.subplots(figsize=(8, 8), dpi=300)
    
            # Plot the GeoDataFrames
            self.mvtxs_layer_gp.plot(ax=ax, color="black", alpha=0.7, marker="^", markersize=264, label='MV tx', zorder=1)
            self.mvlines_layer_gp.plot(ax=ax, color="grey", alpha=0.7, label='MV Lines', zorder=2, linewidth=2.5)
            self.txs_layer_gp.loc[(self.txs_layer_gp["Type"] == "COM") & (self.txs_layer_gp["Conn_Type"].str.len() == 3)].plot(ax=ax, color="blue", alpha=1, marker="o", markersize=32, label='MV/LV txs - 3ph - Com', zorder=3)
            self.txs_layer_gp.loc[(self.txs_layer_gp["Type"] == "RES") & (self.txs_layer_gp["Conn_Type"].str.len() == 3)].plot(ax=ax, color="black", alpha=1, marker="o", markersize=32, label='MV/LV txs - 3ph - Res', zorder=3)
            self.txs_layer_gp.loc[(self.txs_layer_gp["Type"] == "RES") & (self.txs_layer_gp["Conn_Type"].str.len() != 3)].plot(ax=ax, color="sandybrown", alpha=1, marker="o", markersize=32, label='MV/LV txs - 1ph - Res', zorder=3)
            if cap_flag:
                self.caps_layer_gp.plot(ax=ax, color="darkslateblue", alpha=1, marker="s", markersize=64, label='LV txs_res', zorder=3)
    
    
            ax.set_xticks([])
            ax.set_yticks([])
            if cap_flag:
                legend_elements = [
                    # Line2D([0], [0], color='grey', lw=2, label='MV Lines'),
                    Line2D([0], [0], marker='^', color='black', markersize=10, label='Main MV tx', linestyle='None'),
                    Line2D([0], [0], marker='o', color='black', markersize=5, label='LV txs, RES - 3PH', linestyle='None'),
                    Line2D([0], [0], marker='o', color='sandybrown', markersize=5, label='LV txs, RES - SWER', linestyle='None'),
                    Line2D([0], [0], marker='o', color='blue', markersize=5, label='LV txs, COM - 3PH', linestyle='None'),
                    Line2D([0], [0], marker='s', color='darkslateblue', markersize=5, label='MV Caps', linestyle='None')
                ]
            else:
                legend_elements = [
                    # Line2D([0], [0], color='grey', lw=2, label='MV Lines'),
                    Line2D([0], [0], marker='^', color='black', markersize=10, label='Main MV tx', linestyle='None'),
                    Line2D([0], [0], marker='o', color='black', markersize=5, label='LV txs, RES - 3PH', linestyle='None'),
                    Line2D([0], [0], marker='o', color='darkslateblue', markersize=5, label='LV txs, RES - SWER', linestyle='None'),
                    Line2D([0], [0], marker='o', color='blue', markersize=5, label='LV txs, COM - 3PH', linestyle='None')
                ]
    
            ax.legend(handles=legend_elements, loc='lower right')
            ax.set_title("Network Topology", fontsize=10)
            plt.tight_layout()
            plt.show()
            
        return self.network_gp


class DSSDriver:
    """
    This class is to
    """

    def __init__(self, data, gis_data):

        self.data = data
        self.gis_data = gis_data
       

        self.dss = dss_direct.DSS  # type dss_direct.dss_capi_gr.IDSS
        self.dss.Start(0)
        self.dss_text = self.dss.Text
        self.dss_circuit = self.dss.ActiveCircuit
        self.dss_solution = self.dss.ActiveCircuit.Solution
        self.ready = False
        
        # Add more things
        self.build_network
        # Run an snapshot
        # V_vals = self.run_snapshot
        
        

    def build_network(self):
        print("\033[1;92mNetwork building started.\033[0m")
        self.basic_opendss_actions()
        self.voltage_source()
        self.mv_net_tx()
        self.line_codes()
        self.connections()
        try:
            self.capacitors()
        except:
            pass
        try:
            self.mv_txs()
        except:
            pass
        self.lv_tx()
        self.lv_nets()
        self.lv_load_mod()
        self.print_selected_date()
        self.other_opendss_commands()
        print("\033[1;92mNetwork building completed.\033[0m")
    
    def get_date_and_season(day_of_year):
        # Define the start date of the year
        start_of_year = datetime(year=datetime.now().year, month=1, day=1)
    
        # Calculate the date
        date = start_of_year + timedelta(days=day_of_year - 1)  # Subtract 1 because day_of_year is 1-indexed
        date_str = date.strftime("%B %d")
    
        # Determine the season
        if 355 <= day_of_year or day_of_year <= 78:
            season = "Summer"
        elif 79 <= day_of_year <= 170:
            season = "Autumn"
        elif 171 <= day_of_year <= 263:
            season = "Winter"
        else:
            season = "Spring"
    
        return date_str, season
    
    def basic_opendss_actions(self):
        self.dss_text.Command = 'clear'
        self.dss_text.Command = 'Set DefaultBaseFrequency = 50'

    def voltage_source(self):
        self.dss_text.Command = f"New circuit.circuit basekv=66 pu=1 angle=0 phases=3 R1=0.52824 X1=2.113 R0=0.59157 X0=1.7747"
        
        self.dss_text.Command = f"edit vsource.source bus1=sourcebus basekv=66 pu=1 angle=0 phases=3 R1=0.52824 X1=2.113 R0=0.59157 X0=1.7"
        

    def mv_net_tx(self):
        element = self.data["mv_net_txs"]
        for index, row in tqdm(element.iterrows(), desc="Building the circuit - MV Transformers  : ", total=len(element)):
            mv_tx_data = (f"New Transformer.{row['Substation_ID']} "
                          f"phases=3 "
                          f"windings=2 "
                          f"buses=[{row['Bus1']}, mv_f0_n{row['Bus2']}] "
                          f"conns=[{row['Connection_Primary']}, {row['Connection_Secondary']}] "
                          f"kVs=[{row['kvs_primary']}, {row['kvs_secondary']}] "
                          f"kVAs=[{row['kvas_primary']}, {row['kvas_secondary']}] "
                          f"%loadloss={row['loadloss']} "
                          f"%noloadloss={row['noloadloss']} "
                          f"xhl={row['xhl']} "
                          f"enabled=true")
            self.dss_text.Command = mv_tx_data
        
    def line_codes(self):
        element = self.data["linecodes"]
        for index, row in tqdm(element.iterrows(), desc="Building the circuit - Linecodes        : ", total=len(element)):
            linecode_data = (f"new linecode.lc_{row['Linecode_ID']} "
                             f"nphases={row['Phases']} "
                             f"r1={row['r1']} "
                             f"x1={row['x1']} "
                             f"b1={row['b1']} "
                             f"r0={row['r0']} "
                             f"b0={row['b0']} "
                             f"x0={row['x0']} "
                             f"units={row['Units']} "
                             f"normamp={min(row['Ampacity1'], row['Ampacity2'])} ")
            
            # print(linecode_data)
            self.dss_text.Command = linecode_data

    def connections(self):
        element = self.data["lines"]
        for index, row in tqdm(element.iterrows(), desc="Building the circuit - MV Lines         : ", total=len(element)):
            if row["Element_Name"].lower() != "delete":
                mv_line_data = (f"new line.mv_f0_l{row['Line_Number']} "
                                f"bus1=mv_f0_n{row['Start_Node']}.{row['Start_Node_Phase']} "
                                f"bus2=mv_f0_n{row['End_Node']}.{row['End_Node_Phase']} "
                                f"phases={row['Phases']} "
                                f"length={row['Length']} "
                                f"units={row['Units']} "
                                f"linecode=lc_{row['Linecode']}-{row['Phases']}ph "
                                f"enabled=true")
                self.dss_text.Command = mv_line_data
                self.gis_data["MV_lines"].loc[index, "DSSNAME"] = f"mv_f0_l{row['Line_Number']}"
                
                # Ampacity determination:
                amp = self.data["linecodes"].loc[self.data["linecodes"]["Linecode_ID"] == f"{row['Linecode']}-{row['Phases']}ph", "Ampacity1" ].values[0]
                self.gis_data["MV_lines"].loc[index, "Ampacity"] = amp
            
    def capacitors(self):
        element = self.data["mvcaps"]
        for index, row in tqdm(element.iterrows(), desc="Building the circuit - MV Capacitors    : ", total=len(element)):
            mv_caps_data = (f"new capacitor.mv_f0_l{row['Element_ID']} "
                            f"bus1=mv_f0_n{row['Bus1']}.1.2.3 "
                            f"phases={row['phases']} "
                            f"kvar={row['kvar']} "
                            f"kV={row['kvs']}")
            self.dss_text.Command = mv_caps_data
    
    def mv_txs(self):
        three_ph_combi = [''.join(perm) for perm in list(permutations(['R', 'W', 'B']))]
        two_ph_combi = [''.join(perm) for perm in list(permutations(['R', 'B', 'W'], 2))]
        one_ph_combi = [''.join(perm) for perm in list(permutations(['R', 'B', 'W'], 1))]
        char_to_num = {key: '.'.join(str({'R': 1, 'W': 2, 'B': 3}[char]) for char in key) for key in three_ph_combi + two_ph_combi+one_ph_combi}
        
        element = self.data["mvtx"]
        
        for index, row in element.iterrows():
            kva1 = row["kvs_primary"]
            kva2 = row["kvs_secondary"]
            
            if kva1 != kva2:
            
                mv_tx = (f"new transformer.{row['Substation_ID']} "
                                f"buses=[mv_f0_n{row['Bus1']}.{char_to_num[row['Conn_Type']]} mv_f0_n{row['Bus2']}.1.0 mv_f0_n{row['Bus2']}.0.2] "
                                f"phases=1 "
                                f"windings=3 "
                                f"conns=[Delta Wye Wye] "
                                f"kVs=[{row['kvs_primary']} {row['kvs_secondary']} {row['kvs_secondary']}] "
                                f"kVAs=[{row['kvas_primary']} {row['kvas_secondary']} {row['kvas_secondary']}] "
                                f"xhl={row['xhl']} "
                                f"%noloadloss={row['noloadloss']} "
                                f"%loadloss={row['loadloss']}")

                self.dss_text.Command = mv_tx
            
            # if kva1 == kva2, then it is a autotransformer regulator
            else: 
                
                #Regulator
                self.dss_text.Command = "set maxcontroliter=100"
                
                # Reactors
                
                # One reactor per phase
                # There are two jumpers (per phase), one "entry" jumper and one out "jumper".
                # The entry jumper goes from bus1 to the bus for the transformer
                # The out jumper goes from the transformer bus to bus2
                
                # Phase A
                self.dss_text.Command = f"New Reactor.Jumper_{row['Substation_ID']}_A_E phases=1 bus1=mv_f0_n{row['Bus1']}.1 bus2=Jumper_{row['Substation_ID']}_A.2 X=0.0001 R=0.0001"
                self.dss_text.Command = f"New Reactor.Jumper_{row['Substation_ID']}_A_O phases=1  bus1=Jumper_{row['Substation_ID']}_A.1 bus2=mv_f0_n{row['Bus2']}.1 X=0.0001 R=0.0001"
                
                # Phase B
                self.dss_text.Command = f"New Reactor.Jumper_{row['Substation_ID']}_B_E phases=1 bus1=mv_f0_n{row['Bus1']}.2 bus2=Jumper_{row['Substation_ID']}_B.2 X=0.0001 R=0.0001"
                self.dss_text.Command = f"New Reactor.Jumper_{row['Substation_ID']}_B_O phases=1  bus1=Jumper_{row['Substation_ID']}_B.1 bus2=mv_f0_n{row['Bus2']}.2 X=0.0001 R=0.0001"
                
                # Phase C
                self.dss_text.Command = f"New Reactor.Jumper_{row['Substation_ID']}_C_E phases=1 bus1=mv_f0_n{row['Bus1']}.3 bus2=Jumper_{row['Substation_ID']}_C.2 X=0.0001 R=0.0001"
                self.dss_text.Command = f"New Reactor.Jumper_{row['Substation_ID']}_C_O phases=1  bus1=Jumper_{row['Substation_ID']}_C.1 bus2=mv_f0_n{row['Bus2']}.3 X=0.0001 R=0.0001"
                
                
                # Now the transformers (1 per phase)
                
                # Remember, if using a 2 winding model definition for an autotransformer, there
                # has to be a recalculation of the equivalent kVA from the transfomer.
                # Using equation 7.48 from Kersting 3rd edition
                
                # Also, at the end, the regcontrol must be created. Otherwise, there's no regulation
                
                kv = 12.7
                nt = 10/100 # Assuming transformer can regulate up to 10% (typical)
                
                kVAtraf = str(np.round((nt * float(kva1) / (1 + nt)),2))
                
                for phase in ["A", "B", "C"]:
                    mv_tx = (f"new transformer.{row['Substation_ID']}_{phase} "
                             f"phases=1 "
                             f"windings=2 "
                             f"xhl={row['xhl']} "
                             f"%noloadloss={row['noloadloss']} "
                             f"%loadloss={row['loadloss']} "
                             f"wdg=1 "
                             f"Bus=Jumper_{row['Substation_ID']}_{phase}.1.0 "
                             f"kV={kv} "
                             f"kVA={kVAtraf} "
                             f"wdg=2 "
                             f"Bus=Jumper_{row['Substation_ID']}_{phase}.1.2 "
                             f"kV={kv/10} "
                             f"kVA={kVAtraf} "
                             f"Maxtap=1.0 "
                             f"Mintap=-1.0 "
                             f"tap=0.0 "
                             f"numtaps={row['wdg1_numtaps']-1}")
                    
                    # print(mv_tx)
                    self.dss_text.Command = mv_tx
                
                # Now the reg control
                for phase in ["A", "B", "C"]:
                    reg_control = (f" new regcontrol.Reg_{row['Substation_ID']}_{phase} "
                                   f"transformer={row['Substation_ID']}_{phase} "
                                   f"winding=2 "
                                   f"bus=Jumper_{row['Substation_ID']}_{phase}.1 "
                                   f"vreg=100.0 "
                                   f"band=3.0 "
                                   f"ptratio={kv*10} "
                                   f"maxtapchange=1")
                    
                    # print(reg_control)
                    self.dss_text.Command = reg_control
        return 

    def lv_tx(self):
        three_ph_combi = [''.join(perm) for perm in list(permutations(['R', 'W', 'B']))]
        two_ph_combi = [''.join(perm) for perm in list(permutations(['R', 'B', 'W'], 2))]
        one_ph_combi = [''.join(perm) for perm in list(permutations(['R', 'B', 'W'], 1))]
        char_to_num = {key: '.'.join(str({'R': 1, 'W': 2, 'B': 3}[char]) for char in key) for key in three_ph_combi + two_ph_combi+one_ph_combi}

        element = self.data["lvtx"]
        for index, row in tqdm(element.iterrows(), desc="Building the circuit - LV Transformers  : ", total=len(element)):
            if len(row['Conn_Type']) == 3:
                lv_tx_data = (f"new transformer.mv_f0_lv_{row['Substation_ID']} "
                              f"phases=3 "
                              f"windings=2 "
                              f"buses=[mv_f0_n{row['Bus1']} mv_f0_lv{index}_busbar] "
                              f"conns=[{row['Connection_Primary']} {row['Connection_Secondary']}] "
                              f"kVs=[{row['kvs_primary']} {row['kvs_secondary']}] "
                              f"kVAs=[{row['kvas_primary']} {row['kvas_secondary']}] "
                              f"xhl={row['xhl']} "
                              f"%noloadloss={row['noloadloss']} "
                              f"%loadloss={row['loadloss']} "
                              f"wdg=1 "
                              f"numtaps=4 "
                              f"tap={row['wdg1_tap']} "
                              f"maxtap=1.137 "
                              f"mintap=1.028")
                # self.dss_text.Command = lv_tx_data
                # print(lv_tx_data)
            elif len(row['Conn_Type']) == 2:
                lv_tx_data = (f"new transformer.mv_f0_lv_{row['Substation_ID']} "
                              f"phases=1 "
                              f"windings=3 "
                              f"buses=[mv_f0_n{row['Bus1']}.{char_to_num[row['Conn_Type']]} mv_f0_lv{index}_busbar.1.0 mv_f0_lv{index}_busbar.0.2] "
                              f"conns=[Delta Wye Wye] "
                              f"kVs=[22 0.25 0.25] "
                              f"kVAs=[{row['kvas_primary']} {row['kvas_secondary']} {row['kvas_secondary']}] "
                              f"xhl={row['xhl']} "
                              f"%noloadloss={row['noloadloss']} "
                              f"%loadloss={row['loadloss']} "
                              f"wdg=1 "
                              f"numtaps=4 "
                              f"tap={row['wdg1_tap']} "
                              f"maxtap=1.137 "
                              f"mintap=1.028")
            else:
                lv_tx_data = (f"new transformer.mv_f0_lv_{row['Substation_ID']} "
                              f"phases=1 "
                              f"windings=2 "
                              f"buses=[mv_f0_n{row['Bus1']}.{char_to_num[row['Conn_Type']]} mv_f0_lv{index}_busbar.1] "
                              f"conns=[Wye Wye] "
                              f"kVs=[{row['kvs_primary']} {row['kvs_secondary']}] "
                              f"kVAs=[{row['kvas_primary']} {row['kvas_secondary']}] "
                              f"xhl={row['xhl']} "
                              f"%noloadloss={row['noloadloss']} "
                              f"%loadloss={row['loadloss']} "
                              f"wdg=1 "
                              f"numtaps=4 "
                              f"tap={row['wdg1_tap']} "
                              f"maxtap=1.137 "
                              f"mintap=1.028")
                # self.dss_text.Command = lv_tx_data
            
            self.dss_text.Command = lv_tx_data
            self.gis_data["MVLV_txs"].loc[index, "DSSNAME"] = f"mv_f0_lv_{row['Substation_ID']}"

    def lv_nets(self):
        element = self.data["lv_lines"]
        for index, row in tqdm(element.iterrows(), desc="Building the circuit - LV Lines         : ", total=len(element)):
            if row['phases'] == 3:   
                bus_conn = ".1.2.3"
            else:
                bus_conn = ".1"
                
            line_data = (f"new line.{row['line_name']} "
                         f"bus1={row['bus1']}{bus_conn} "
                         f"bus2={row['bus2']}{bus_conn} "
                         f"phases={row['phases']} "
                         f"length={row['length']} "
                         f"units={row['units']} "
                         f"linecode={row['linecode']} ")
            # print(line_data)
            self.dss_text.Command = line_data

    def lv_load_mod(self, selected_day=0):
        # Loadshape initial settings
        profile_locataion = f"{os.getcwd()}"
        house_data = np.load(f"{profile_locataion}\Res_load_data_30min_res.npy")
        com_data = np.load(f"{profile_locataion}\Com_load_data_30min_res.npy")
        time_res = 30
        if selected_day==0:
            np.random.seed(100)
            selected_day = int(np.random.randint(0, 365))
        else:
            selected_day = selected_day
        
        # Load definition
        element = self.data["lv_loads"]
        for index, row in tqdm(element.iterrows(), desc="Building the circuit - Customer loadings: ", total=len(element)):
                    
            if row['phases'] == 1:
                
                load_data = (f"new load.{row['load_name']} "
                              f"phases={row['phases']} "
                              f"bus1={row['bus1']} "
                              f"kw=1 "
                              f"conn=wye " 
                              f"kv={row['kv']} "
                              f"pf={row['pf']} "
                              f"model=1 "
                              f"vminpu=0.0 vmaxpu=2 "
                              f"status={row['model']} "
                              f"enabled=True")
                
                self.dss_text.Command = load_data

                load_profile_res = house_data[np.random.randint(len(house_data)), selected_day, :]
                load_shape = (f'New Loadshape.Load_shape_res_{index} '
                              f'npts={int((24 * 60) / time_res)} '
                              f'minterval={time_res} '
                              f'Pmult={load_profile_res.tolist()} '
                              f'useactual=no')
                self.dss_text.Command = load_shape
                
                # Then, associate the profile to a customer
                self.dss_circuit.SetActiveElement(f"load.{row['load_name']}")
                self.dss_circuit.ActiveElement.Properties('daily').Val = f"Load_shape_res_{index}"
                

            else:
                load_data = (f"new load.{row['load_name']} "
                              f"phases={row['phases']} "
                              f"bus1={row['bus1']} "
                              f"kw=1 "
                              f"conn=wye " 
                              f"kv={row['kv']} "
                              f"pf={row['pf']} "
                              f"model=1 "
                              f"vminpu=0.0 vmaxpu=2 "
                              f"status={row['model']} "
                              f"enabled=True")
        
                self.dss_text.Command = load_data
                
                while True:
                
                    load_profile_com = com_data[np.random.randint(len(com_data)), selected_day, :]
                    load_av = np.mean(load_profile_com)
                    load_max = np.max(load_profile_com)
                    
                    if load_max < row['tx_cap']/2:
                        break
                
                load_shape = (f'New Loadshape.Load_shape_com_{index} '
                              f'npts={int((24 * 60) / time_res)} '
                              f'minterval={time_res} '
                              f'Pmult={load_profile_com.tolist()} '
                              f'useactual=no')
                self.dss_text.Command = load_shape
                
                # Then, associate the profile to a customer
                self.dss_circuit.SetActiveElement(f"load.{row['load_name']}")
                self.dss_circuit.ActiveElement.Properties('daily').Val = f"Load_shape_com_{index}"
        
        return selected_day
        
    def print_selected_date(self, selected_day):
        
        date_str, season = DSSDriver.get_date_and_season(selected_day)
        
        print(f"The random day selected for the simulation is {date_str}, and it correspond to the {season} season.")
    
    def other_opendss_commands(self):
        self.dss_text.Command = 'Set VoltageBases=[66.0, 22.0, 12.7 0.400, 0.2309]'
        self.dss_text.Command = 'calcv'
        
    def sel_time(snapshottime):
        
        # Input for specifying the snapshot time
        
        while True:
            # Prompt user for time input
            # snapshottime = input("Enter time (hh:mm): ")
    
            # Validate the format and values
            try:
                h, m = snapshottime.split(':')
                h, m = int(h), int(m)
    
                if not (0 <= h < 24) or not (0 <= m < 60):
                    raise ValueError("Invalid hour or minute value.")
    
                # Round the minutes
                if m < 15:
                    m = '00'
                elif 15 <= m < 45:
                    m = '30'
                else:
                    h = h + 1 if h < 23 else 0
                    m = '00'
    
                # Combine hours and minutes into the rounded time
                rounded_time = f"{h:02d}:{m}"
                print(f"Rounded time: {rounded_time}")
                break
    
            except ValueError:
                print("Invalid input. Please enter time in the format hh:mm where hh is 00-23 and mm is 00-59.")
        
        sim_hh = rounded_time.split(":")[0]
        sim_mm = rounded_time.split(":")[1]
        
        return sim_hh, sim_mm
    
    def Voltge_profile_plot(self):
        plt.rcParams["font.size"] = 18
    
        nodes = list(self.temp_all_V_values.index)
    
        mv_nodes = {phase: [] for phase in ['1', '2', '3']}
        lv_nodes = {phase: [] for phase in ['1', '2', '3']}
    
        for node in nodes:
            node_str = str(node).lower()
            if ("mv" in node_str and "lv" not in node_str) or "source" in node_str:
                phase = node_str[-1]
                mv_nodes[phase].append(node)
            elif "lv" in node_str:
                phase = node_str[-1]
                lv_nodes[phase].append(node)
    
        V_mv = {phase: pd.DataFrame(index=mv_nodes[phase], columns=['Val', 'Distance']) for phase in ['1', '2', '3']}
        V_lv = {phase: pd.DataFrame(index=lv_nodes[phase], columns=['Val', 'Distance']) for phase in ['1', '2', '3']}
    
        for node in list(self.temp_all_V_values.index):
            phase = str(node)[-1]
            if node in mv_nodes[phase]:
                V_mv[phase].loc[node, 'Val'] = self.temp_all_V_values.loc[node, 'Val']
                V_mv[phase].loc[node, 'Distance'] = self.temp_all_V_values.loc[node, 'Distance']
            elif node in lv_nodes[phase]:
                if self.temp_all_V_values.loc[node, 'Val'] > 0.1:
                    V_lv[phase].loc[node, 'Val'] = self.temp_all_V_values.loc[node, 'Val']
                    V_lv[phase].loc[node, 'Distance'] = self.temp_all_V_values.loc[node, 'Distance']
    
        # PLOT
        legend = []
        fig1, ax1 = plt.subplots(dpi=300, figsize=(16, 6))
    
        markers = {'1': {'mv': 'ro', 'lv': 'r.'},
                   '2': {'mv': 'ko', 'lv': 'k.'},
                   '3': {'mv': 'bo', 'lv': 'b.'}}
    
        for phase in ['1', '2', '3']:
            if not V_mv[phase].empty:
                plt.plot(V_mv[phase]['Distance'], V_mv[phase]['Val'], markers[phase]['mv'], markersize=3.5)
                legend.append(f"MV phase {phase}")
            if not V_lv[phase].empty:
                plt.plot(V_lv[phase]['Distance'], V_lv[phase]['Val'], markers[phase]['lv'], markersize=2.5)
                legend.append(f"LV phase {phase}")
    
        ax1.grid(color='grey', linestyle='-', linewidth=0.2)
        ax1.legend(legend, loc='lower center', bbox_to_anchor=(0.5, -0.325), borderaxespad=0., ncol=3)
        ax1.set_ylabel("Voltage, pu")
        ax1.set_xlabel("Distance from the MV substation, km")
        plt.xticks()
        plt.yticks()
        #plt.tight_layout()
        plt.show()
        
        # Explanation
        
        n_reg_txs = 0
        try:
            n_reg_txs = self.data['mvtx']["Substation_ID"].str.contains('_REG').sum()
        except:
            pass
        
        if n_reg_txs > 0:
            print("This circuit has voltage regulators. These are devices typically used on long distribution networks with the aim of raising the voltage levels on zones far from the source. These regulators are essentially autotransformers with a control on the tap positions, which will act according to the settings of the control. In this case, the settings on the voltage regulators are working to maintain a voltage of 1.00 pu (22 kV line to line) at the secondary side, explaining why there are some gaps between the MV points on the voltage profile plot.")
        
    def Voltge_timeseries_plot(self):
        plt.rcParams["font.size"] = 18
        idx_date = pd.date_range('2021-01-01 00:00', '2021-01-01 23:30', freq = '30min')
        hours = mdates.HourLocator(interval = 4)
        h_fmt = mdates.DateFormatter('%H:%M')

        fig, axes = plt.subplots(dpi=300, figsize=(16, 6))  # Adjust figsize as needed

        for column in self.Vdata.columns:
            axes.plot(idx_date, self.Vdata[column], color='blue', alpha=0.2)
            axes.xaxis.set_major_locator(hours)
            axes.xaxis.set_major_formatter(h_fmt)
            axes.set_xlabel('Time of the day')
            axes.set_ylabel('Voltages [V]')
            axes.grid(color='grey', linestyle='-', linewidth=0.2)
        axes.axhline(y=253, color='red', linestyle='--')
        # Adjust layout and show the plot
        legend_elements = [
            Line2D([0], [0], color='blue', alpha=0.2,  lw=2, label='Customers'),
             Line2D([0], [0], color='red', linestyle='--', lw=2, label='Upper V limit'),
        ]
        axes.legend(handles=legend_elements)
        plt.tight_layout()
        plt.show()
        
        
    def Tx_utilisation_plot(self):
        plt.rcParams["font.size"] = 18
        idx_date = pd.date_range('2021-01-01 00:00', '2021-01-01 23:30', freq = '30min')
        hours = mdates.HourLocator(interval = 4)
        h_fmt = mdates.DateFormatter('%H:%M')
        
        if (self.data["lvtx"]["Conn_Type"].str.len() == 3).all():
            
            fig, axes = plt.subplots(dpi=300, figsize=(16, 6))  # Adjust figsize as needed
            
            # MV transformer
            tx_cap = self.data["mv_net_txs"].loc[0, "kvas_primary"]
            axes.plot(idx_date, 100*self.Sdata["Value"]/tx_cap, color="blue", zorder=3, linewidth=3.5)
            
            # MV/LV transformers
            for i, column in enumerate(self.Sdata_txs.columns):
                tx_cap =  self.data["lvtx"].loc[i, "kvas_primary"]
                tx_type = self.data["lvtx"].loc[i, "Type"]
                if tx_type == "RES":
                    axes.plot(idx_date, 100*self.Sdata_txs[column]/tx_cap, color="silver", zorder=1)
                else:
                    axes.plot(idx_date, 100*self.Sdata_txs[column]/tx_cap, color="black", zorder=2)
            axes.xaxis.set_major_locator(hours)
            axes.xaxis.set_major_formatter(h_fmt)
            axes.set_xlabel('Time of the day')
            axes.set_ylabel('Transformer utilisation level [%]')
            axes.axhline(y=100, color='red', linestyle='--')
            axes.grid(color='grey', linestyle='-', linewidth=0.2)
            
            # Create a legend
            legend_elements = [
                Line2D([0], [0], color='blue', lw=3.5, label='MV Tx'),
                Line2D([0], [0], color='silver', lw=2, label='LV Txs, RES'),
                Line2D([0], [0], color='black', lw=2, label='LV Txs, COM'),
                Line2D([0], [0], color='red', linestyle='--', lw=2, label='Max allowed Tx %'),
            ]
            
            axes.legend(handles=legend_elements)
            
            # Adjust layout and show the plot
            plt.tight_layout()
            plt.show()
        
        else:
            
            fig, axes = plt.subplots(dpi=300, figsize=(16, 6))  # Adjust figsize as needed
            
            # MV transformer
            tx_cap = self.data["mv_net_txs"].loc[0, "kvas_primary"]
            axes.plot(idx_date, 100*self.Sdata["Value"]/tx_cap, color="blue", zorder=4, linewidth=3.5)
            
            # MV/LV transformers
            for i, column in enumerate(self.Sdata_txs.columns):
                tx_cap =  self.data["lvtx"].loc[i, "kvas_primary"]
                tx_type = self.data["lvtx"].loc[i, "Type"]
                tx_conn = self.data["lvtx"].loc[i, "Conn_Type"]
                
                if tx_type == "RES":
                    if len(tx_conn) == 3:
                        axes.plot(idx_date, 100*self.Sdata_txs[column]/tx_cap, color="darkgrey", zorder=1)
                    else:
                        axes.plot(idx_date, 100*self.Sdata_txs[column]/tx_cap, color="sandybrown", zorder=2)
                else:
                    axes.plot(idx_date, 100*self.Sdata_txs[column]/tx_cap, color="black", zorder=3)
            
            axes.xaxis.set_major_locator(hours)
            axes.xaxis.set_major_formatter(h_fmt)
            axes.set_xlabel('Time of the day')
            axes.set_ylabel('Transformer utilisation level [%]')
            axes.axhline(y=100, color='red', linestyle='--')
            axes.grid(color='grey', linestyle='-', linewidth=0.2)
            
            # Create a legend
            legend_elements = [
                Line2D([0], [0], color='blue', lw=3.5, label='MV Tx'),
                Line2D([0], [0], color='silver', lw=2, label='LV Txs, RES - 3ph'),
                Line2D([0], [0], color='sandybrown', lw=2, label='LV Txs, RES - SWER'),
                Line2D([0], [0], color='black', lw=2, label='LV Txs, COM - 3ph'),
                Line2D([0], [0], color='red', linestyle='--', lw=2, label='Max allowed Tx %'),
            ]
            
            axes.legend(handles=legend_elements)
            
            # Adjust layout and show the plot
            plt.tight_layout()
            plt.show()
            
    
    def gis_tx_utilisation(self):
            
        utilisation_pct = 100*self.gis_data["MVLV_txs"]["SNAP_val"]/self.gis_data["MVLV_txs"]["kvas_primary"]
        self.gis_data["MVLV_txs"]['Ut_pct'] = utilisation_pct

        # Plot
        fig, ax = plt.subplots(figsize=(8, 8), dpi=300)
        
        # Plot the GeoDataFrames
        self.gis_data["MV_tx"].plot(ax=ax, color="black", alpha=0.7, marker="^", markersize=256, label='MV tx', zorder=1)
        self.gis_data["MV_lines"].plot(ax=ax, color="grey", alpha=0.7, label='MV Lines', zorder=2, linewidth=2.5)
        self.gis_data["MVLV_txs"].plot(column='Ut_pct', ax=ax, cmap='RdYlGn_r', markersize=16, vmin=0, vmax=100)
        
        # Create color bar using a color map
        mappable = plt.cm.ScalarMappable(cmap='RdYlGn_r')
        mappable.set_array(range(101))
        cbar = fig.colorbar(mappable, ax=ax, shrink=0.5)  # Adjust 'shrink' as needed
        cbar.set_label('Tx Utilisation [%]', fontsize=10)
        cbar.set_ticks([0, 20, 40, 60, 80, 100])
        cbar.ax.set_yticklabels(['0', '20', '40', '60', '80', '100'], fontsize=10)
        
        for label in (ax.get_xticklabels() + ax.get_yticklabels()):
            label.set_fontsize(10)

        ax.title.set_fontsize(10)
        ax.xaxis.label.set_fontsize(10)
        ax.yaxis.label.set_fontsize(10)

        ax.set_xticks([])
        ax.set_yticks([])
        plt.tight_layout()
        plt.show()
    
    def gis_lines_utilisation(self):
        
        utilisation_pct = 100*self.gis_data["MV_lines"]["SNAP_val"]/self.gis_data["MV_lines"]["Ampacity"]
        self.gis_data["MV_lines"]['Ut_pct'] = utilisation_pct
        # Create a figure and axis
        fig, ax = plt.subplots(figsize=(8, 8), dpi=300)
        
        # Plot the GeoDataFrames
        self.gis_data["MV_tx"].plot(ax=ax, color="black", alpha=0.7, marker="^", markersize=256, label='MV tx', zorder=1)
        self.gis_data["MVLV_txs"].plot(ax=ax, color="grey", alpha=1, marker="o", markersize=16, label='LV txs', zorder=3)
        self.gis_data["MV_lines"].plot(column='Ut_pct', ax=ax, cmap='RdYlGn_r', linewidth=2.5, markersize=64, vmin=0, vmax=100, zorder=3)
        
        # Create color bar using a color map
        mappable = plt.cm.ScalarMappable(cmap='RdYlGn_r')
        mappable.set_array(range(101))
        cbar = fig.colorbar(mappable, ax=ax, shrink=0.5)  # Adjust 'shrink' as needed
        cbar.set_label('Line Utilisation [%]', fontsize=10)
        cbar.set_ticks([0, 20, 40, 60, 80, 100])
        cbar.ax.set_yticklabels(['0', '20', '40', '60', '80', '100'], fontsize=10)
        
        for label in (ax.get_xticklabels() + ax.get_yticklabels()):
            label.set_fontsize(10)

        ax.title.set_fontsize(10)
        ax.xaxis.label.set_fontsize(10)
        ax.yaxis.label.set_fontsize(10)
        
        ax.set_xticks([])
        ax.set_yticks([])
        plt.tight_layout()
        plt.show()

    def gis_tx_utilisation_daily(self):

        utilisation_pct = 100*self.gis_data["MVLV_txs"]["DAILY_max"]/self.gis_data["MVLV_txs"]["kvas_primary"]
        self.gis_data["MVLV_txs"]['Ut_max'] = utilisation_pct
        # Create a figure and axis
        fig, ax = plt.subplots(figsize=(8, 8), dpi=300)

        # Plot the GeoDataFrames
        self.gis_data["MV_tx"].plot(ax=ax, color="black", alpha=0.7, marker="^", markersize=256, label='MV tx', zorder=1)
        self.gis_data["MV_lines"].plot(ax=ax, color="grey", alpha=0.7, label='MV Lines', zorder=2, linewidth=2.5)
        self.gis_data["MVLV_txs"].plot(column='Ut_max', ax=ax, cmap='RdYlGn_r', markersize=16, zorder=3, vmin=0, vmax=100)

        # Create color bar using a color map
        mappable = plt.cm.ScalarMappable(cmap='RdYlGn_r')
        mappable.set_array(range(101))
        cbar = fig.colorbar(mappable, ax=ax, shrink=0.5)  # Adjust 'shrink' as needed
        cbar.set_label('Max Tx Utilisation [%]', fontsize=10)
        cbar.set_ticks([0, 20, 40, 60, 80, 100])
        cbar.ax.set_yticklabels(['0', '20', '40', '60', '80', '100'], fontsize=10)
        
        for label in (ax.get_xticklabels() + ax.get_yticklabels()):
            label.set_fontsize(10)

        ax.title.set_fontsize(10)
        ax.xaxis.label.set_fontsize(10)
        ax.yaxis.label.set_fontsize(10)
        
        ax.set_xticks([])
        ax.set_yticks([])

        plt.show()

    def gis_lines_utilisation_daily(self):

        utilisation_pct = 100*self.gis_data["MV_lines"]["DAILY_max"]/self.gis_data["MV_lines"]["Ampacity"]
        self.gis_data["MV_lines"]['Ut_max'] = utilisation_pct
        # Create a figure and axis
        fig, ax = plt.subplots(figsize=(8, 8), dpi=300)

        # Plot the GeoDataFrames
        self.gis_data["MV_tx"].plot(ax=ax, color="black", alpha=0.7, marker="^", markersize=256, label='MV tx', zorder=1)
        self.gis_data["MVLV_txs"].plot(ax=ax, color="grey", alpha=1, marker="o", markersize=16, label='LV txs', zorder=3)
        self.gis_data["MV_lines"].plot(column='Ut_max', ax=ax, cmap='RdYlGn_r', linewidth=2.5, markersize=64, vmin=0, vmax=100, zorder=3)

        # Create color bar using a color map
        mappable = plt.cm.ScalarMappable(cmap='RdYlGn_r')
        mappable.set_array(range(101))
        cbar = fig.colorbar(mappable, ax=ax, shrink=0.5)  # Adjust 'shrink' as needed
        cbar.set_label('Max Line Utilisation [%]', fontsize=10)
        cbar.set_ticks([0, 20, 40, 60, 80, 100])
        cbar.ax.set_yticklabels(['0', '20', '40', '60', '80', '100'], fontsize=10)
        
        for label in (ax.get_xticklabels() + ax.get_yticklabels()):
            label.set_fontsize(10)

        ax.title.set_fontsize(10)
        ax.xaxis.label.set_fontsize(10)
        ax.yaxis.label.set_fontsize(10)

        ax.set_xticks([])
        ax.set_yticks([])

        plt.show()
        

    
    

        
