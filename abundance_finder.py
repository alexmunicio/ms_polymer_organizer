import os
import xlrd
import pandas as pd

# path = "C:\\Users\\munic\\OneDrive\\Documents\\Atmospheric_Chemistry_Research\\6_29_2022\\test.xlsx"
# path = "C:\\Users\\munic\\Projects\\Research\\Python Scripts\\Data\\test2.xls"
# path = "C:\\Users\\munic\\OneDrive\\Documents\\Atmospheric_Chemistry_Research\\6_29_2022\\test.xls"
# excel_workbook = xlrd.open_workbook(path)
# excel_worksheet = excel_workbook.sheet_by_index(0)

# project_dir = "C:\\Users\\munic\\Projects\\Research\\Python_Scripts\\experimental_software\\"
project_dir = "C:\\Users\\munic\\OneDrive\\Documents\\Atmospheric_Chemistry_Research\\01_20_2023\\"
project_dir = "C:\\Users\\munic\\Documents\\Improvised_Research\\"
experiment_file = "01202023_Glyoxal_Python_Workup.csv"
experiment_file = "01132023_redone_Python_Workup.csv"
#water_1_file = "CC_BS_Methylglyoxal0,0500M_08_01_18_Modified_Data_all_withtime_WH_02.csv"
#water_2_file = "CC_BS_Methylglyoxal0,0500M_08_01_18_Modified_Data_withtime_MG.csv"
#water_3_file = "6_29_2022_Water_Wash_3_Python_Work_Up.csv"
results_file = project_dir + "results.xlsx"

experiment_df = pd.read_csv(project_dir + experiment_file)
#water_1_df = pd.read_csv(project_dir + water_1_file)
#water_2_df = pd.read_csv(project_dir + water_2_file)
#water_3_df = pd.read_csv(project_dir + water_3_file)

if os.path.exists(results_file): 
    os.remove(results_file)

# C3H8O4_Na = (chicken["Molecular Formula"] == "C3H8O4") & (chicken["Adduct"] == "Na")
# C3H6O3_Na = 
# print(chicken[C3H8O4_Na])
# chicken[C3H8O4_Na].to_csv(project_dir + results_file)

# mgly adducts (comment out glyoxal if running for mgly)
compound_adducts = [
    ("C3H6O3", "Na"), 
    ("C3H8O4", "Na"), 
    ("C6H8O4", "H"),
    ("C6H10O5", "Na"),
    ("C6H12O6", "Na"),
    ("C9H12O6", "H"),
    ("C9H16O8", "Na"),
    ("C9H18O9", "Na"),
    ("C12H16O8", "H"),
    ("C12H22O11", "Na"),
    ("C15H20O10", "H"),
    ("C15H22O11", "H"),
    ("C15H26O13", "Na"),
    ("C15H28O14", "Na"),
    ("C18H24O12", "H"),
    ("C18H26O13", "H"),
    ("C18H30O15", "Na"),
    ("C18H32O16", "Na"),
    ("C21H28O14", "H"),
    ("C21H30O15", "H"),
    ("C21H34O17", "Na"),
    ("C21H36O18", "Na"),
    ("C24H34O17", "H"),
    ("C24H36O18", "H"),
    ("C24H38O19", "Na"),
    ("C24H40O20", "Na")]

# glyoxal adducts
#compound_adducts = [
#    ("C2H4O3", "H"),
#    ("C2H6O4", "Na"),
#    ("C4H8O6", "Na"),
#    ("C4H10O7", "Na"),
#    ("C6H10O8", "Na"),
#    ("C6H12O9", "Na"),
#    ("C6H14O10", "Na"),
#    ("C8H14O11", "Na"),
#    ("C8H16O12", "Na"),
#    ("C8H18O13", "Na"),
#    ("C10H18O14", "Na"),
#    ("C12H22O17", "Na")
#]
writer = pd.ExcelWriter(results_file, engine = "xlsxwriter")

def abundance_generator(compound, adduct):
    abundance_filter = (experiment_df["Molecular Formula"] == compound) & (experiment_df["Adduct"] == adduct) 
    print(experiment_df[abundance_filter])
    experiment_df[abundance_filter].to_excel(writer, sheet_name = compound, index = False)
    
    #water_1_filter = (water_1_df["Molecular Formula"] == compound) & (water_1_df["Adduct"] == adduct)
    #water_2_filter = (water_2_df["Molecular Formula"] == compound) & (water_2_df["Adduct"] == adduct)
    #water_3_filter = (water_3_df["Molecular Formula"] == compound) & (water_3_df["Adduct"] == adduct)

    last_row = len(experiment_df[abundance_filter].index) + 2

    #water_1_df[water_1_filter].to_excel(writer, sheet_name = compound, startrow = last_row, index = False)
    #last_row = last_row + len(water_1_df[abundance_filter].index) + 2
    #water_2_df[water_2_filter].to_excel(writer, sheet_name = compound, startrow = last_row, index = False)
    #last_row = last_row + len(water_2_df[abundance_filter].index) + 2
    #water_3_df[water_3_filter].to_excel(writer, sheet_name = compound, startrow = last_row, index = False)

for compound_adduct in compound_adducts:
    abundance_generator(compound_adduct[0], compound_adduct[1])

writer.save()