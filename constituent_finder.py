'''
This Python script finds the Constituent ID using the Net IDs in the Excel Spreadsheet.
Make sure the Excel Spreadsheet is mapped correctly. G7 should be in NetID
'''

import openpyxl
from raiseredge import RaiserEdge
import pandas as pd

access = '' # Add your Primary Access Key
auth = '' # Add your OAUTH Key
re = RaiserEdge(access_key=access, oauth=auth)

# Function that checks the NetID is valid
def valid_net_id(s:str):
    
    if(len(s) != 6 or not s[:2].isalpha() or not s[2:].isdigit()):
        return False
    return True

# Function that find the Constituent using NetID only
def find_constituent_id(net_id: str):

    if valid_net_id(str(net_id)) == False:
        return None

    search_net_id = {
        'search_text': net_id
    }
    constituent_netid = re.search_constituent(**search_net_id)
    
    if constituent_netid['count'] > 0:
        print(constituent_netid)
        return constituent_netid
    
    return None

if __name__ == "__main__":
    
    # Creates a new XLXS Notebook
    workbook = openpyxl.Workbook()

    workbook.active.title = "Known_Event_Constituent"
    KConstituent = workbook["Known_Event_Constituent"]
    KConstituent.append(["Name", "Net_ID", "Constituent_ID", "Event_Name", "Event_ID"])

    workbook.save(filename='Test2.xlsx')


    # Set the file path
    # path = "C:\\Users\\STSC\\Documents\\Programming\\Python\\Raiser's Edge NXT API Wrapper\\MasterEventList 20-21 (Cleaned).xlsx"
    path = "C:\\Users\\STSC\\Documents\\Programming\\Python\\Raiser's Edge NXT API Wrapper\\MasterEventList 20-21 (4).xlsx"


    # Load the Excel file
    excel_file = pd.ExcelFile(path)
    
    
    # Get the sheet names
    sheet_names = excel_file.sheet_names


    # event_id = ["FY21Alum_PGN1","FY21Alum_EBITOR","FY21Alum_PGNL1","FY21Alum_FL1","FY21Alum_CP1","FY21Alum_PGNTI","FY21Alum_PWE1","FY21Alum_NH","FY21Alum_PGNL2","FY21Alum_FLR","FY21Alum_FLIp","FY21Alum_CPV","FY21Alum_ToBaN","FY21Alum_PGNFoF","FY21Alum_FC","FY21Alum_FL2","FY21Alum_CPH","FY21Alum_EBITP","FY21Alum_EBITSoS","FY21Alum_PWE2","FY21Alum_PGNL3","FY21Alum_FL3","FY21Alum_BIapoce","FY21Alum_FC1","FY21Alum_CUAaaA","FY21Alum_PGN2","FY21Alum_ADoS","FY21Alum_Em","FY21Alum_FC2","FY21Alum_CHtBC","FY21Alum_WotRB","FY21Alum_PGNL4","FY21Alum_FL4","FY21Alum_CTO","FY21Alum_FC3","FY21Alum_WotRLT1","FY21Alum_PGNSM","FY21Alum_FC4","FY21Alum_C1","FY21Alum_EBIT1","FY21Alum_WotRBYL","FY21Alum_PGNL5","FY21Alum_FL5","FY21Alum_FC5","FY21Alum_WotRLT2","FY21Alum_EB5","FY21Alum_PGN78a9M","FY21Alum_EBIT2","FY21Alum_WotR","FY21Alum_PGNL6","FY21Alum_FL6","FY21Alum_WotRLT3","FY21Alum_FL7","FY21Alum_WotRLT4","FY21Alum_PGNL7","FY21Alum_WotRLT5","FY21Alum_AWP","FY21Alum_PGNH","FY21Alum_PGNT&G","FY21Alum_FPWP&S","FY21Alum_C5","FY21Alum_FPWPT","FY21Alum_GG","FY21Alum_IIOF","FY21Alum_4","FY21Alum_AtF","FY21Alum_CP2","FY21Alum_FPWSJitA","FY21Alum_CME","FY21Alum_SAN1","FY21Alum_DSPP","FY21Alum_ADD","FY21Alum_IT","FY21Alum_PGN3","FY21Alum_CP3","FY21Alum_CP4","FY21Alum_CP5","FY21Alum_PGN","FY21Alum_WotRLT6","FY21Alum_PGN4","FY21Alum_CNJ","FY21Alum_PGNIStoT","FY21Alum_PST","FY21Alum_SwDN","FY21Alum_PGNL8","FY21Alum_SAN2","FY21Alum_CVN1","FY21Alum_LAC1","FY21Alum_PGNMQ1","FY21Alum_LAEBAyOi","FY21Alum_PGNL9","FY21Alum_WotRLT7","FY21Alum_PGNL10","FY21Alum_PGNMQ2","FY21Alum_C2","FY21Alum_LAC2","FY21Alum_TNBA","FY21Alum_PGNL11","FY21Alum_C3","FY21Alum_WotRL","FY21Alum_TCoyB","FY21Alum_AGn","FY21Alum_C4","FY21Alum_PGNL12","FY21Alum_LAC3","FY21Alum_R","FY21Alum_PGN5","FY21Alum_C6","FY21Alum_WotRLT8","FY21Alum_EA"]
    event_id = [] # Add Event ID here in order of the worksheet

    counter = 0
    # Loop through each sheet and print the value of cell G7
    for sheet in sheet_names:
        id = event_id[counter]
        if sheet != 'Event_Date_Time_Details':
            # Read the sheet into a DataFrame
            df = pd.read_excel(path, sheet_name=sheet)
            
            # Print the value of cell G7
            print(f"Sheet: {sheet}, Value: {df.iloc[5, 6]}")

            column_name = df.columns[6]
            column_values = df[column_name]

            # Print the column values from row 8 until the last row
            for i in range(6, len(column_values)):
                net_id = column_values[i]
                constituent_data = find_constituent_id(net_id)
                if constituent_data != None:
                    try:
                        KConstituent.append([constituent_data['value'][0]['name'], net_id, constituent_data['value'][0]['lookup_id'], sheet , id])
                    except:
                        print("Can't Find Right Data")

            workbook.save(filename='Test2.xlsx')
            counter+=1

