import pandas as pd

shipToFile = pd.read_excel(".\data\Ship-to_VBU-GBM_DueDate_2025-06-02.xlsx", sheet_name="JP Assignment List",header=2)
cols = [
    "ECC Ship-to",
    "Cust",
    "Cust_Name", 
    "Address_1",
    "City",
    "Country",
    "Sales_Coverage",
    "End_Mkt_Segment",
    # "PostalCode"
    ]
shipToFile[cols].to_csv(".\data\Ship-to_VBU-GBM_DueDate_2025-06-02_cleaned.csv", index=False)
# list1 = shipToFile["ECC Ship-to"].tolist()
# print(len(list1))

