#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders


# In[2]:

def main():
    df = pd.read_excel("input_folder/import.xlsx",1)


    # In[3]:


    def mail_sender(error_type):
        sender_address = 'propriedades.figueiredo@gmail.com'
        receiver_address = "leo.defig@gmail.com"
        sender_pass = "Celia4040"

        #Setup the MIME
        message = MIMEMultipart()
        message['From'] = sender_address
        message['To'] = receiver_address
        message['Subject'] = 'Stock Import Notification'   #The subject line

        #Logic based on error
        if error_type == "header":
            mail_content = '''The following required columns are missing or misspelled: ''' + listToString(missing_columns)
        elif error_type == "nan":
            mail_content = '''The attached Excel contains columns with missing data, please change these columns in the original excel and try again: '''
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open("error_folder/missing_data.xlsx", "rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="missing_data.xlsx"')
            message.attach(part)


        #The body and the attachments for the mail
        message.attach(MIMEText(mail_content, 'plain'))
        #Create SMTP session for sending the mail
        session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
        session.starttls() #enable security
        session.login(sender_address, sender_pass) #login with mail_id and password
        text = message.as_string()
        session.sendmail(sender_address, receiver_address, text)
        session.quit()
        print('Mail Sent')


    # In[4]:


    def listToString(s):
        str1 = " "
        return (str1.join(s))


    # In[5]:


    missing_columns = []
    nan_df = df[df.isna().any(axis=1)]
    print(df.isna().any(axis=1))
    nan_df.to_excel("error_folder/missing_data.xlsx", index = False)
    print(nan_df.empty)

    def criteria_checker():

        required_columns = ["name","size","cat","qty","sku"]
        #missing_columns = []


        #Column header check
        for col in required_columns:

            if col not in df:
                missing_columns.append(col)

        if missing_columns:
            print(missing_columns)
            mail_sender("header")
            print("header issue")
            return False

        #Nan check
        if not nan_df.empty:
            mail_sender("nan")
            print("Missing Entry")
            return False

        #All conditions met return True
        if not missing_columns and nan_df.empty:
            return True


    # In[7]:


    if criteria_checker():

        mapper1 = {
        0:"unknown",
        1:"Services",
        2:"Inventory Goods",
        3:"Armchairs",
        4:"Sofas",
        5:"Bedroom",
        6:"Dining Room",
        7:"Coffee and Lamp Tables",
        8:"Consoles and TV Units",
        9:"Mirror Funiture",
        10:"Bar Stools and Ottomans",
        11:"Mirrors and Wall Art",
        12:"Lighting",
        13:"Clocks",
        14:"New Arrivals",
        15:"Glassware and Crystal",
        16:"Porcelain and Ceramics",
        17:"Planters and Greenery",
        18:"Boxes and Trays",
        19:"Ornaments and Objects",
        20:"Drinks Trolleys, Shelves and Plinths",
        21:"Specials",
        22:"Rugs/Cushion and Linen",
        23:"Outdoor Furniture"
        }

        mapper2 = {
            "Consoles and TV Units":"Furniture",
            "Ornaments and Objects":"Accessories",
            "Glassware and Crystal":"Accessories",
            "Drinks Trolleys, Shelves and Plinths":"Lounge",
            "Coffee and Lamp Tables":"Furniture",
            "Outdoor Furniture":"",
            "Planters and Greenery":"Outdoor",
            "Mirror Furniture":"Furniture",
            "Armchairs":"Lounge",
            "Bar Stools & Ottomans":"Lounge",
            "Dining Room":"Rooms",
            "Sofas":"Lounge",
            "Bedroom":"Rooms",
            "Porcelain and Ceramics":"Accessories",
            "Lighting":"Accessories",
            "Bar Stools and Ottomans":"Furniture",
            "Mirrors and Wall Art":"Wall Items",
            "Rugs/Cushions and Linen":"Accessories",
            "Boxes and Trays":"Accessories",
            "Clocks":"Wall Items",
        }

        df["ProductCode"] = df["sku"]
        df["Name"] = df["name"]
        df["Category"] = df["cat"]
        df = df.replace({"Category": mapper1})
        df["Brand"] = "Primavera"
        df["Type"] = "Stock"
        df["FixedAssetType"] = ""
        df["CostingMethod"] = "FIFO"
        df[['Length','Width','Height']] = df["size"].str.split(pd.read_excel("input_folder/import.xlsx",2).at[0,'Dimension Delimiter'],expand=True)
        df["Weight"] = ""
        df["WeightUnits"] = "kg"
        df["DimensionUnits"] = pd.read_excel("input_folder/import.xlsx",2).at[0,'Dimension Unit']
        df["Barcode"] = ""
        df["MinimumBeforeReorder"] = 0
        df["ReorderQuantity"] = 0
        df["DefaultLocation"] = "Primavera Location: Warehouse"
        df["LastSuppliedBy"] = ""
        df["SupplierProductCode"] = df["sc"]
        df["SupplierProductName"] = df["sn"]
        df["SupplierFixedPrice"] = 0
        df["PriceTier1"] = ""
        df["PriceTier2"] = ""
        df["PriceTier3"] = ""
        df["PriceTier4"] = ""
        df["PriceTier5"] = 0
        df["PriceTier6"] = 0
        df["PriceTier7"] = 0
        df["PriceTier8"] = 0
        df["PriceTier9"] = 0
        df["PriceTier10"] = 0
        df["AssemblyBOM"] = "No"
        df["AutoAssembly"] = "No"
        df["AutoDisassemble"] = "No"
        df["DropShip"] = "No Drop Ship"
        df["AverageCost"] = ""
        df["DefaultUnitOfMeasure"] = "Item"
        df["InventoryAccount"] = "Inventory Control"
        df["RevenueAccount"] = "Sales - " + df["Category"]
        df["ExpenseAccount"] = ""
        df["COGSAccount"] = "COS - " + df["Category"]
        df["AdditionalAttribute1"] = df["Category"]
        df = df.replace({"AdditionalAttribute1": mapper2})
        df["ProductAttributeSet"] = "Product Set " + df["AdditionalAttribute1"]
        df["AdditionalAttribute2"] = ""
        df["AdditionalAttribute3"] = ""
        df["AdditionalAttribute4"] = ""
        df["AdditionalAttribute5"] = ""
        df["AdditionalAttribute6"] = ""
        df["AdditionalAttribute7"] = ""
        df["AdditionalAttribute8"] = ""
        df["AdditionalAttribute9"] = ""
        df["AdditionalAttribute10"] = ""
        df["DiscountName"] = ""
        df["ProductFamilySKU"] = ""
        df["ProductFamilyName"] = ""
        df["ProductFamilyOption1Name"] = ""
        df["ProductFamilyOption1Value"] = ""
        df["ProductFamilyOption2Name"] = ""
        df["ProductFamilyOption2Value"] = ""
        df["ProductFamilyOption3Name"] = ""
        df["ProductFamilyOption3Value"] = ""
        df["CommaDelimitedTags"] = ""
        df["StockLocator"] = ""
        df["PurchaseTaxRule"] = "Products and Goods Imported"
        df["SaleTaxRule"] = "Standard Rate Sales"
        df["Status"] = "ACTIVE"
        df["Description"] = ""
        df["ShortDescription"] = ""
        df["Sellable"] = "Yes"
        df["PickZones"] = ""
        df["AlwaysShowQuantity"] = 0
        df["WarrantySetupName"] = ""
        df["InternalNote"] = ""
        df["ProductionBOM"] = "No"
        df["MakeToOrderBom"] = "No"
        df["QuantityToProduce"] = 1

        #Reordeing columns

        cols = df.columns.tolist()
        column_to_move = "Description"
        new_position = 78

        cols.insert(new_position, cols.pop(cols.index(column_to_move)))

        df = df[cols]

        column_to_move = "ProductAttributeSet"
        new_position = 53

        cols.insert(new_position, cols.pop(cols.index(column_to_move)))

        df = df[cols]


        print("wrote to output")
        df = df.loc[:, 'ProductCode':]

        #Writing to excel and csv
        df.to_excel("output_folder/output.xlsx", index = False)
        df.to_csv("output_folder/output.csv", index = False)
