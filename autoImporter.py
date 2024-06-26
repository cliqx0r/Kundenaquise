import json, xlwings, time, os
import pandas as pd
import xml.etree.ElementTree as ET
from sys import exit

def writeLog(errorlist):
        if len(errorlist) > 0:
            with open("logfile.txt", "w") as logfile:
                for error in errorlist:
                    logfile.write(error)
        if len(errorlist) > 0:
            for error in errorlist:
                print(error)
            input("press any Button to close")
        else:
            pass

def addInfo(value, errorList):
    errormessage = f"{time.strftime('%Y-%m-%d %H:%M')}_____{value}\n"
    errorList.append(errormessage)


def main(): 
    
    errors = []

    if os.path.exists("config.xml") == False:
        addInfo(f"No such file or directory: config.xml", errors)
        addInfo(f"Try to create file: config.xml", errors)
        configfile = ET.Element('config')
        serialnumbersElement = ET.SubElement(configfile, "filepath_serialnumbers")
        customerexcelElement = ET.SubElement(configfile, "filepath_customerexcel")
        serialnumbersElement.text = "serialnumbers.json"
        customerexcelElement.text = " "
        b_xml = ET.tostring(configfile)
        

        with open("config.xml", "wb") as config:
            config.write(b_xml)

        if os.path.exists("config.xml") == True:
            addInfo(f"Created File: config.xml :: please fill out the correct paths in config.xml and retry", errors)
        else:
            addInfo(f'File could not be created: "config.xml"', errors)

        
    else:
        global filepath_serialnumbers
        global filepath_customerexcel

        config = ET.parse("config.xml").getroot()

        filepath_serialnumbers = config[0].text
        if filepath_serialnumbers is None:
            addInfo('parameter "filepath_serialnumbers" has no value \n', errors)
        elif os.path.exists(filepath_serialnumbers) == False:
            addInfo(f"No such file or directory: {filepath_serialnumbers}", errors)
        else: 
            pass

        filepath_customerexcel = config[1].text
        if filepath_customerexcel == " ":
            addInfo('parameter "filepath_customerexcel" has no value \n', errors)
        elif os.path.exists(filepath_customerexcel) == False:
            addInfo(f"No such file or directory: {filepath_customerexcel}", errors)
        else:
            pass
    writeLog(errors)

    wb = xlwings.Book(filepath_customerexcel)
    sheet= wb.sheets[2]

    ## reading through the serialnumbers.json file
    with open(filepath_serialnumbers, 'r', encoding='utf-8') as sn:
        serialnumbers = json.load(sn)
        sn.close()

    dataframe = pd.read_excel(filepath_customerexcel, "Anfragen über TIBEK")

    for index, row in dataframe.iterrows():
        for rowrow in serialnumbers[2]["data"]:
            if row.iloc[13] == rowrow["serialnumber"] and pd.isna(row.iloc[17]):
                if rowrow["activationdate"] != None:
                    sheet.range((index+2,17)).value = rowrow["activationdate"][:-9]
                if rowrow["testenddate"] != None:
                    sheet.range((index+2,18)).value = rowrow["testenddate"][:-9]
    wb.save()
    wb.close()
    addInfo("List succesfully updated", errors)
    writeLog(errors)

if __name__ == "__main__":
    main()
