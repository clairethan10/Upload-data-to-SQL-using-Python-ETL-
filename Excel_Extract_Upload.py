# -*- coding: utf-8 -*-
"""
Created on Wed Feb 25 15:20:05 2026

@author: vhaglathanc

This file upload an excel file onto SQL for invoicing data.
"""

from pathlib import Path 
import pandas as pd 
from datetime import datetime 
from sqlalchemy import create_engine
import sqlalchemy
import logging 
import configparser
import shutil 
import psutil 
import os 
import win32com.client as win32 
import re 



# find excel files in the folder path 

def extract_excel(folder, outputfolder): 
    
    log.info("Extracting Excel STARTED")
    
    files = list(Path(folder).glob("*.xls*"))  # find both xls and xlsx files 
    
    log.info(f"Reading {files}")
    
    df_list=[] # create a list for all excel files 
    
    for f in files:
        df=pd.read_excel(f, header=2)
            
        df["filename"]=f.name # attaching the file name to the table  
        df["ImportDate"]=datetime.today().strftime('%Y-%m-%d') # get today date  
        
        df_list.append(df) # adding all excel files into the list  
        
    final_df=pd.concat(df_list, ignore_index=True) #stack all files into one dataframe
    
    try: # save a temp csv table in the load folder before uploading onto SQL 
        final_df.to_csv(outputfolder + "ES_Inv.csv", index=False, header=True, encoding="utf-8-sig")
    except Exception as e:
        print(f"Pandas saving to_csv error: {e}")
            
    log.info("Extracting Excel DONE")
    return final_df



def upload_csv(path, engine, if_exists):
    log.info("Uploading ES_Invoicing STARTED")
    
    # create a data dictionary for SQL table
    dt={
        "Account Number" : sqlalchemy.NVARCHAR(length=20),
        "Account Name"   : sqlalchemy.NVARCHAR(length=100), 
        "Invoice Period" : sqlalchemy.NVARCHAR(length=20),
        "Accession Number": sqlalchemy.NVARCHAR(length=20), 
        "Requisition Number": sqlalchemy.NVARCHAR(length=20), 
        "DOS" : sqlalchemy.DATE,
        "PatientName": sqlalchemy.NVARCHAR(length=100), 
        "Patient DOB": sqlalchemy.DATE,
        "Product": sqlalchemy.NVARCHAR(length=100),
        "MRN": sqlalchemy.NVARCHAR(length=20), 
        "PO#": sqlalchemy.NVARCHAR(length=20), 
        "Charges": sqlalchemy.NVARCHAR(length=100), 
        "ImportDate": sqlalchemy.DATE           
        }
    
    try:
        df=pd.read_csv(path + "ES_Inv.csv", encoding="utf-8-sig")
    except Exception as e: 
        log.exception(f"CSV reading exception raised: {e}")
    
    df.drop_duplicates() 
    
    try: 
        df.to_sql(
            "Invoicing",
            engine,
            schema="INV",
            index=False,
            if_exists=if_exists,
            dtype=dt,
        )
    except Exception as e:
        log.exception(f"Uploading CSV to SQL error: {e}")
        
    log.info("Uploading Invoicing DONE")
    return 

def is_outlook_running():
    for p in psutil.process_iter(attrs=["name", "pid"]):
        if "OUTLOOK.EXE" in p.info["name"]:
            return True
    return False 
    
    
               

def send_last_log_line(log_path, addresses, subject="Invoice Loading Log Update"): 
    
    
    try:
        #check if outlook is running 
        
        if not is_outlook_running():
            os.startfile("outlook")
            
        # get the latest line from the log file 
        
        logfiles=list(Path(log_path).glob("*.log"))
        
        if not logfiles:
            raise FileNotFoundError("No log files found.")
            
        dated_log=[]
        
        for f in logfiles:
            match = re.search(r"\d{4}-\d{2}-\d{2}".f.name) #extract dates from the file 
            if match:
                file_date=datetime.strptime(match.group(),"%Y-%m-%d")
                dated_log.append((file_date, f))
            
        
        if not dated_log: 
            raise ValueError("No log files with dates found.")
            
        latest_file = max(dated_log, key=lambda x: x[0])[1]
        

        with latest_file.open("r", encoding="utf-8") as f:
            lines=f.readlines()
            if not lines:
                last_line="Log file is empty"
            else: 
                last_line=lines[-1].strip() 
                
                
        # begin email message
        
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)      

        if isinstance(addresses, str):
            addresses=[a.strip() for a in re.split(r"[;,]", addresses) if a.strip()]
        mail.To=';'.join(addresses)
        
                       
        mail.subject=f"{subject} - {latest_file.name}"
        mail.Body = (
            f"Last log entry of invoice data upload:\n"
            f"Last line: {latest_file}\n\n"
            f"Last entry: {last_line}"
            )
        
        mail.Send()

    except RecursionError as e:
        log.debug(e) 
    except Exception as e:
        log.debug(e)    
             
    return     
if __name__=="__main__": 
    config=configparser.ConfigParser()
    config.read("config.ini")
    
    # get informations from the ini file 
    
    excel_path=config.get("path", "excel_path")
    outputtable_path=config.get("path", "output_path")
    archive = config.get("path", "archive")
    log_dir = config.get("path", "log_path")
    

    addresses = config.get("email", "addresses")
    
    if_exists=config.get("misc", "if_exists")
    
    
    # set up logging 
    today=datetime.today().strftime('%Y-%m-%d')
    
    log=logging.basicConfig(
        filename=f"{log_dir}\\invoicing_{today}.log", 
        filemode="a", #append the existing log file 
        level=logging.DEBUG, #add details information 
        format="%(levelname)s;%(asctime)s;%(filename)s;"
        + "%(funcName)s;%(lineno)s;%(message)s"
        )
    
    log=logging.getLogger(__name__) 
    
    log.info("Reading excel file started")
    
    # starts reading the files 
    
    ifiles=list(Path(excel_path).glob("*.xls*"))
    
    if len(ifiles)>0 :
        
        extract_excel(excel_path, outputtable_path)
        
        engine = create_engine(
                    "mssql+pyodbc:<--path-->?Trusted_Connection=yes&driver=OhDBC+Driver+17+for+SQL+Server"
                )
        
        upload_csv(outputtable_path, engine, if_exists)
        
        # move the file to archive folder 
        
        for f in ifiles: 
            try:
                shutil.move(f, archive)
                log.info(f"Excel {f} moved successfully into {archive}")
            except Exception as e: 
                log.info(f"{f} cannot be moved: {e}")
        
        log.info("Invoicing excel file successfully loaded")
        

    else:
        log.info(f"No excel file found in {excel_path}")
        
    send_last_log_line(log, addresses, subject="Invoice Loading Log Update")
        
