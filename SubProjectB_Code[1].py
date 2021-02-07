# -*- coding: utf-8 -*-
"""
Created on Tue Nov  3 08:52:47 2020

@author: carso
"""
#Project B

#Delivery 1
#start by importing pandas and reading the excel file
import pandas as pd
df = pd.read_excel("Cases.xlsx")

#collecting 100 random rows as samples
data = df.sample(n=100)
data.to_excel("Data.xlsx")
#reopen the random data to sort it by date
Df = pd.read_excel("Data.xlsx")
#make ProcDate into Date-Time format that pandas can read
pd.to_datetime(Df["ProcDate"])
#finally we sort by dates in the column "ProcDate"
Df = Df.sort_values(by='ProcDate')
#if needed, print to check that the dates are in order
#print(Df['ProcDate'])
#export the data to the Excel file for use in the project
Df.to_excel("Data.xlsx")

#Delivery 2
import pandas as pd
#Use excel writer to write data into a new excel file called Cases2
writer = pd.ExcelWriter("Cases2.xlsx", engine = "xlsxwriter")
#Main function sends inputs to the other functions and prints results
def main():
    read_file = read("Cases.xlsx")
    print(read_file)
    diagnosis_unique = diagnosis(read_file["Diagnosis"])
    print("***UNIQUE DIAGNOSIS AND FREQUENCY***")
    print(diagnosis_unique)
    procedures_unique = procedures(read_file["Procedure"])
    print("***UNIQUE PROCEDURES AND FREQUENCY***")
    print(procedures_unique)
    dia = input("Enter a diagnosis: ")
    pro = input("Enter a procedure: ")
    f = frequency(dia, pro, read_file)
    print(pro,"has been ordered",f,"times for a diagnosis that includes",dia,".")
#1 - Reads a file and creates a dataframe for it
def read(file):
    df = pd.read_excel(file)
    return df

#2 - Creates a dictionary of unique diagnosis and their frequencies
def diagnosis(clm):
    unique = {} #Creates an empty dictionary called unique
    for d in clm:
        d_split = d.split(",") #Seperates multiple diagnosis by comma
        for e in d_split:
            e = e.strip() #Gets rid of unwanted space
            if e not in unique: 
                unique[e] = 1  #Adds new diagnosis as a key and initializes frequency value as 1
            else:
                unique[e] = unique[e] + 1  #If a diagnosis is already in the dictionary, adds 1 to frequency
    Data = pd.DataFrame.from_dict(unique, orient = "index", columns = ["Frequency"])  #Converts the dictionary into a 2 column dataframe
    Data = Data.sort_values("Frequency", ascending = False) #Sorts the data by frequency
    Data.to_excel(writer, sheet_name = "diagnosis")  #Writes the data into an excel sheet in Cases2
    return unique

#3 - Does the same as #2 but for procedures
def procedures(clm2):
    unique2 = {} 
    for d in clm2:
        d_split = d.split(",")
        for e in d_split:
            e = e.strip()
            if e not in unique2:
                unique2[e] = 1
            else:
                unique2[e] = unique2[e] + 1
    Data2 = pd.DataFrame.from_dict(unique2, orient = "index", columns = ["Frequency"])
    Data2 = Data2.sort_values("Frequency", ascending = False)
    Data2.to_excel(writer, sheet_name = "procedures")
    return unique2

#4 - Takes user input diagnosis and procedure and gives how many times that procedure was used for that diagnosis
def frequency(dia, pro, df):
    count = 0  #Initialize count
    for i in range(df["Diagnosis"].count()):  #For i in range of # of rows
        d = df.iloc[i, 0]  #The location of a diagnosis is row i in the diagnosis column (column 0)
        d = d.split(",") #Seperates multiple diagnosis by comma
        for e in d:
            if e == dia: #If the diagnosis == the user input diagnosis
                p = df.iloc[i, 1] #The location of a procedure is row i in the procedure column (column 1)
                p = p.split(",") #Seperates multiple procedures by comma
                for c in p:
                    if c == pro:
                        count = count + 1 #If the procedure == the user input produre, adds 1 to count
    return count

main() #Closes main function
writer.save() #Saves the Cases2 excel file that was written