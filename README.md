# Healthcare_Dataset_-Analyzer

I have a real-life data set containing more than 1600 records. Each record contains partial information about a patient appointment at an electrophysiology lab at a hospital.  Note that all identifiable patient information are omitted.  At an electrophysiology lab, an electrophysiologist (a doctor who specializes in the diagnosis and treatment of abnormal heart rhythms) will numb the skin in patient’s groin (a few inches to the side of the genitalia) with medication. Then, he’ll insert several catheters, or tubes, into a vein. The catheters are threaded to patient’s heart, where they’ll sense the electrical activity there and evaluate the performance of the heart.

I wrote a Python program that reads the given data set and select a random sample of 100 records from this data set and store them in an Excel file. Next, I will use the data set to learn about the “nature of the business”, what happens in an EP lab as it relates to diagnosis and EP procedure.  

Then, The program includes a function that when it is called it reads the data from the Excel and stores it in a pandas dataframe.  It then slices the columns for the diagnoses and the procedures. 
The program includes a function to find the list of all unique diagnoses and writes the results in an Excel workbook.  The table is sorted by frequencies and the name of the sheet is “diagnosis”.  
The program includes a function to find the list of all unique procedures and writes the results in the same Excel workbook.  The table is sorted by frequencies and the name of the sheet is “procedures”.  
The program includes a function that asks a user to enter a diagnosis, e.g. “Afib” and a procedure, e.g. “DCCV”.   It then determines the number of appointments that this procedure was ordered for the entered diagnosis.  

