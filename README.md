## RUIDsExcelMerger
Are you a believer/worker in one of the temples of systematised slavery called SSC? Are you fed up with shitty excel copypasting tasks? Would you like to copy excel rows to the best place where they can be copied, namely to an other excel file?
Yes, of course 😊 
Here is a python script, which hopefully makes your life a bit easier. 

This script will iterate all the rows of the first excel sheets of all xlsx files in the folder of py file. First, it checks the specified column (KeywordIndex). If it finds any of your keywords (Keyword) AND finds a value in the specified column (MinimumIndex) which is not a None and higher than the specified value (Minimum)  then it will copy this row cells (between A:CA) and append it to the file called output.xlsx.  

# How it works: 
Assuming that you are working with tables containing some identificator strings (text) and values, it needs 4 inputs. You could provide these in a command line, after you start the py file. 
The system will iterate all the rows of the first excel sheets of all xlsx files in the py’s folder . First, it checks the specified column (KeywordIndex). If it finds any of your keywords (Keyword) AND finds a value in the specified column (MinimumIndex) which is not a None and higher than the specified value (Minimum)  then it will copy this row cells (between A:CA) and append it to the file called output.xlsx.  
Inputs: 
-	Keyword(s) – string (text)
-	KeywordIndex  - integer (number) 
-	Minimum – integer  (number)
-	MinimumIndex-integer  (number)

Keyword(s) – It could be anything, the system will consider them as a string. You can provide more than one. Those needs to be separated with commas. Mind that, the system is case sensitive and consider spaces as characters. 
KeywordIndex  - An index number for the column of the specified keyword’s row. Mind that, indexing starts from zero, therefore 0 = column A  
Minimum – It should be a number. If the script finds a row with higher value than this number in the specified cell by Minimumindex AND containing the keyword in the specified cell by KeywordIndex , then it will copy the values
MinimumIndex - An index number for the column of the specified minimum’s row. Mind that, indexing starts from zero, therefore 0 = column A  
 
#An example: 

After you copied the files, started the exe, you need to provide te following inputs:

Enter the examined keywords, separate by commas : Foo   
Enter the examined column index, which should be containing the keyword(s): 2
Enter the minimum value: 12
Enter the examined column index, which should be greather than the minimum: 10

The system will search for the „Foo” word (Keyword) in the 3rd column of the table, as you provide 2 as KeyboardIndex (COUNTING STARTS FROM ZERO!). If it founds a row with the specifics above, it will check the 11th column of the row, since you provided 10 as MinimumIndex. In case if this cell contains a value which is higher than 12 (Minimum), then it will copy the row cells (A:CA) and append it to the output.xlsx file. The script will iterate this process in all of the xlsx files in the directory where you placed the exe file. 

The system uses Python’s openpyxl library,
Output.xlsx needs to be empty every time before you starts the process. 
The current script copies only values, not formulas nor charts or pictures. 

You can use this as a standalone exe with no Python, just unzip the zip file and read the readme file. 

