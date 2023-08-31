# VBA-challenge
The VBA script for this file is based on the codes provided by the in-class assignments from the U of T Data Analysis bootcamp.  ChaGPT was also used in the assistance of debugging and refactoring the script.  

The purpose of this file is to loop through all the stocks for one year and output information for each ticker symbol.  This information included the "yearly change", :the percentage change", and "the total volume of the stock".  This summary of information is then used to provide the "greatest % increase", "greatest % decrease", and "greatest total volume".  Finally, conditional formatting is utilized to make it user friendly.

My key findings with this project included how it is necessary to refactor the vba script in order for the file to run smoothly.  This is necessary for large datasets to avoid any potential crashes.  Another finding is to ensure that any cells, columns or ranges being referenced in the vba script have "ws." in front of it, in order for the script to loop through each worksheet and fill in those values successfully.  Otherwise the program will not attempt to run at all.
