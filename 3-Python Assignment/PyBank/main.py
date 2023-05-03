#import libraries
import os
import csv

#Set input path
budget_path = os.path.join("Resources","budget_data.csv")

#Open csv file
with open(budget_path,'r') as budget_file:
#store the header row
    budget_header =  next(budget_file)
#Read csv file
    budget_rows = csv.reader(budget_file, delimiter=",")
#Initialize month count,Profit/Loss Totals, List for Profit/Loss Changes month to month
    mnth_cnt=0
    PL_tot = 0
    PL_curr = 0
    PL_prev = 0 
    PL_chg = []
#Variables to store the greatest/lowest change and corresponding months
    GC = 0
    GCM = ""
    LC = 0
    LCM = ""

#Loop thorugh rows in file
    for rows in budget_rows:
#Every row gets counted as a month
        mnth_cnt = mnth_cnt+1
#Sum the Profit/Loss with every row
        PL_tot = PL_tot + int(rows[1])
#Calculate difference of current row PL from previous PL value
#Store change in list
        PL_curr = int(rows[1]) 
        PL_chg.append(PL_curr - PL_prev)
#Store current row in previous PL value
        PL_prev = int(rows[1]) 

        print(PL_chg)
        #Reassign the greatest and lowest change values in variables
        #Jan 10 change needs to be discarded. Start with 2nd row change value(index 1)
        if len(PL_chg)>1 :
            if PL_chg[mnth_cnt-1]>GC :
                GC =PL_chg[mnth_cnt-1]
                GCM = rows[0]
            elif PL_chg[mnth_cnt-1]<LC :
                LC =PL_chg[mnth_cnt-1]
                LCM = rows[0]
    
#Exclude first month change because previous month data does not exist
PL_chg.pop(0)

#Print all required values
print(f'Financial Analysis')
print("------------------")
print(f'Total Months: {mnth_cnt}')
print(f'Total Profit/Loss: ${PL_tot}')
print(f'Average Change in Profit/Loss: ${round(sum(PL_chg)/len(PL_chg),2)}')
print(f'Greatest Increase in Profits: {GCM} : $({GC})')
print(f'Greatest Increase in Profits: {LCM} : $({LC})')

#Create output path
budget_op_path =os.path.join("Analysis","budget_output.csv")
#Open output path as writable file object
with open(budget_op_path,'w') as budget_output:
#Create file writer object in csv format
    budget_writerows = csv.writer(budget_output, delimiter=",")
    budget_writerows.writerow(["Financial Analysis"])
    budget_writerows.writerow(["------------------"])
    budget_writerows.writerow([f'Total Months: {mnth_cnt}'])
    budget_writerows.writerow([f'Total Profit/Loss: ${PL_tot}'])
    budget_writerows.writerow([f'Average Change in Profit/Loss: ${round(sum(PL_chg)/len(PL_chg),2)}'])
    budget_writerows.writerow([f'Greatest Increase in Profits: {GCM} : $({GC})'])
    budget_writerows.writerow([f'Greatest Increase in Profits: {LCM} : $({LC})'])