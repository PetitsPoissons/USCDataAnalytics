#Import modules
import os
import csv

#Identify file (budget_data.csv is in the same PyBank directory as this code,
# so I don't need to set the path to retrieve this data file)
file = "budget_data.csv"

#Define variables
budget_data = []
change_pl = {}
total_months = 0
total_pl = 0
previous_month_pl = 0

#Define functions
def total_change(dict_name):
    sum_change_pl = 0.0
    for k in dict_name:
        sum_change_pl += dict_name[k]
    return(sum_change_pl)

def print_results(rows_list):
    print('\n')
    for row in rows_list:
        print(row)
    print('\n')

#Open the CSV
with open(file) as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')

    #Store header and move one row down
    csv_header = next(csvreader)

    #Loop through and store into list budget_data
    for row in csvreader:
        budget_data.append(row)

#Loop through budget_data list and perform various calculations:
for item in budget_data:

    #total number of months
    total_months += 1

    #total net P/L
    month_pl = int(item[1])
    total_pl += month_pl

    #month-to-month change in P/L
    if total_months > 1:
        #store month-to-month change in P/L into a dictionary
        change_pl[item[0]] = month_pl - previous_month_pl
    
    #reset value for previous_month_pl
    previous_month_pl = month_pl

#Calculate average change in P/L
avg_change_pl = round(total_change(change_pl) / (total_months - 1),2)

#Calculate greatest increase in profits and in losses
max_profit_chg = max(change_pl.items(), key=lambda item: item[1])
max_loss_chg = min(change_pl.items(), key=lambda item: item[1])

#Prepare presentation of results
row1 = 'Financial Analysis'
row2 = '--------------------------'
row3 = f"Total Months: {total_months}"
row4 = f'Total: ${total_pl}'
row5 = f'Average Change: ${avg_change_pl}'
row6 = f'Greatest Increase in Profits: {max_profit_chg[0]} (${max_profit_chg[1]})'
row7 = f'Greatest Decrease in Profits: {max_loss_chg[0]} (${max_loss_chg[1]})'
rows = [row1,row2,row3,row4,row5,row6,row7]

#Print results to the terminal
print_results(rows)

#Write results in text file
output_file = 'PyBank_results.txt'
with open(output_file,'w',encoding='utf8',newline='') as txtfile:
    txtfile.writelines(row + '\n' for row in rows)