#Import modules
import os
import csv

#Identify file (election_data.csv is in the same PyPoll directory as this code,
# so I don't need to set the path to retrieve this data file)
file = "election_data.csv"

#Define variables
election_data = []
number_votes = []
percent_votes = []


#Define function
def vote_count(name):
    count = 0
    for row in election_data:
        if row[2] == name:
            count += 1
    return(count)

def output(command_name):
    command_name('\n' + title + '\n' + line + '\n' + total_votes + '\n' + line + '\n')
    for result in results:
        command_name(f'{result[0]}: {result[1]}% ({result[2]})\n')
    command_name(line + '\n' + winner + '\n' + line + '\n')

#Open the CSV
with open(file) as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')

    #Store header and move one row down
    csv_header = next(csvreader)

    #Loop through and store into list election_data
    for row in csvreader:
        election_data.append(row)

#Calculate total number of votes
total_votes = len(election_data)

#Create list of candidates
candidates = sorted(list(set([vote[2] for vote in election_data])))

#Calculate percentage votes for each candidate
for candidate in candidates:
    number_votes.append(vote_count(candidate))
    percent_votes.append(format(vote_count(candidate)/total_votes*100,'.3f'))

winner_index = number_votes.index(max(number_votes))
winner_name = candidates[winner_index]

#Prepare for presentation of results
title = 'Election Results'
line = '--------------------------'
total_votes = f"Total Votes: {total_votes}"
winner = f"Winner: {winner_name}"
results_zipped = zip(candidates, percent_votes, number_votes)
results = list(results_zipped)
results.sort(key=lambda x:x[2],reverse=True)
print(results)

#Print results to the terminal
output(print)

#Write results in text file
output_file = 'PyPoll_results.txt'
with open(output_file,'w',encoding='utf8',newline='') as txtfile:
    output(txtfile.writelines)