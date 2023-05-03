#import libraries
import os
import csv

#Set input path
poll_path = os.path.join("Resources","election_data.csv")

#Open as file object
with open(poll_path,'r') as poll_file:
#Store the header row
    poll_header =  next(poll_file)
#Create a file reader object in csv format
    poll_rows = csv.reader(poll_file, delimiter=",")
#Initialize the vote counts and candidate dictionary
    vote_cnt=0
    can = ""
    can_v = 1
    candidates = {}

#Loop through rows in the file
    for rows in poll_rows:
#Each row is couted as a vote
        vote_cnt = vote_cnt+1 
#When candidate name is not the same as row and also not present in dictionary - initialize candidate name and candidate vote variables, add both to candidate dictionary
        if can != rows[2] and candidates.get(rows[2],"Null") == "Null":
            can = rows[2]
            can_v =1
            candidates[rows[2]]=can_v
#When candidate name is not the same as row and but IS present in dictionary - initialize candidate name and add to candidate vote variable, add both to candidate dictionary
        elif can != rows[2] and candidates.get(rows[2],"Null") != "Null":
            can = rows[2]
            can_v = candidates.get(rows[2])+1
            candidates[rows[2]]=can_v
#When candidate name is the same as row (has lopped to next row within same candidate) - add to candidate vote variable, add to candidate dictionary
        else:
            can_v = candidates.get(rows[2])+1
            candidates[rows[2]]=can_v

#Print all required values             
print("Election Results\n-----------------------------")
print(f'Total Votes : {vote_cnt} \n -----------------------------')
#Loop to calculate the maximum votes candidate
max_val = 0
max_key =""
for key,value in candidates.items():
#Print key value pair from the candidate dicionary
        print(f'{key} : {round((int(value)/vote_cnt)*100,3)}% : ({value})')
        if value >max_val:
             max_val = value
             max_key = key
print(f'-----------------------------\nWinner is: {max_key}')
print(f'-----------------------------')

#Set output path
poll_op_path =os.path.join("Analysis","election_output.csv")
#Open path as wriatble file object
with open(poll_op_path,'w') as poll_output:
#Create writer object in csv format
    poll_writerows = csv.writer(poll_output, delimiter=",")
    poll_writerows.writerow([f'Election Results'])
    poll_writerows.writerow([f'-----------------------------'])
    poll_writerows.writerow([f'Total Votes : {vote_cnt}'])  
    poll_writerows.writerow([f'-----------------------------'])
#Print key value pair from the candidate dicionary
    for key,value in candidates.items():
        poll_writerows.writerow([f'{key} : {round((int(value)/vote_cnt)*100,3)}% : ({value})'])
    poll_writerows.writerow([f'-----------------------------'])
    poll_writerows.writerow([f'Winner is: {max_key}'])  
    poll_writerows.writerow([f'-----------------------------'])