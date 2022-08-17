import csv

with open('./data.csv') as f:
  reader = csv.reader(f)
  data_set = [row for row in reader]

print(data_set)