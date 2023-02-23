import csv
with open('src/names.csv', 'a') as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow([f'{1}'])