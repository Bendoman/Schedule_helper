names = []
import csv
csv_file = open('src/names.csv', 'r')
csv_reader = csv.reader(csv_file)
for row in csv_reader:
    names.append(row[0])
names = names[1:]
print(names)

csv_file.close()
csv_file = open('src/names.csv', 'a', newline='\n')
writer = csv.writer(csv_file, delimiter='\n')
value = 'Time Lightowler'
writer.writerow([value])