import os
from fnmatch import fnmatch
import csv
import re

root = "."
path = os.path.join(root, "\\attachments")
pattern = "*.csv"
filenames = []

for path, subdirs, files in os.walk(root):
   for name in files:
        if fnmatch(name, pattern):
            fn = path + "\\" + name
            filenames.append(fn)

fields = ['Source', 'Username', 'Masked Password']
merged_filename = "credparser.csv"

with open(merged_filename, 'w', encoding="UTF-8", newline='') as csvfile:
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(fields)
    for fn in filenames:
        f = open(fn, encoding="UTF-8")
        f.readline()
        csv_f = csv.reader(f)
        for row in csv_f:
            if re.search('\[HASH\]',str(row[0])) or re.findall(r"(^[a-fA-F\d]{32})", row[2]):
                row[2] = "<Password Hash>"
                row.pop
                row[0] = fn.split("\\")[2]
                csvwriter.writerow(row)
            elif not row[2] == 'NF' or row[2] == "EMPTY:NONE":
                row.pop
                row[0] = fn.split("\\")[2]
                strlength = len(row[2])
                masked = int(strlength/3)
                slimstr = row[2][masked:strlength-masked]
                row[2] = "*" * masked + slimstr + "*" * masked
                csvwriter.writerow(row)


