import pandas
import os
import time
from pprint import pprint

def addRow(dictArr, preCount, curRow):
  dictRow = curRow
  dictRow.append(preCount)
  dictRow.append(curRow[5])

  if preCount == ">20":
    preCount = 21
  if curRow[5] == ">20":
    curRow[5] = 21

  diff = int(preCount) - int(curRow[5])
  dictRow.append(int(diff))

  if int(preCount) > 20 or int(curRow[5]) > 20:
    dictRow.append('appr')

  dictArr.append(dictRow)
  return dictArr

# input
print('rows to output: ')
outputRowsCount = input()
outputRowsCount = int(outputRowsCount)
print('rows selected: ' + str(outputRowsCount))
print()

# File 1
full_start_time = time.time()
start_time = time.time()
print('Opening first file...')

df1 = pandas.read_excel('1.xlsx')
preArr = df1.values.tolist()
print(f"Time spent: {time.time() - start_time:.4f} seconds")

# File 2
start_time = time.time()
print('Opening second file...')

df2 = pandas.read_excel('2.xlsx')
curArr = df2.values.tolist()
print(f"Time spent: {time.time() - start_time:.4f} seconds")


# Calculate
start_time = time.time()
print('Calculating...')

dictArr = []
length = len(preArr)

# Map
mappedPreArr = {}
for i, preRow in enumerate(preArr):
  print('mapping left:'  + str(length - i))
  mappedPreArr[preRow[1]] = preRow[5]

# Calculate
length = len(curArr)
for i, curRow in enumerate(curArr):
  if mappedPreArr.get(curRow[1]) == None:
    continue
  dictArr = addRow(dictArr, mappedPreArr.get(curRow[1]), curRow)
  print('calculate left:'  + str(length - i))

print('-----')
print(f"time spent calculating: {time.time() - start_time:.4f} seconds")

start_time = time.time()
print('-----')
print(f"sorting...")

start_time = time.time()

# Sort
def getDiff(e):
  return e[9]
dictArr.sort(reverse=True, key=getDiff)

print('-----')
print(f"time spent sorting: {time.time() - start_time:.4f} seconds")

# File write
start_time = time.time()
print('-----')
print(f"file write: " + str(len(dictArr)) + " ...")

dictArr = dictArr[:outputRowsCount]

df = pandas.DataFrame(dictArr)
excel_file = "output_" + str(time.time()) + ".xlsx"
df.to_excel(excel_file, index=False)

print(f"time spent file write: {time.time() - start_time:.4f} seconds")
print(f"Full time spend: {time.time() - full_start_time:.4f} seconds")

while 1:
  time.sleep(60)

