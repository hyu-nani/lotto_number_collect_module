'''
복권 번호 수집기
author : hyu-nani
'''

import os
from openpyxl import load_workbook
import shutil
from datetime import datetime
import time

print()
print("\t┌───────────────────────────────────────────┐")
print("\t│                                           │")
print("\t│       복권 번호 확률정리기                │")
print("\t│       Version 1.2                         │")
print("\t│                             [ NANI ]      │")
print("\t└───────────────────────────────────────────┘\n")

file_list   =   os.listdir()

if len(file_list) == 0:
    print("파일이 없습니다.\n")
else:
    print("\t-- file list --")
    for i in range(len(file_list)):
        print("\t",end='')
        print(i,end=' ')
        print(file_list[i])
    print("\t---------------")
selectNum = int(input("\t 파일 선택 : "))

print("===================================================")
xlFileName = file_list[selectNum]
print(xlFileName)
book = load_workbook(str(xlFileName))
sheet = book['excel']

roundOfEvent = int(sheet['B4'].value)+1

print("숫자 추출중",end='')

#정렬을 위한 단순 리스트
numberSortList = []
#1~45 번의 각 자리의 나온 횟수 저장 리스트
numberCameOut = []
numberCameOutSort = []

#최근 나온 회차 저장 리스트
recentList1 = []
recentList2 = []
recentList3 = []
recentList4 = []
recentList5 = []

#1~45 번의 각 중복 발생 수
duplicateNumber = []
duplicateNumberSort = []
for i in range(45):
    numberSortList.append(i+1)
    numberCameOut.append(0)
    numberCameOutSort.append(0)
    duplicateNumber.append(0)
    duplicateNumberSort.append(0)
# 비교할 횟수 
compareNum = 2
countList = []
preNumList = []
for i in range(compareNum*7):
    preNumList.append(0)
for i in range(45):
    countList.append(0)

for i in range(roundOfEvent-1,1,-1):
    nowNumList = []
    num = int(sheet['N'+str(i+3)].value)
    nowNumList.append(num)
    preNumList[0] = num
    num = int(sheet['O'+str(i+3)].value)
    nowNumList.append(num)
    preNumList[1] = num
    num = int(sheet['P'+str(i+3)].value)
    nowNumList.append(num)
    preNumList[2] = num
    num = int(sheet['Q'+str(i+3)].value)
    nowNumList.append(num)
    preNumList[3] = num
    num = int(sheet['R'+str(i+3)].value)
    nowNumList.append(num)
    preNumList[4] = num
    num = int(sheet['S'+str(i+3)].value)
    nowNumList.append(num)
    preNumList[5] = num
    num = int(sheet['T'+str(i+3)].value)
    nowNumList.append(num)
    preNumList[6] = num

    if i > roundOfEvent-5:
        recentList1.append(nowNumList)
    # 각 번호 나온 횟수 추가
    for j in range(7):
        for k in range(1,46):
            if nowNumList[j] == k:
                numberCameOut[k-1] += 1
    # 기록중 같은 번호 찾기
    count = 0
    for j in range(1,46):
         if compareNum == preNumList.count(j):
            duplicateNumber[j-1] += 1
            count += 1
    if count > 0:
        countList[count-1] += 1
    #중복 기록 찾기 위한 리스트 쉬프트
    for j in range(compareNum-1,0,-1):
        for k in range(7):
            preNumList[j*7+k] = preNumList[(j-1)*7+k]
    if i % 100 == 0:
        print(".",end='')
print()
    

print("\t==== 전체 나온 횟수 ====")
for i in range(45):
    print('\t',end='')
    if i < 9:
        print(i+1,end=' ')
    else:
        print(i+1,end='')
    print("숫자 ",end='가 나올확률은')
    print(round((numberCameOut[i]/roundOfEvent)*100.0,3),end='%이고 연속으로 나온횟수는')
    print(duplicateNumber[i],end='회 이고\n')
print("\t이라면 1~45의 숫자 중 6개의 중복되지 않는 번호를 뽑아 나열한다면 다음번에 가장 나올 확률이 높은 숫자로 이루어진 숫자의 나열을 10가지 알려줘")
print("\t단, 10가지 전체 나열된 숫자들이 중복된 횟수는 3번 이하여야해.")

print("중복 갯수 계산..............")
print("\t2회차 연속 중복 수")
for i in range(45):
    if countList[i] != 0:
        print("\t",end='')
        print(i+1,end='개 : ')
        print(countList[i],end='번 \n')
print("\t합  : ",end='')
print(sum(countList),end='')
print("번")

print("정렬중",end='')
count = 0
for k in range(45):
    MAXNUM = 0
    for i in range(45):
        if numberCameOut[i] > MAXNUM:
            MAXNUM = numberCameOut[i]
    if MAXNUM > 0:
        for j in range(45):
            if MAXNUM == numberCameOut[j]:
                numberSortList[count] = j+1
                numberCameOutSort[count] = numberCameOut[j]
                duplicateNumberSort[count] = duplicateNumber[j]
                numberCameOut[j] = 0
                count += 1
                print(".",end='')
print()
print("\t==== 나온 횟수 정렬 ====")
for i in range(45):
    print('\t',end='')
    if numberSortList[i] < 10:
        print(numberSortList[i],end='  : ')
    else:
        print(numberSortList[i],end=' : ')
    print(numberCameOutSort[i],end=' 회 / 중복 : ')
    print(duplicateNumberSort[i],end='회 \n')
for i in 4:
    print(recentList1[i][])
