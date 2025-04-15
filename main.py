'''
복권 번호 추첨기
author : hyu-nani
가중치 : 거이 나올 확률 없음 -1- ~ 확실함 : +10
'''


#예시 B4     6 7 27 29 38 45 17


import os
from openpyxl import load_workbook
import shutil
from datetime import datetime
import time
import random
from collections import Counter
#5번 구간별 출현횟수 [0~9] 0:1~5, 1:6~10....

test = True

WinNum = "6 7 27 29 38 45 17"

if test == True:
    start = 'B5'
    testNum = 100000
else:
    start = 'B4'
    testNum = 1000

showNum5 = [0] * 9
def num5check(val):
    if val <= 5:
        showNum5[0] += 1
    elif val <= 10:
        showNum5[1] += 1
    elif val <= 15:
        showNum5[2] += 1
    elif val <= 20:
        showNum5[3] += 1
    elif val <= 25:
        showNum5[4] += 1
    elif val <= 30:
        showNum5[5] += 1
    elif val <= 35:
        showNum5[6] += 1
    elif val <= 40:
        showNum5[7] += 1
    else:
        showNum5[8] += 1

print()
print("\t┌───────────────────────────────────────────┐")
print("\t│                                           │")
print("\t│       복권 번호 확률정리기                │")
print("\t│       Version 1.3                         │")
print("\t│                             [ NANI ]      │")
print("\t└───────────────────────────────────────────┘\n")

file_list   =   os.listdir()

A = False
if len(file_list) == 0:
    print("파일이 없습니다.\n")
else:
    print("\t-- file list --")
    for i in range(len(file_list)):
        print("\t",end='')
        print(i,end=' ')
        print(file_list[i])
        if file_list[i] == 'data.xlsx':
            A = True
            selectNum = i
    print("\t---------------")
if A == False:
    selectNum = int(input("\t 파일 선택 : "))

print("===================================================")
xlFileName = file_list[selectNum]
print(xlFileName)
book = load_workbook(str(xlFileName))
sheet = book['excel']
roundOfEvent = int(sheet[start].value)+1

print("숫자 추출중",end='')

#정렬을 위한 단순 리스트
numberSortList = []
#1~45 번의 각 자리의 나온 횟수 저장 리스트
numberCameOut = [0] * 453
numberCameOutSort = [0] * 45

#각 번호의 가중치 (다음 회차 확률 계산용) [0~45]
numberWeight = [0] * 45
WeightNumSort = [0] * 45
WeightValSort = [0] * 45
#최근 나온 회차 저장 리스트
recentList1 = []
recentList2 = []
recentList3 = []
recentList4 = []
recentList5 = []

#1~45 번의 각 중복 발생 수
duplicateNumber = [0] * 45
duplicateNumberSort = [0] * 45
for i in range(45):
    numberSortList.append(i+1)
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

    if i > roundOfEvent-6:
        recentList1.append(nowNumList)
    # 각 번호 나온 횟수 추가
    for j in range(7):
        for k in range(1,46):
            if nowNumList[j] == k:
                numberCameOut[k-1] += 1
                num5check(k);
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
    

print("==== 전체 나온 횟수 ====")
for i in range(45):
    print('',end='')
    if i < 9:
        print(i+1,end=' ')
    else:
        print(i+1,end='')
    print("숫자 ",end='가 나올확률은')
    print(round((numberCameOut[i]/roundOfEvent)*100.0,3),end='%이고 연속으로 나온횟수는')
    print(duplicateNumber[i],end='회 이고\n')
print("가장 최근에 뽑았던 숫자 나열은")
for i in range(4):
    print(i+1, end='. ')
    for j in range(6):
        print(recentList1[i][j],end='')
        numberWeight[recentList1[i][j]-1] -=  2*(5 - i)
        if j == 5:
            print('')
        else:
            print(',',end='')
print("와 같아")
print("1~45의 숫자 중 6개의 중복되지 않는 번호를 뽑아 나열한다면 다음번에 가장 나올 확률이 높은 숫자로 이루어진 숫자의 나열을 10가지 알려줘")
print("단, 10가지 전체 나열된 숫자들이 중복된 횟수는 3번 이하여야해.")

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

for i in range(45):
    print('\t',end='')
    if i < 10:
        print(i,end='  : ')
    else:
        print(i,end=' : ')
    print(numberCameOut[i],end=' 회')
    print()


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
    print(duplicateNumberSort[i],end='회 ')
    
    # 가중치 추가
    if i < 10:
        numberWeight[numberSortList[i]-1] += 40
    elif i < 20:
        numberWeight[numberSortList[i]-1] += 50
    elif i < 30:
        numberWeight[numberSortList[i]-1] += 5
    else:
        numberWeight[numberSortList[i]-1] += 2

print("정렬중",end='')
count = 0
for k in range(45):
    MAXNUM = 0
    for i in range(45):
        if duplicateNumber[i] > MAXNUM:
            MAXNUM = duplicateNumber[i]
    if MAXNUM > 0:
        for j in range(45):
            if MAXNUM == duplicateNumber[j]:
                numberSortList[count] = j+1
                numberCameOutSort[count] = numberCameOut[j]
                duplicateNumberSort[count] = duplicateNumber[j]
                duplicateNumber[j] = 0
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
    print(duplicateNumberSort[i],end='회 ')
    
    # 가중치 추가
    if i < 10:
        numberWeight[numberSortList[i]-1] += 20
    elif i < 20:
        numberWeight[numberSortList[i]-1] += 10
    elif i < 30:
        numberWeight[numberSortList[i]-1] += 1
    else:
        numberWeight[numberSortList[i]-1] += 10


for i in range(45):
    if numberSortList[i] < 10:
        numberWeight[numberSortList[i]-1] += 1
    elif numberSortList[i] < 20:
        numberWeight[numberSortList[i]-1] += 1
    elif numberSortList[i] < 30:
        numberWeight[numberSortList[i]-1] += 5
    elif numberSortList[i] < 40:
        numberWeight[numberSortList[i]-1] += 1
    else:
        numberWeight[numberSortList[i]-1] += 1

#출현 번호대 가중치
print("출현 번호대 가중치 증가",end='')
count = 4
for k in range(9):
    MAXNUM = 0
    for i in range(9):
        if showNum5[i] > MAXNUM:
            MAXNUM = showNum5[i]
    if MAXNUM > 0:
        for j in range(9):
            if MAXNUM == showNum5[j]:
                for l in range(5):
                    numberWeight[l + j * 5] += count
                if count > 0:
                    count -= 1
                showNum5[j] = 0
                print(".",end='')


print("\t==== 가중치정렬 ====")
for i in range(45):
    print('\t',end='')
    if i < 10:
        print(i,end='  : ')
    else:
        print(i,end=' : ')
    print(numberWeight[i],end='')
    print()

print("정렬중",end='')
count = 0
MAXWEIGHT = 0
for k in range(45):
    MAXNUM = 0
    for i in range(45):
        if numberWeight[i] > MAXNUM:
            MAXNUM = numberWeight[i]
    if MAXNUM > 0:
        if MAXWEIGHT == 0:
            MAXWEIGHT = MAXNUM
        for j in range(45):
            if MAXNUM == numberWeight[j]:
                WeightNumSort[count] = j+1
                WeightValSort[count] = numberWeight[j]
                numberWeight[j] = 0
                count += 1
                print(".",end='')
print()
print("\t==== 가중치정렬 ====")
for i in range(45):
    print('\t',end='')
    if WeightNumSort[i] < 10:
        print(WeightNumSort[i],end='  : ')
    else:
        print(WeightNumSort[i],end=' : ')
    print(WeightValSort[i],end='/')
    print(MAXWEIGHT)


for i in range(45):
    WeightValSort[i] = WeightValSort[i] / MAXWEIGHT

numPacket = []
odd = 0
even = 0
i = 0
jung = 0
print("계산적용 랜덤숫자생성중 : 짝홀수 균형")
while True:
    if i > testNum:
        break
    else: 
        i+=1
 
    while True:
        # 가중치를 반영하여 6개 숫자 선택 (중복 없음)
        stat = False
        selected_numbers = random.choices(WeightNumSort, weights=WeightValSort, k=6)
        # 중복을 방지하기 위해 반복 선택
        while len(set(selected_numbers)) < 6:
            selected_numbers = random.choices(WeightNumSort, weights=WeightValSort, k=6)

        # 홀짝수 균형 유지 3:3 2:4  (5:1 x)
        even = 0
        odd = 0
        for j in range(len(selected_numbers)):
            if selected_numbers[j]%2 == 0:
                even += 1
            else:
                odd += 1
        if even >= 2 and even <= 4:
            stat = True
        
        # 합계 200 이상 지양
        if stat == True:
            sum = 0
            for j in range(len(selected_numbers)):
                sum += selected_numbers[j]
            if sum > 200 or sum < 70:
                stat = False

        # 끝자리 같은 숫자 3개 이상이면 베재
        if stat == True:
            endNum = []
            for j in range(len(selected_numbers)):
                endNum.append(selected_numbers[j]%10)
            count_dict = Counter(endNum)
            for num, count in count_dict.items():
                if count > 2:
                    stat = False
            
        if stat == True:
            # 선택된 숫자 정렬 후 출력
            selected_numbers.sort()
            if (test == False):
                if selected_numbers in numPacket:
                    jung += 1
                else:
                    numPacket.append(selected_numbers)
            else:
                numPacket.append(selected_numbers)

            break

    

# 랜덤 테스트 결과

if test == True:
    corllectTotal = 0
    corllect1 = 0
    corllect2 = 0
    corllect3 = 0
    corllect4 = 0
    print("예시답 입력:", end='')
    insert = WinNum  # 문자열 입력 받기
    numberList = list(map(int, insert.split()))  # 공백 기준으로 나누고 정수 변환

    for i in range(testNum):
        selected_numbers = random.choices(WeightNumSort, k=6)
        while len(set(selected_numbers)) < 6:
            selected_numbers = random.choices(WeightNumSort, k=6)
        selected_numbers.sort()

        counter1 = selected_numbers
        counter2 = numberList
        common_counts = 0
        for j in range(len(counter1)):
            for k in range(len(counter2)):
                if counter1[j] == counter2[k]:
                    common_counts += 1
        if common_counts == 3:
            corllectTotal+=1
            corllect4 += 1
        if common_counts == 4:
            corllectTotal += 1
            corllect3 += 1
        if common_counts == 5:
            corllectTotal += 1
            corllect2 += 1
        if common_counts == 6:
            corllectTotal += 1
            corllect1 += 1

    print("================================\t 계산식 적용 전 \t================================")
    print("시도수:",end='')
    print(testNum)
    print("성공수:",end='')
    print(corllectTotal)
    print("확률 :",end='')
    print(round(corllectTotal/testNum*100,3),end='%')
    print()
    print("1등 : ", end='')
    print(round(corllect1/testNum*100,3), end='% : ')
    print(corllect1)
    print("2등 : ", end='')
    print(round(corllect2/testNum*100,3), end='% : ')
    print(corllect2)
    print("3등 : ", end='')
    print(round(corllect3/testNum*100,3), end='% : ')
    print(corllect3)
    print("4등 : ", end='')
    print(round(corllect4/testNum*100,3), end='% : ')
    print(corllect4)


    corllectTotal = 0
    corllect1 = 0
    corllect2 = 0
    corllect3 = 0
    corllect4 = 0
    for i in range(testNum):
        counter1 = numPacket[i]
        counter2 = numberList
        common_counts = 0
        for j in range(len(counter1)):
            for k in range(len(counter2)):
                if counter1[j] == counter2[k]:
                    common_counts += 1
        if common_counts == 3:
            corllectTotal+=1
            corllect4 += 1
        if common_counts == 4:
            corllectTotal += 1
            corllect3 += 1
        if common_counts == 5:
            corllectTotal += 1
            corllect2 += 1
        if common_counts == 6:
            corllectTotal += 1
            corllect1 += 1


    print("")
    print("================================\t 계산식 적용 후 \t================================")
    print("시도수:",end='')
    print(testNum)
    print("성공수:",end='')
    print(corllectTotal)
    print("확률 :",end='')
    print(round(corllectTotal/testNum*100,3),end='%')
    print()
    print("1등 : ", end='')
    print(round(corllect1/testNum*100,3), end='% : ')
    print(corllect1)
    print("2등 : ", end='')
    print(round(corllect2/testNum*100,3), end='% : ')
    print(corllect2)
    print("3등 : ", end='')
    print(round(corllect3/testNum*100,3), end='% : ')
    print(corllect3)
    print("4등 : ", end='')
    print(round(corllect4/testNum*100,3), end='% : ')
    print(corllect4)

print("================================\t 공식 적용 후 추첨번호 \t================================")
for i in range(20):
    print(random.choice(numPacket))