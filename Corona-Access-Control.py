#!/usr/bin/env python
# coding: utf-8

# 라이브러리 불러오기
from pandas import DataFrame
from datetime import datetime
import os

import cv2
import pyzbar
from pyzbar.pyzbar import decode
from pyzbar.pyzbar import ZBarSymbol

import winsound as ws

# 비프음 함수
def beepsound():
    freq = 1000    # range : 37 ~ 32767
    dur = 200     # ms
    ws.Beep(freq, dur) # winsound.Beep(frequency, duration)

# 오늘 날짜 정보 가져오기
today = datetime.today().strftime('%Y%m%d')

# PC 이름 입력받기
pc_num= input('PC 번호를 입력하세요 : ')

# 파일 이름 생성하기
file_name = today + '_' + 'PC' + '_' + pc_num

# 같은 이름의 파일이 있다면 파일 이름 수정
if os.path.exists('./'+ file_name +'.xlsx') == True:
    print('\n오늘 자 파일이 이미 존재합니다')
    num = input('몇 번째 파일입니까? : ')
    file_name = today + '_' + 'PC' + '_' + pc_num + '_' + '(' + num + ')'

# 성도 정보 데이터 프레임 생성
df = DataFrame(columns = ['교회', '구역', '이름','소속','연락처','체온','문진사항','방문시간'])
count = 1
brethren = []

# QRCode 스캔 및 성도 정보 입력
capture = cv2.VideoCapture(1)

while True:
    _, frame = capture.read()
    cv2.imshow('QR Code Scanner', frame)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    decoded_data = decode(gray, symbols=[ZBarSymbol.QRCODE])
    
    try:
        data = decoded_data[0][0]
        
        # 스캔 정보 utf-8로 디코딩
        data = data.decode('utf-8')
        
        # 비프음 내기
        print(beepsound())
        print("\nQR코드가 스캔되었습니다.")
        print(data)
            
        # 성도 정보 입력
        # QRcode 형식: "교회/구역/이름/소속/연락처/"
        brethren = data
        brethren = list(brethren.split('/'))
                
        # 체온과 문진사항 입력
        brethren[5] = input('체온을 입력해주세요: ')
        brethren.append(input('문진사항을 입력해주세요(o/x): '))
        if brethren[6] == '': # Enter 입력 'x' 처러
            brethren[6] = 'x'
            
        # 현재시간을 방문시간으로 저장
        brethren.append(datetime.today().strftime("%Y/%m/%d %H:%M"))
        
        # 성도 정보 데이터 프레임에 저장 및 엑셀 저장
        df.loc[count] = brethren
        df.to_excel(file_name+ '.xlsx')
        count = count + 1
        
        # 종료 or 이전 기록 삭제
        opt = input("프로그램 종료: q / 이전 기록 삭제: d / 다음 단계로 이동: Enter: ")
        if opt == 'q':
            break
        elif opt == 'd':
            df.loc[count-1] = ['','','','','','','','']
            df.to_excel(file_name + '.xlsx')
            count = count - 1
        else:
            pass       
    except:
        pass
    
    # 'q'입력 시 스캔 종료
    key = cv2.waitKey(1)
    if key == ord('q'):
        break
