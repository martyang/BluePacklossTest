#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import serial
import time
import xlwt
from serial import SerialException
import serial.tools.list_ports_windows

CLEAR = "01e0fc0191"
GET_PACK = "01e0fc0190"
DUT = "01e0fc02af01"
DH5 = "2DH5"
DH3 = "2DH3"
RxHECErrcntr = "RxHECErrorCntr"
RxCRCErrcntr = "RxCRCErrorCntr"
TIME = 10


def getData(data):
    return int(data.strip().split(' ')[-1])


timestr = time.strftime('%Y%m%d%H%M%S', time.localtime())
print(timestr)
path = os.getcwd()  # 获取当前工作路径
print(path)
file = open(path + '\\config.txt', 'rb')
a = file.read().decode('utf-8')
PORT = a.split('\n')[0].split(' ')[0]
BAUD = a.split('\n')[0].split(' ')[1]
TIME = int(a.split('\n')[1].split(' ')[0])
print(PORT, BAUD, TIME)

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet("3296丢包测试")
worksheet.write(0, 0, "2DH5")
worksheet.write(0, 1, "2DH3")
worksheet.write(0, 2, "ERROR")
worksheet.write(0, 3, "Rate")
row = 1

try:
    sercom = serial.Serial(PORT, BAUD, timeout=5)
    print("serial open success!")
    time.sleep(0.5)
    while True:
        dh5 = 1
        dh3 = 1
        error = 0
        sercom.write(bytes.fromhex(CLEAR))
        print(CLEAR)
        time.sleep(TIME)
        sercom.write(bytes.fromhex(GET_PACK))
        print(GET_PACK)
        time.sleep(0.1)
        while sercom.inWaiting():
            recv_data = sercom.readline().decode("utf-8")
            print(recv_data)
            if DH5 in recv_data:
                dh5 = getData(recv_data)
                print(dh5)
            if DH3 in recv_data:
                dh3 = getData(recv_data)
                print(dh3)
            if RxHECErrcntr in recv_data or RxCRCErrcntr in recv_data:
                error += getData(recv_data)
                print(error)
        worksheet.write(row, 0, dh5)
        worksheet.write(row, 1, dh3)
        worksheet.write(row, 2, error)
        if dh5 + dh3 == 0:
            worksheet.write(row, 3, 0)
        else:
            worksheet.write(row, 3, error / (dh3 + dh5))
        row += 1
        workbook.save('./%s.xls' % timestr)
except SerialException:
    print('串口异常')
    time.sleep(2)
