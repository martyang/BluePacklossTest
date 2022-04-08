import time

import pyvisa as visa
import xlwt
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QDialog
from pyvisa import VisaIOError
from serial import SerialException
from packagetest import Ui_MainWindow
import serial
import serial.tools.list_ports_windows


class runThread(QThread):
    CLEAR = "01e0fc0191"
    GET_PACK = "01e0fc0190"
    DUT = "01e0fc02af01"
    DH5 = "2DH5"
    DH3 = "2DH3"
    RxHECErrcntr = "RxHECErrorCntr"
    RxCRCErrcntr = "RxCRCErrorCntr"
    data = pyqtSignal(str)
    status = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, ser_com, freq_list, interval, psa):
        super(runThread, self).__init__()
        print('创建子线程')
        self.ser_com = ser_com
        self.freq_list = freq_list
        self.interval = interval
        self.psa = psa
        self.running = True

    def stopTest(self):
        self.running = False
        print('stoptest', self.running)
        # time.sleep(0.5)

    def run(self):
        print('run')
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet("丢包测试")
        worksheet.write(0, 0, "Freq")
        print('初始化sheet')
        for i in range(4):
            worksheet.write(0, 4 * i + 1, "2DH5")
            worksheet.write(0, 4 * i + 2, "2DH3")
            worksheet.write(0, 4 * i + 3, "ERROR")
            worksheet.write(0, 4 * i + 4, "Rate")
        time_str = time.strftime('%Y%m%d%H%M%S', time.localtime())
        row = 1

        length = len(self.freq_list) // 2
        for i in range(length):
            start_fre = int(self.freq_list[2 * i])
            end_fre = int(self.freq_list[2 * i + 1])
            test_fre = start_fre
            while test_fre <= end_fre and self.running:
                print("Freq:" + str(test_fre) + "MHz")
                worksheet.write(row, 0, test_fre)
                self.psa.write('FREQ:FIX %dMHZ\n' % test_fre)
                self.pack_test(worksheet, row)
                test_fre += self.interval
                row += 1
                workbook.save('./%s.xls' % time_str)
            self.progress.emit(i + 1)
        self.status.emit('开始')
        self.ser_com.close()

    def pack_test(self, worksheet, row):
        print('连接串口')
        time.sleep(0.5)
        test_count = 0
        while test_count < 4:
            dh5 = 1
            dh3 = 1
            error = 0
            if self.ser_com.isOpen():
                self.ser_com.write(bytes.fromhex(self.CLEAR))
                time.sleep(5)
                self.ser_com.write(bytes.fromhex(self.GET_PACK))
                time.sleep(0.1)
                while self.ser_com.inWaiting():
                    recv_data = self.ser_com.readline().decode("utf-8")
                    self.data.emit(recv_data)
                    if self.DH5 in recv_data:
                        dh5 = int(recv_data.strip().split(' ')[-1])
                    if self.DH3 in recv_data:
                        dh3 = int(recv_data.strip().split(' ')[-1])
                    if self.RxHECErrcntr in recv_data or self.RxCRCErrcntr in recv_data:
                        error += int(recv_data.strip().split(' ')[-1])
                worksheet.write(row, 4 * test_count + 1, dh5)
                worksheet.write(row, 4 * test_count + 2, dh3)
                worksheet.write(row, 4 * test_count + 3, error)
                if dh5 + dh3 == 0:
                    worksheet.write(row, 4 * test_count + 4, 0)
                else:
                    worksheet.write(row, 4 * test_count + 4, error / (dh3 + dh5))
                test_count += 1


class UiWindows(QMainWindow, Ui_MainWindow):
    stop = pyqtSignal()

    def __init__(self):
        super(UiWindows, self).__init__()
        self.setupUi(self)
        self.freq_list = []
        self.pushButton.clicked.connect(self.add_freq)
        self.pushButton_2.clicked.connect(self.startTest)
        self.set_ser_list()
        self.ser_com = None
        self.psa = None
        self.show()

    def set_ser_list(self):
        """
        获取串口列表
        :return:
        """
        port_list = serial.tools.list_ports_windows.comports()
        self.comboBox.clear()
        for port in port_list:
            self.comboBox.addItem(str(port))

    def getCom(self):
        combo_str = self.comboBox.currentText()
        return combo_str.strip().split(' ')[0]

    def add_freq(self):
        start_freq = self.lineEdit_2.text()
        end_freq = self.lineEdit_3.text()
        if start_freq == '' or end_freq == '':
            QMessageBox(QMessageBox.Warning, '错误警告', '频率不能为空！').exec_()
        elif int(end_freq) > 6000 or float(start_freq) < 0.25:
            QMessageBox(QMessageBox.Warning, '错误警告', '超出信号源频率范围！').exec_()
        elif int(start_freq) > int(end_freq):
            QMessageBox(QMessageBox.Warning, '错误警告', '起始频率不能比终止频率大！').exec_()
        else:
            self.freq_list.append(start_freq)
            self.freq_list.append(end_freq)
            self.textBrowser_2.append(start_freq + "--" + end_freq)

    def update_text(self, data):  # 更新
        self.textBrowser.append(data)

    def changeStatus(self, status):
        self.pushButton_2.setText(status)

    def update_progress(self, value):
        self.progressBar.setValue(value)

    def startTest(self):
        PORT = self.getCom()
        BAUD = self.lineEdit_4.text()
        IP = self.lineEdit.text()
        INTERVAL = int(self.lineEdit_5.text())
        POWER = int(self.lineEdit_6.text())
        self.progressBar.setMaximum(len(self.freq_list) // 2)

        if len(self.freq_list) == 0:
            QMessageBox(QMessageBox.Warning, '错误警告', '没有加入任何测试频点！').exec_()
        else:
            if self.pushButton_2.text() == '开始':
                try:
                    self.ser_com = serial.Serial(PORT, BAUD, timeout=5)
                    self.textBrowser.append('串口初始化成功！')
                    if self.psa is None:
                        print('开始连接信号源')
                        self.textBrowser.append('开始连接信号源')
                        rm = visa.ResourceManager()
                        self.psa = rm.open_resource('TCPIP0::%s::inst0::INSTR' % IP, open_timeout=1000)
                        self.textBrowser.append('信号源初始化成功！')

                    self.pushButton_2.setText('停止')
                    self.textBrowser.append('开始测试')
                    self.progressBar.setValue(0)
                    self.psa.write('POWer:AMPLitude %dDBM\n' % POWER)
                    self.psa.write('OUTPut:STATe ON\n')
                    self.myThread = runThread(self.ser_com, self.freq_list, INTERVAL, self.psa)  # 不加self程序闪退？
                    self.myThread.data.connect(self.update_text)
                    self.myThread.status.connect(self.changeStatus)
                    self.myThread.progress.connect(self.update_progress)
                    self.stop.connect(self.myThread.stopTest)
                    self.myThread.start()
                    print('start')
                except SerialException:
                    QMessageBox(QMessageBox.Warning, '错误警告', '串口无法打开').exec_()
                except VisaIOError:
                    QMessageBox(QMessageBox.Warning, '错误警告', '无法连接信号源，请确保电脑与设备处于同一IP地址').exec_()
            else:
                self.stop.emit()
