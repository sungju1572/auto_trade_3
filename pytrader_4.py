import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from Kiwoom_4 import *
import time
from os import environ


form_class = uic.loadUiType("exam.ui")[0]


#해상도 고정 함수 추가
def suppress_qt_warning():
    environ["QT_DEVICE_PIXEL_RATIO"] = "0"
    environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    environ["QT_SCREEN_SCALE_FACTORS"] = "1"
    environ["QT_SCALE_FACTOR"] = "1"

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
            
        self.trade_set = False
        
        self.trade_stocks_done = False

        self.kiwoom = Kiwoom(self) #객체생성
        self.kiwoom.comm_connect() #연결


        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)


        self.timer2 = QTimer(self)
        self.timer2.start(1000 *10)
        self.timer2.timeout.connect(self.timeout2)

        accouns_num = int(self.kiwoom.get_login_info("ACCOUNT_CNT"))
        accounts = self.kiwoom.get_login_info("ACCNO")

        accounts_list = accounts.split(';')[0:accouns_num]
        
        
        self.comboBox.addItems(accounts_list) #콤보박스 1에 계좌목록 추가
        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.check_balance)
        self.pushButton_3.clicked.connect(self.check_stock)
        self.pushButton_5.clicked.connect(self.trade_start)
        self.pushButton_2.clicked.connect(self.delete_row)
        self.pushButton_4.clicked.connect(self.ready_trade)


        self.row_count = 0 #tableWidget_3 에서 행 카운트하는용
        self.stock_list = [] #주시종목 담은 리스트
        self.stock_ticker_list = [] #주시종목 티커 리스트 
        self.account_number ="" #계좌
        self.take_profit = 0 #익절기준
        
        
        #self.lineEdit_8.textChanged.connect(self.profit_percent)# 익절 기준
        
    


    #종목 ui에 띄우기
    def code_changed(self):
        code = self.lineEdit.text()
        name = self.kiwoom.get_master_code_name(code)
        self.lineEdit_2.setText(name)
        
    #서버연결
    def timeout(self):
        market_start_time = QTime(9, 0, 0)
        current_time = QTime.currentTime()

        if current_time > market_start_time and self.trade_stocks_done is False:
            #self.trade_stocks()
            self.trade_stocks_done = True

        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time


        state = self.kiwoom.get_connect_state()
        if state == 1:
            state_msg = "서버 연결 중"
            
        else:
            state_msg = "서버 미 연결 중"

        self.statusbar.showMessage(state_msg + " | " + time_msg)
        

    #잔고 실시간으로 갱신
    def timeout2(self):
        if self.checkBox.isChecked():
            self.check_balance()


    #현재가격저장        
    def present_price(self):
        price = self.kiwoom.price
        self.lineEdit_3.setText(str(price))
     
    #주식 잔고 
    def check_balance(self):
        self.kiwoom.reset_opw00018_output()
        account_number = self.comboBox.currentText()

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")

        while self.kiwoom.remained_data:
            time.sleep(0.2)
            self.kiwoom.set_input_value("계좌번호", account_number)
            self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 2, "2000")

        # opw00001
        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00001_req", "opw00001", 0, "2000")

        # balance
        item = QTableWidgetItem(self.kiwoom.d2_deposit)
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.tableWidget.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw00018_output['single'][i - 1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)

        self.tableWidget.resizeRowsToContents()

        # Item list
        item_count = len(self.kiwoom.opw00018_output['multi'])
        self.tableWidget_2.setRowCount(item_count)

        for j in range(item_count):
            row = self.kiwoom.opw00018_output['multi'][j]
            for i in range(len(row)):
                item = QTableWidgetItem(row[i])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(j, i, item)

        self.tableWidget_2.resizeRowsToContents()
        
    
        
    #주시 종목에 설정한 종목 넣기
    def check_stock(self):
        code = self.lineEdit.text()
        name = self.kiwoom.get_master_code_name(code)
        count = self.comboBox_2.currentText()
        
        if count =="3개":
            high = self.lineEdit_3.text()
            middle = self.lineEdit_4.text()
            low = self.lineEdit_5.text()
            price = self.lineEdit_9.text()
            
            self.tableWidget_3.setRowCount(self.row_count+1)
            self.tableWidget_3.setColumnCount(7)
            self.tableWidget_3.setItem(self.row_count,0,QTableWidgetItem(name))
            self.tableWidget_3.setItem(self.row_count,1,QTableWidgetItem(high))
            self.tableWidget_3.setItem(self.row_count,2,QTableWidgetItem(middle))
            self.tableWidget_3.setItem(self.row_count,3,QTableWidgetItem(low))
            self.tableWidget_3.setItem(self.row_count,4,QTableWidgetItem(code))
            self.tableWidget_3.setItem(self.row_count,5,QTableWidgetItem(price))
            self.tableWidget_3.setItem(self.row_count,6,QTableWidgetItem(count))
            self.row_count+=1
            
            self.plainTextEdit.appendPlainText("종목추가 : "+ name)
            
            
        elif count =="2개":
            high = self.lineEdit_6.text()
            low = self.lineEdit_7.text()
            middle = (float(high)+float(low))/2
            price = self.lineEdit_9.text()
            
            
            self.tableWidget_3.setRowCount(self.row_count+1)
            self.tableWidget_3.setColumnCount(7)
            self.tableWidget_3.setItem(self.row_count,0,QTableWidgetItem(name))
            self.tableWidget_3.setItem(self.row_count,1,QTableWidgetItem(high))
            self.tableWidget_3.setItem(self.row_count,2,QTableWidgetItem(str(middle)))
            self.tableWidget_3.setItem(self.row_count,3,QTableWidgetItem(low))
            self.tableWidget_3.setItem(self.row_count,4,QTableWidgetItem(code))
            self.tableWidget_3.setItem(self.row_count,5,QTableWidgetItem(price))
            self.tableWidget_3.setItem(self.row_count,6,QTableWidgetItem(count))
            self.row_count+=1
            
            self.plainTextEdit.appendPlainText("종목추가 : "+ name)
            
    #제거 버튼 눌렀을때 테이블에서 행삭제        
    def delete_row(self):
        select = self.tableWidget_3.selectedItems()
        for i in select:
            row = i.row()
            self.tableWidget_3.removeRow(row)
            self.row_count-=1
            self.plainTextEdit.appendPlainText("선택 종목삭제")
        
        
    #호가 받아오는 함수
    def get_hoga(self, trcode):
        self.kiwoom.set_input_value("종목코드", trcode)
        self.kiwoom.comm_rq_data("opt10004_req", "opt10004", 0, "3000")
        
    #전일종가 받아오는 함수
    def get_last_close(self, trcode):
        self.kiwoom.set_input_value("종목코드", trcode)
        self.kiwoom.comm_rq_data("opt10002_req", "opt10002", 0, "3000")



    #tableWidget_3 에서 값 얻어오기
    def get_label(self):
        init_num = 0
        init_list = []
        while init_num < self.row_count:
            sec_list = []
            for i in range(7):
                sec_list.append(self.tableWidget_3.item(init_num,i).text())
            init_list.append(sec_list)
            init_num += 1
        print(init_list)
        return init_list
                
    
    
    def ready_trade(self):
        self.account_number = self.comboBox.currentText()
        self.stock_list = self.get_label()
        
        for i in range(len(self.stock_list)):
            self.kiwoom.dic[self.stock_list[i][0] + '_status'] = '초기상태' 
            self.kiwoom.dic[self.stock_list[i][0] + '_rebuy'] = 1  
            self.kiwoom.dic[self.stock_list[i][0] + '_initial'] = 0 
            self.kiwoom.dic[self.stock_list[i][0] + '_buy_count'] = 0 
            self.kiwoom.dic[self.stock_list[i][0] + '_sell_price'] = 0 
            self.kiwoom.dic[self.stock_list[i][0] + '_rebuy_count'] = 0
            self.kiwoom.dic[self.stock_list[i][0] + '_buy_line'] = ""
            
            
            self.get_hoga(self.stock_list[i][4])
            self.kiwoom.dic[self.stock_list[i][0] + '_hoga'] = self.kiwoom.hoga
            time.sleep(0.5)
            
            self.get_last_close(self.stock_list[i][4])
            self.kiwoom.dic[self.stock_list[i][0] + '_last_close'] = self.kiwoom.last_close 
            time.sleep(0.5)

            
            #매도조건 상태 2가지
            self.kiwoom.dic[self.stock_list[i][0] + '_sell_status1'] = '초기상태'
            self.kiwoom.dic[self.stock_list[i][0] + '_sell_status2'] = '초기상태'
        
        
            self.plainTextEdit.appendPlainText("거래준비완료 | 종목 :" + self.stock_list[i][0]  )

        print(self.kiwoom.dic)
    

    #거래시작 버튼눌렀을때 주시 종목별 구독
    def trade_start(self):
        self.account_number = self.comboBox.currentText()
        self.stock_list = self.get_label()
        

        
        for i in range(len(self.stock_list)):
            
            if i ==0:
                self.kiwoom.SetRealReg("1000", self.stock_list[i][4], "20;10", "0")
                self.stock_ticker_list.append(self.stock_list[i][4]) 
            else: 
                self.kiwoom.SetRealReg("1000", self.stock_list[i][4], "20;10", "1")
                self.stock_ticker_list.append(self.stock_list[i][4])
            




    #익절기준 변경점
   # def profit_percent(self):
        #self.take_profit = float(self.lineEdit_8.text())




if __name__ == "__main__":
    
    import sys
    
    suppress_qt_warning()
    
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()

