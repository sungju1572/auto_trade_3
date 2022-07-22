import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
import pandas as pd
import sqlite3
import datetime
import numpy as np


TR_REQ_TIME_INTERVAL = 0.2

class Kiwoom(QAxWidget):
    def __init__(self, ui):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.ui = ui
        
        
        self.dic = {}
        
        self.rebuy = 1 #재매수 횟수 (1번만 가능하도록)
      

        
    #COM오브젝트 생성
    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") #고유 식별자 가져옴

    #이벤트 처리
    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._event_connect) # 로그인 관련 이벤트 (.connect()는 이벤트와 슬롯을 연결하는 역할)
        self.OnReceiveTrData.connect(self._receive_tr_data) # 트랜잭션 요청 관련 이벤트
        self.OnReceiveChejanData.connect(self._receive_chejan_data) #체결잔고 요청 이벤트
        self.OnReceiveRealData.connect(self._handler_real_data) #실시간 데이터 처리

    #로그인
    def comm_connect(self):
        self.dynamicCall("CommConnect()") # CommConnect() 시그널 함수 호출(.dynamicCall()는 서버에 데이터를 송수신해주는 기능)
        self.login_event_loop = QEventLoop() # 로그인 담당 이벤트 루프(프로그램이 종료되지 않게하는 큰 틀의 루프)
        self.login_event_loop.exec_() #exec_()를 통해 이벤트 루프 실행  (다른데이터 간섭 막기)

    #이벤트 연결 여부
    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()

    #종목리스트 반환
    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market) #종목리스트 호출
        code_list = code_list.split(';')
        return code_list[:-1]

    #종목명 반환
    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code) #종목명 호출
        return code_name

    #통신접속상태 반환
    def get_connect_state(self):
        ret = self.dynamicCall("GetConnectState()") #통신접속상태 호출
        return ret

    #로그인정보 반환
    def get_login_info(self, tag):
        ret = self.dynamicCall("GetLoginInfo(QString)", tag) #로그인정보 호출
        return ret

    #TR별 할당값 지정하기
    def set_input_value(self, id, value):
        self.dynamicCall("SetInputValue(QString, QString)", id, value) #SetInputValue() 밸류값으로 원하는값지정 ex) SetInputValue("비밀번호"	,  "")

    #통신데이터 수신(tr)
    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no) 
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    #실제 데이터 가져오기
    def _comm_get_data(self, code, real_type, field_name, index, item_name): 
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)", code, #더이상 지원 안함??
                               real_type, field_name, index, item_name)
        return ret.strip()

    #수신받은 데이터 반복횟수
    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    #주문 (주식)
    def send_order(self, rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no):
        self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                         [rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no])
        
    #주문 (선물)    
    def send_order_fo(self, rqname, screen_no, acc_no,  code, order_type, slbytp, hoga, quantity, price, order_no):
        self.dynamicCall("SendOrderFO(QString, QString, QString, QString, int, QString, QString, int, QString, QString)",
                         [rqname, screen_no, acc_no, code, order_type, slbytp, hoga, quantity, price, order_no])


    #실시간 타입 구독신청
    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)

####
    #실시간 조회관련 핸들
    def _handler_real_data(self, trcode, real_type, data):
        
        # 체결 시간 
        time =  self.get_comm_real_data(trcode, 20)
        date = datetime.datetime.now().strftime("%Y-%m-%d ")
        time =  datetime.datetime.strptime(date + time, "%Y-%m-%d %H%M%S")
        print("체결시간 :", time, end=" ")

        # 현재가 
        price =  self.get_comm_real_data(trcode, 10)
        if price != "":
            price = float(price[1:])
            #print(trcode, ":", price)
            
        #시가
        start_price = self.get_comm_real_data(trcode, 16)

       
        
        for i in range(len(self.ui.stock_list)):
            if trcode == self.ui.stock_list[i][4]:
                print(i, "번째 :", self.ui.stock_list[i])
                
                if self.ui.stock_list[i][6] == "3개":
                    start_price = self.get_comm_real_data(trcode, 16)
                    price = self.get_comm_real_data(trcode, 10)
                    high = float(self.ui.stock_list[i][1])
                    middle = float(self.ui.stock_list[i][2])
                    low = float(self.ui.stock_list[i][3])
                    name = self.ui.stock_list[i][0]
                    buy_total_price = self.ui.stock_list[i][5]
                    
                    
                    if start_price  == "":
                        self.ui.plainTextEdit_2.appendPlainText("시가 입력 대기중 :" + name )
                    else:
                        start_price  = float(start_price[1:])
                        price = float(price[1:])
                    
                    
                        self.dic[self.ui.stock_list[i][0] + '_start_price'] = start_price  
                        self.dic[self.ui.stock_list[i][0] + '_high'] = high
                        self.dic[self.ui.stock_list[i][0] + '_middle'] = middle
                        self.dic[self.ui.stock_list[i][0] + '_low'] = low
                        self.dic[self.ui.stock_list[i][0] + '_price'] = price
                        self.dic[self.ui.stock_list[i][0] + '_trcode'] = trcode
                        self.dic[self.ui.stock_list[i][0] + '_name'] = name
                        self.dic[self.ui.stock_list[i][0] + '_buy_total'] = buy_total_price
                        
                        print("3개 list", self.dic)
                        
                        self.strategy(name)
                        
                    
                    #strategy(시가 , 상단, 중단, 하단, 현재가, trcode)
                elif self.ui.stock_list[i][6] == "2개":
                    
                    start_price = self.get_comm_real_data(trcode, 16)
                    price = self.get_comm_real_data(trcode, 10)
                    high = float(self.ui.stock_list[i][1])
                    middle = float(self.ui.stock_list[i][2])
                    low = float(self.ui.stock_list[i][3])
                    name = self.ui.stock_list[i][0]
                    buy_total_price = self.ui.stock_list[i][5]

                    
                    
                    if start_price  == "":
                        self.ui.plainTextEdit_2.appendPlainText("시가 입력 대기중 :" + name )
              
                    else:
                        start_price  = float(start_price[1:])
                        price = float(price[1:])
                        
                        
                        self.dic[self.ui.stock_list[i][0] + '_start_price'] = start_price  
                        self.dic[self.ui.stock_list[i][0] + '_high'] = high
                        self.dic[self.ui.stock_list[i][0] + '_middle'] = middle
                        self.dic[self.ui.stock_list[i][0] + '_low'] = low
                        self.dic[self.ui.stock_list[i][0] + '_price'] = price
                        self.dic[self.ui.stock_list[i][0] + '_trcode'] = trcode
                        self.dic[self.ui.stock_list[i][0] + '_name'] = name
                        self.dic[self.ui.stock_list[i][0] + '_buy_total'] = buy_total_price
                        
                      
                        print("2개 list", self.dic)


                        self.strategy_2(name)
                       
        
      

    #실시간 데이터 가져오기
    def get_comm_real_data(self, trcode, fid):
        ret = self.dynamicCall("GetCommRealData(QString, int)", trcode, fid)
        return ret

    #체결정보
    def get_chejan_data(self, fid):
        ret = self.dynamicCall("GetChejanData(int)", fid)
        return ret
    

    def get_server_gubun(self):
        ret = self.dynamicCall("KOA_Functions(QString, QString)", "GetServerGubun", "")
        return ret

    def _receive_chejan_data(self, gubun, item_cnt, fid_list):
        print(gubun)
        print(self.get_chejan_data(9203))
        print(self.get_chejan_data(302))
        print(self.get_chejan_data(900))
        print(self.get_chejan_data(901))

    #받은 tr데이터가 무엇인지, 연속조회 할수있는지
    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if next == '2': 
            self.remained_data = True
        else:
            self.remained_data = False
            
        #받은 tr에따라 각각의 함수 호출
        if rqname == "opt10081_req": #주식일봉차드 조회
            self._opt10081(rqname, trcode)
        elif rqname == "opw00001_req": #예수금 상세현황 요청
            self._opw00001(rqname, trcode)
        elif rqname == "opw00018_req": #계좌평가잔고 내역 요청
            self._opw00018(rqname, trcode)


        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    @staticmethod
    #입력받은데이터 정제    
    def change_format(data):
        strip_data = data.lstrip('-0')
        if strip_data == '' or strip_data == '.00':
            strip_data = '0'

        try:
            format_data = format(int(strip_data), ',d')
        except:
            format_data = format(float(strip_data))
        if data.startswith('-'):
            format_data = '-' + format_data

        return format_data

    #입력받은데이터(수익률) 정제
    @staticmethod
    def change_format2(data):
        strip_data = data.lstrip('-0')

        if strip_data == '':
            strip_data = '0'

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data

    def _opw00001(self, rqname, trcode):
        d2_deposit = self._comm_get_data(trcode, "", rqname, 0, "d+2추정예수금")
        self.d2_deposit = Kiwoom.change_format(d2_deposit)


    def _opt10081(self, rqname, trcode):
        data_cnt = self._get_repeat_cnt(trcode, rqname) #데이터 갯수 확인

        for i in range(data_cnt): #시고저종 거래량 가져오기
            date = self._comm_get_data(trcode, "", rqname, i, "일자")
            open = self._comm_get_data(trcode, "", rqname, i, "시가")
            high = self._comm_get_data(trcode, "", rqname, i, "고가")
            low = self._comm_get_data(trcode, "", rqname, i, "저가")
            close = self._comm_get_data(trcode, "", rqname, i, "현재가")
            volume = self._comm_get_data(trcode, "", rqname, i, "거래량")

            self.ohlcv['date'].append(date)
            self.ohlcv['open'].append(int(open))
            self.ohlcv['high'].append(int(high))
            self.ohlcv['low'].append(int(low))
            self.ohlcv['close'].append(int(close))
            self.ohlcv['volume'].append(int(volume))
            

    

    #opw박스 초기화 (주식)
    def reset_opw00018_output(self):
        self.opw00018_output = {'single': [], 'multi': []}

    #여러 정보들 저장 (주식)
    def _opw00018(self, rqname, trcode):
        # single data
        total_purchase_price = self._comm_get_data(trcode, "", rqname, 0, "총매입금액")
        total_eval_price = self._comm_get_data(trcode, "", rqname, 0, "총평가금액")
        total_eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, 0, "총평가손익금액")
        total_earning_rate = self._comm_get_data(trcode, "", rqname, 0, "총수익률(%)")
        estimated_deposit = self._comm_get_data(trcode, "", rqname, 0, "추정예탁자산")

        self.opw00018_output['single'].append(Kiwoom.change_format(total_purchase_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_profit_loss_price))

        total_earning_rate = Kiwoom.change_format(total_earning_rate)
        

        if self.get_server_gubun():
            total_earning_rate = float(total_earning_rate) / 100
            total_earning_rate = str(total_earning_rate)

        self.opw00018_output['single'].append(total_earning_rate)

        self.opw00018_output['single'].append(Kiwoom.change_format(estimated_deposit))

        # multi data
        rows = self._get_repeat_cnt(trcode, rqname)
        for i in range(rows):
            name = self._comm_get_data(trcode, "", rqname, i, "종목명")
            quantity = self._comm_get_data(trcode, "", rqname, i, "보유수량")
            purchase_price = self._comm_get_data(trcode, "", rqname, i, "매입가")
            current_price = self._comm_get_data(trcode, "", rqname, i, "현재가")
            eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, i, "평가손익")
            earning_rate = self._comm_get_data(trcode, "", rqname, i, "수익률(%)")

            quantity = Kiwoom.change_format(quantity)
            purchase_price = Kiwoom.change_format(purchase_price)
            current_price = Kiwoom.change_format(current_price)
            eval_profit_loss_price = Kiwoom.change_format(eval_profit_loss_price)
            earning_rate = Kiwoom.change_format2(earning_rate)

            self.opw00018_output['multi'].append([name, quantity, purchase_price, current_price, eval_profit_loss_price,
                                                  earning_rate])


    def strategy(self, name):
        
        list_1 = [k for k in self.dic.keys() if name in k ]
        
        
        status = self.dic[list_1[0]]
        rebuy = self.dic[list_1[1]]
        initial = self.dic[list_1[2]]
        buy_count = self.dic[list_1[3]]
        sell_price = self.dic[list_1[4]] 
        rebuy_count = self.dic[list_1[5]] 
        start_price = self.dic[list_1[6]]
        high = self.dic[list_1[7]]
        middle = self.dic[list_1[8]]
        low = self.dic[list_1[9]]
        price = self.dic[list_1[10]]
        trcode = self.dic[list_1[11]]
        name = self.dic[list_1[12]]
        buy_total_price = self.dic[list_1[13]]
       
        middle_low = float((float(middle) + float(low)) / 2 )#중하중단선
        
        buy_number = int(int(buy_total_price) / int(price))

        
        #초기상태
        if status == "초기상태":
            #시가가 중하중단선, 하단선 사이에 있으면 매수
            if start_price <= middle_low and start_price >= low :
                self.send_order('send_order', "0101", self.ui.account_number, 1, trcode, 10,  51600 ,"00", "" ) #지정가
                self.dic[list_1[0]] = "매수상태"
                self.dic[list_1[2]] = price
                self.dic[list_1[3]] = buy_number
                
                
                self.send_order('send_order', "0101", self.ui.account_number, 5, trcode, 10,  61700 ,"00", "" ) #지정가
                
                self.ui.plainTextEdit.appendPlainText("매수 :"+ name + " 매수가격 :" + str(price) + " 매수수량 : " + str(buy_number))

                
            #시가가 두개 사이에 있지 않다가, 현재가가 하단선을 뚫고 올라오면 매수
            elif start_price < low  :
                if price >=low :
                    self.send_order('send_order', "0101", self.ui.account_number, 1, trcode, buy_number,  0 ,"03", "" )
                    
                    self.dic[list_1[0]] = "매수상태"
                    self.dic[list_1[2]] = price
                    self.dic[list_1[3]] = buy_number
                    
                    self.ui.plainTextEdit.appendPlainText("현재가 하단선 돌파 매수 :" + name + " 매수가격 :" + str(price) + " 매수수량 : " + str(buy_number))
                else : 
                    self.ui.plainTextEdit_2.appendPlainText("현재가 하단선 밑 :" + name + "| 현재가 : " + str(price))
                    
            
            elif start_price > middle_low : 
                self.ui.plainTextEdit_2.appendPlainText("시가 중간선 위 : "+ name + "| 현재가 : " + str(price))   
            
            """
            elif start_price > middle_low and stay == 0: 
                self.ui.plainTextEdit_2.appendPlainText("대기중 | "+ name)
                self.dic[list_1[3]] = 1
            """
    
        #매수 상태
        elif status == "매수상태":
            #현재가가 하단선의 1%밑으로 내려가면 강제 청산 
                self.send_order('send_order', "0101", self.ui.account_number, 3, trcode, 10,  61700,"00", "1" ) #지정가
                
                

#재매수 고치기 (수량 얼마나 살지?)      
                
        elif status == "재매수대기상태" and rebuy == 1 :
            if price >=low :
                    sell_count = int(int(sell_price) / int(price))
                    self.send_order('send_order', "0101", self.ui.account_number, 1, trcode, sell_count,  0 ,"03", "" )
                    self.dic[list_1[0]] = "재매수상태"
                    self.ui.plainTextEdit.appendPlainText("현재가 하단선 돌파 매수 (재매수) :" + name + " 매수수량 : " + str(sell_count))
                    self.dic[list_1[1]] = 0
                    self.dic[list_1[5]] = sell_count
                   
            else :
                    self.ui.plainTextEdit_2.appendPlainText("재매수 대기중 :" + name + "| 현재가 : " + str(price))
                    
        elif status == "재매수대기상태" and rebuy == 0 :
            self.ui.plainTextEdit.appendPlainText("재매수 횟수 0 회, 거래종료 :" + name)
            self.dic[list_1[1]] = 2
            
            
            ######
            
        elif status == "재매수상태":
            if price < low * 0.99 :
                self.send_order('send_order', "0101", self.ui.account_number, 2, trcode, rebuy_count,  0 ,"03", "" )
                self.dic[list_1[0]] = "재매수대기상태"
                self.dic[list_1[4]] = price * buy_count
                self.ui.plainTextEdit.appendPlainText("하단가 1% 미만 지점 이탈 | 매도 :" + name)
            
            #보유중일때 현재가가 익절기준을 넘으면 매도   
            elif price >= initial + (initial * self.ui.take_profit/100):
                self.send_order('send_order', "0101", self.ui.account_number, 2, trcode, rebuy_count,  0 ,"03", "" )
                self.ui.plainTextEdit.appendPlainText("익절 퍼센트 초과 | 매도 :" + name)
                self.dic[list_1[0]] = "익절매도"
                
            




                

    #거래 전략 strategy(dic(시가 , 상단, 중단, 하단, 현재가, 종목코드, 종목이름, 현재상태, rebuy) 2개인경우
    
    def strategy_2(self, name):
        
        list_1 = [k for k in self.dic.keys() if name in k ]
        
        
        status = self.dic[list_1[0]]
        rebuy = self.dic[list_1[1]]
        initial = self.dic[list_1[2]]
        buy_count = self.dic[list_1[3]]
        sell_price = self.dic[list_1[4]] 
        rebuy_count = self.dic[list_1[5]] 
        start_price = self.dic[list_1[6]]
        high = self.dic[list_1[7]]
        middle = self.dic[list_1[8]]
        low = self.dic[list_1[9]]
        price = self.dic[list_1[10]]
        trcode = self.dic[list_1[11]]
        name = self.dic[list_1[12]]
        buy_total_price = self.dic[list_1[13]]

        buy_number = int(int(buy_total_price) / int(price))

        #초기상태
        if status == "초기상태":
            #시가가 상하준단선, 하단선 사이에 있으면 매수
            if start_price <= middle and start_price >= low :
                self.send_order('send_order', "0101", self.ui.account_number, 1, trcode, buy_number,  0 ,"03", "" )
                self.dic[list_1[0]] = "매수상태"
                self.dic[list_1[3]] = buy_number
                self.dic[list_1[2]] = price
                
                   
                self.ui.plainTextEdit.appendPlainText("매수 :"+ name + " 매수가격 :" + str(price) + " 매수수량 : " + str(buy_number))

                
            #시가가 두개 사이에 있지 않다가, 현재가가 하단선을 뚫고 올라오면 매수
            elif start_price < low  :
                if price >=low :
                    self.send_order('send_order', "0101", self.ui.account_number, 1, trcode, buy_number,  0 ,"03", "" )
                    
                    self.dic[list_1[0]] = "매수상태"
                    self.dic[list_1[2]] = price
                    self.dic[list_1[3]] = buy_number
                    
                    self.ui.plainTextEdit.appendPlainText("현재가 하단선 돌파 매수 :" + name + " 매수가격 :" + str(price)  + " 매수수량 : " + str(buy_number))
                else : 
                    self.ui.plainTextEdit_2.appendPlainText("현재가 하단선 밑 :" + name + "| 현재가 : " + str(price))
                    
                    
                    
            elif start_price > middle: 
                self.ui.plainTextEdit_2.appendPlainText("시가 중간선 위 : "+ name + "| 현재가 : " + str(price))   
            
 
    
        #매수 상태
        elif status == "매수상태":
            #현재가가 하단선의 1%밑으로 내려가면 강제 청산 
            if price < low * 0.99 :
                self.send_order('send_order', "0101", self.ui.account_number, 2, trcode, buy_count,  0 ,"03", "" )
                self.dic[list_1[0]] = "재매수대기상태"
                self.dic[list_1[4]] = price * buy_count
                self.ui.plainTextEdit.appendPlainText("하단가 1% 미만 지점 이탈 | 매도 :" + name)
            
            #보유중일때 현재가가 익절기준을 넘으면 매도
            elif price >= initial + (initial * self.ui.take_profit/100):
                self.send_order('send_order', "0101", self.ui.account_number, 2, trcode, buy_count,  0 ,"03", "" )
                self.ui.plainTextEdit.appendPlainText("익절 퍼센트 초과 | 매도 :" + name)
                self.dic[list_1[0]] = "익절매도"
               
        #재매수 대기상태 수정(얼만큼 살지)     
        elif status == "재매수대기상태" and rebuy == 1 :
            if price >=low :
                    sell_count = int(int(sell_price) / int(price))
                    self.send_order('send_order', "0101", self.ui.account_number, 1, trcode, sell_count,  0 ,"03", "" )
                    self.dic[list_1[0]] = "재매수상태"
                    self.ui.plainTextEdit.appendPlainText("현재가 하단선 돌파 매수 (재매수) :" + name + " 매수수량 : " + str(sell_count))
                    self.dic[list_1[1]] = 0
                    self.dic[list_1[5]] = sell_count
                   
            else :
                    self.ui.plainTextEdit_2.appendPlainText("재매수 대기중 :" + name  + "| 현재가 : " + str(price))
                    
                    
        elif status == "재매수대기상태" and rebuy == 0 :
            self.ui.plainTextEdit.appendPlainText("재매수 횟수 0 회, 거래종료 :" + name)
            self.dic[list_1[1]] = 2


                
        elif status == "재매수상태":
            if price < low * 0.99 :
                self.send_order('send_order', "0101", self.ui.account_number, 2, trcode, rebuy_count,  0 ,"03", "" )
                self.dic[list_1[0]] = "재매수대기상태"
                self.dic[list_1[4]] = price * buy_count
                self.ui.plainTextEdit.appendPlainText("하단가 1% 미만 지점 이탈 | 매도 :" + name)
            
            #보유중일때 현재가가 익절기준을 넘으면 매도   
            elif price >= initial + (initial * self.ui.take_profit/100):
                self.send_order('send_order', "0101", self.ui.account_number, 2, trcode, rebuy_count,  0 ,"03", "" )
                self.ui.plainTextEdit.appendPlainText("익절 퍼센트 초과 | 매도 :" + name)
                self.dic[list_1[0]] = "익절매도"
                
            
                

        
            

        



if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect() #연결
    

    
    
