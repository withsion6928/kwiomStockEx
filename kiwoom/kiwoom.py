from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwom 클래스 입니다.")

        #####eventLoop 모듈
        self.login_event_loop = None
        self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        #######################

        ###스크린 번호 모음
        self.screen_my_info = "2000"



        #######변수모듈####
        self.account_num = None
        self.account_stock_dict = {}
        #################


        ## 계좌 관련 변수##

        self.use_money = 0
        self.use_money_percent = 0.01
        ####




        ########################
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info() #예수금
        self.detail_account_mystock() #계좌평가잔고내역요청
        self.not_cncs_ordr() #미체결


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)


    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))


        self.login_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num  = account_list.split(";")[0]

        print("나의 보유 계좌번호 %s" % self.account_num) ##8154121511

    def detail_account_info(self): ## tr  요청하는 부분
        print("예수금을 요청하는 부분")

        self.dynamicCall("SetInputValue(String, String)","계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0" ,self.screen_my_info)
        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):  ## tr  요청하는 부분
        print("게좌평가 잔고내역 요청")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)


        self.detail_account_info_event_loop.exec_()

    def not_cncs_ordr(self, sPrevNext="0"):  ## tr  요청하는 부분
        print("미체결주문내역확인")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opw10075", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext): ## tr 요청 후 결과 받는 부분
        '''
        tr요청을 받는 구역, 슬롯이다.
        :param sScrNo: 스크린번호
        :param sRQName: 내가 요청했을 때 지은 이름
        :param sTrCode: 요청id, tr코드
        :param sRecordName: 사용안함
        :param sPrevNext: 다음페이지 유무
        :return:
        '''
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0, "예수금")
            print("예수금[%s] " % int(deposit))


            self.use_money = int(deposit)  * self.use_money_percent
            self.use_money = self.use_money / 10


            chuguem = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액[%s]" % int(chuguem))

            d2chuguem = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0, "d+2출금가능금액")
            print("d+2출금가능금액[%s]" % int(d2chuguem))

            self.detail_account_info_event_loop.exit()

        if sRQName == "계좌평가잔고내역요청":
            total_buy_money =  self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money_result = int(total_buy_money)
            print("총매입금액[%s]" % total_buy_money_result)

            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)

            print("총수익률(%%) : %f" % (total_profit_loss_rate_result))

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]

                if code in self.account_dict:
                    pass
                else:
                    self.account_dic.update({code:{}})

                prdt__nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                prdt__nm = prdt__nm.strip()

                stock_bln = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                stock_bln = int(stock_bln.strip())

                buy_unpr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                buy_unpr = int(buy_unpr.strip())

                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                learn_rate = float(learn_rate.strip())

                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                current_price = int(current_price.strip())

                total_buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                total_buy_price = int(total_buy_price.strip())

                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명":prdt__nm})
                self.account_stock_dict[code].update({"보유수량":stock_bln})
                self.account_stock_dict[code].update({"매입가":buy_unpr})
                self.account_stock_dict[code].update({"수익률(%)":learn_rate})
                self.account_stock_dict[code].update({"현재가":current_price})
                self.account_stock_dict[code].update({"매입금액":total_buy_price})
                self.account_stock_dict[code].update({"매매가능수량":possible_quantity})

                cnt += 1
            print("계좌에 가지고 있는 종목 %s" % self.account_stock_dict)
            print("계좌에 가지고 있는 카운트 %s" % cnt)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            print("47강~ 48 강  참조해랏 :https://youtu.be/MlPfJwn_6rA")
