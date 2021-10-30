import os

from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import QTest
from config.kiwoomType import *
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwom 클래스 입니다.")
        self.realType = RealType()
        #####eventLoop 모듈
        self.login_event_loop = None
        self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        #######################

        ###스크린 번호 모음
        self.screen_start_stop_info = "1000"
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"
        self.screen_real_stock = "5000" ## 종목별로 할당할 스크린 번호
        self.screen_meme_stock = "6000"  ## 종목별로 할당할 주문용 스크린 번호

        #######변수모듈####
        self.account_num = None
        self.account_stock_dict = {}
        self.not_account_stock_dic = {}
        self.portfolio_stock_dict = {}
        #################
        ##종목분석용
        self.calcul_data = []


        ## 계좌 관련 변수##

        self.use_money = 0
        self.use_money_percent = 0.01
        ####

        ########################
        self.get_ocx_instance()
        self.event_slots()
        self.real_event_slots()


        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()  # 예수금
        self.detail_account_mystock()  # 계좌평가잔고내역요청
        #self.not_cncs_ordr() #미체결
        #self.calculator_fnc()  # 종목 분석용 임시용으로 실행


        self.read_code()   ####저장된종목을 불러온다.
        self.screen_number_settion() ##스크린 번호를 할당
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_info, '', self.realType.REALTYPE['장시작시간']['장운영구분'],'0')

        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]['스크린번호']
            fids = self.realType.REALTYPE['주식체결']['체결시간']
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_num, code, fids, '1')
            print("실시간 등록 코드 : %s,  스크린번호: %s, fid번호: %s" % (code, screen_num, fids))

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def real_event_slots(self):
        self.OnReceiveRealData.connect(self.realdata_slot)



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
        self.account_num = account_list.split(";")[0]

        print("나의 보유 계좌번호 %s" % self.account_num)  ##8154121511

    def detail_account_info(self):  ## tr  요청하는 부분
        print("예수금을 요청하는 부분")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info)
        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):  ## tr  요청하는 부분
        print("게좌평가 잔고내역 요청")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext,
                         self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def not_cncs_ordr(self, sPrevNext="0"):  ## tr  요청하는 부분
        print("미체결주문내역확인")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opw10075", sPrevNext,self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):  ## tr 요청 후 결과 받는 부분
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

            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 10

            chuguem = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액[%s]" % int(chuguem))

            d2chuguem = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0,
                                         "d+2출금가능금액")
            print("d+2출금가능금액[%s]" % int(d2chuguem))

            self.detail_account_info_event_loop.exit()

        if sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0,
                                               "총매입금액")
            total_buy_money_result = int(total_buy_money)
            print("총매입금액[%s]" % total_buy_money_result)

            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0,
                                                      "총수익률(%)")
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
                    self.account_dic.update({code: {}})

                prdt__nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                prdt__nm = prdt__nm.strip()

                stock_bln = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                stock_bln = int(stock_bln.strip())

                buy_unpr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                buy_unpr = int(buy_unpr.strip())

                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "수익률(%)")
                learn_rate = float(learn_rate.strip())

                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "현재가")
                current_price = int(current_price.strip())

                total_buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                   "매입금액")
                total_buy_price = int(total_buy_price.strip())

                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                     "매매가능수량")
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명": prdt__nm})
                self.account_stock_dict[code].update({"보유수량": stock_bln})
                self.account_stock_dict[code].update({"매입가": buy_unpr})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_buy_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt += 1
            print("계좌에 가지고 있는 종목 %s" % self.account_stock_dict)
            print("계좌에 가지고 있는 카운트 %s" % cnt)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            print("47강~ 48 강  참조해랏 :https://youtu.be/MlPfJwn_6rA")

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태")
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"주문구분")
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"주문미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"체결량")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = order_no.strip()
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip(' +').lstrp('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dic:
                    pass
                else:
                    self.not_account_stock_dic[order_no] = {}

                self.not_account_stock_dic[order_no].update({"종목코드": code})
                self.not_account_stock_dic[order_no].update({"종목명": code_nm})
                self.not_account_stock_dic[order_no].update({"주문번호": order_no})
                self.not_account_stock_dic[order_no].update({"주문상태": order_status})
                self.not_account_stock_dic[order_no].update({"주문수량": order_quantity})
                self.not_account_stock_dic[order_no].update({"주문가격": order_price})
                self.not_account_stock_dic[order_no].update({"주문구분": order_gubun})
                self.not_account_stock_dic[order_no].update({"미체결수량": not_quantity})
                self.not_account_stock_dic[order_no].update({"체결량": ok_quantity})
                print("미체결 종목 : %s" % self.not_account_stock_dic[order_no])

            self.detail_account_info_event_loop.exit()

        elif sRQName == "주식일봉차트조회":
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print("%s 일봉데이터 요청 " % code)

            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            print("데이터 일수 : %s" % cnt)
            #한번 조회 하면 600일치까지 일봉데이터를 받아올수 있다.
            for i in range(cnt):
                data = []
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"거래량")
                trading_value= self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"거래대금")
                date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"일자")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"저가")

                data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append("")

                self.calcul_data.append(data.copy())

            ##print(self.calcul_data)


            if sPrevNext == "2":
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print("총 일수 :%s" % len(self.calcul_data))
                #120일 이평선을 그릴만큼의 데이터가 있는지를 체크
                pass_success = False

                if self.calcul_data == None or len(self.calcul_data) < 120:
                    pass_success = False
                else:

                    total_price = 0
                    for value in self.calcul_data[:120]:
                        total_price += int(value[1])#120일간의 종가를 다 더함

                    moving_average_price = total_price / 120
                    bottom_stock_price = False
                    check_price = None
                    if int(self.calcul_data[0][7]) <= moving_average_price and moving_average_price <= int(self.calcul_data[0][6]): #저가보다 높고 고가보다 낮고
                        print("오늘 주가 120이평선에 걸쳐있는것 확인")
                        bottom_stock_price = True
                        check_price = int(self.calcul_data[0][6])
                        prev_price = None # 과거의 120일선 위에 있는 일봉 저가
                    if bottom_stock_price == True:
                        moving_average_price_prev = 0
                        price_top_moving = False

                        idx = 1
                        while True:
                            if len(self.calcul_data[idx:]) < 120 :  #120일이치가 있는지 계속 확인
                                print("120일치가 없음!")
                                break

                            total_price = 0
                            for value in self.calcul_data[idx:120+idx]:
                                total_price += int(value[1])
                                moving_average_price_prev = total_price / 120

                            if moving_average_price_prev <= int (self.calcul_data[idx][6]) and idx <= 20: ##20일전에 고가가
                                print("20일동안 주가가 120일 이평선과 같거나 위에 있으면 조건 통과 못함 ")
                                price_top_moving = False
                                break
                            elif int(self.calcul_data[idx][7]) > moving_average_price_prev and idx > 20:
                                print("120일 이평선 위에 있는 일봉 확인됨  ")
                                price_top_moving = True
                                prev_price = int(self.calcul_data[idx][7])
                                break
                            idx += 1

                        #해당 부분 이평선이 가장최근일자의 이평선 가격보다 낮은지 확인
                        if price_top_moving == True:
                            if moving_average_price > moving_average_price_prev and check_price > prev_price:
                                print("포착된 이평선의 가격이 오늘자(최근일자) 이평선가격보다 낮은것으로 확인됨 ")
                                print("포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮은지 확인됨")
                                pass_success = True

                if pass_success == True:
                    print("조건부 통과됨 ")
                    code_nm = self.dynamicCall("GetMasterCodeName(QString", code)

                    f = open("files/condition_stock.txt", "a", encoding="utf8")
                    f.write("$s\t%s\t%s\n" % (code, code_nm, str(self.calcul_data[0][1])))
                    f.close()
                elif pass_success == False:
                    print("조건부 통과 못함")


                self.calculator_event_loop.exit()

    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]
        return code_list

    def calculator_fnc(self):
        code_list = self.get_code_list_by_market("10")
        print("코스닥 갯수 %s" % len(code_list))

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisConnectRealData(QString)", self.screen_calculation_stock)
            print("%s /%s : KOSDAK Stock Code : %s is updating.... " % (idx + 1, len(code_list), code))
            self.day_kiwoom_db(code=code)

    def day_kiwoom_db(self, code=None, date=None, sPrevNext="0"):

        QTest.qWait(3600)

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가코드", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식일봉차트조회", "opt10081", sPrevNext,
                         self.screen_calculation_stock)

        self.calculator_event_loop.exec_()


    def read_code(self):
        if os.path.exists("files/condition_stock.txt"):
            f = open("files/condition_stock.txt","r",encoding="utf8")
            lines = f.readLine()
            for line in lines:
                if line != "":
                    ls = line.split("\t")

                    stock_code = ls[0]
                    stock_name = ls[1]
                    stock_price = int(ls[2].split("\n"))
                    stock_price = abs(stock_price)
                    self.portfolio_stock_dict.update({stock_code:{"종목명" : stock_name, "현재가" : stock_price}})

            f.close()
            print(self.portfolio_stock_dict)


    def screen_number_settion(self):
        screen_overwrite = []

        for code in self.account_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)


        for order_number in self.not_account_stock_dic.keys():
            code = self.not_account_stock_dic[order_number]['종목코드']

            if code not in screen_overwrite:
                screen_overwrite.append(order_number)

        for code in self.portfolio_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)

        cnt = 0
        for code in screen_overwrite:

            temp_screen = int(self.screen_real_stock)
            meme_screen = int(self.screen_meme_stock)

            if (cnt % 50) == 0:
                temp_screen += 1
                self.screen_real_stock = str(temp_screen)

            if (cnt % 50 ) == 0:
                meme_screen += 1
                self.screen_meme_stock = str(meme_screen)

            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({"스크린번호": str(self.screen_real_stock)})
                self.portfolio_stock_dict[code].update({"주문용스크린번호": str(self.screen_meme_stock)})
            elif code not in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict.update({"스크린번호": str(self.screen_real_stock), "주문용스크린번호": str(self.screen_meme_stock)})

            cnt += 1
        print(self.portfolio_stock_dict)


    def realdata_slot(self, sCode, sRealType, sRealData):

        if sRealType == "장시작시간":
            fid = self.realType.REALTYPE[sRealType]['장운영구분']
            value = self.dynamicCall("GetCommRealData(QString, int)", sCode, fid)

            if value == "0":
                print("장 시작 전")
            elif value == "3":
                print("장 시작")
            elif value == "2":
                print("장 종료, 동시호가로 넘어감")
            elif value == "4":
                print("3시30분 장 종료")

        elif sRealType == "주식체결":
            print(sCode)


