from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwom 클래스 입니다.")

        #####eventLoop 모듈
        self.login_event_loop = None
        #######################

        #######변수모듈####
        self.account_num = None
        #################


        #################이벤트루프 모듈#################
        self.detail_account_info_event_loop = None


        ########################
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info() #예수금


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
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0" , "2000")
        self.detail_account_info_event_loop = QEventLoop()
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

            chuguem = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액[%s]" % int(chuguem))

            d2chuguem = self.dynamicCall("GetCommData(String, String, String, String)", sTrCode, sRQName, 0, "d+2출금가능금액")
            print("d+2출금가능금액[%s]" % int(d2chuguem))

            self.detail_account_info_event_loop.exit()

