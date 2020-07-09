# -*- coding: utf-8 -*-

# date : 2020/07/08
# xing api sample
#  - login
#  - 잔고조회 : T0424
#  - 주문조회 : T0425
#  - Q검색리스트 : T1826
#  - Q검색 : T1825
#
# 보다 자세한 내용을 아래 tistory 참고
# https://money-expert.tistory.com/14
# https://money-expert.tistory.com/17

import win32com.client
import pythoncom
import sys
import time
from PyQt5 import QtWidgets
from PyQt5 import QtGui
from PyQt5 import QtCore
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox

# ======================================================
# 위치가 틀리다면 수정하여야 하는 부분
# ======================================================
XING_PATH = "C:\\eBEST\\xingAPI\\"
# 위치가 틀리다면 수정하여야 하는 부분 끝
# ======================================================

# ======================================================
# 수정하여야 하는 부분
# ======================================================
server_add = "hts.ebestsec.co.kr"
id = "ebest id"
passwd = "로그인 암호"
cert_passwd = "공인인증서 암호"
account_number = "주식 계좌번호" 
account_pwd = "주식계좌 암호"   
if 1 : #모의투자
    server_add = "demo.ebestsec.co.kr"
    passwd = "모의투자 사이트 로그인암호"
    account_number = '모의주식 계좌번호'
    account_pwd = "모의 주식계좌 암호"           
# ======================================================
# 수정하여야 하는 부분 끝
# ======================================================

TODAY = time.strftime("%Y%m%d")
TODAY_TIME = time.strftime("%H%M%S")
TODAY_S = time.strftime("%Y-%m-%d")

class Form(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.ui = uic.loadUi("xing_sample_ui.ui", self)
        #init
        self.query_list = []

    def clear_message(self) :
        self.ui.listWidget_msg.clear()
    def show_message(self, pr) :
        self.ui.listWidget_msg.addItem(pr)
        self.ui.listWidget_msg.scrollToBottom()

    # T0424 잔고 받기
    def Balance_0424(self) :
        ret, bals = get_balance('all')  # 모든 종목 정보를 얻는다. 특정 종목을 원하면 해당하는 코드입력
        if ret >= 0 :
            pr = '=== 잔고 ==='
            self.show_message(pr)
            pr = ' code  balance '
            self.show_message(pr)
            pr = '--------------'
            self.show_message(pr)

            for bal in bals[0] :        
                pr = bal['code'] + ' ' + str(bal['total'])
                self.show_message(pr)

    # T0424 잔고 받기
    def OrderResults_0425(self) :
        self.clear_message()
        ordered = order_status_tr(kind='0', code='all') # kind = '0'(전체), '1'(체결), '2'(미체결)
        if 'error' in ordered[0] : # 오류
            self.show_message("0425 : error returned")
            return

        # orders[1] : 주문 내역
        order_num = ordered[2][0]['total']
        if order_num > 0 :
            pr = '  주문결과 '
            self.show_message(pr)
            pr = '--------------------------'
            self.show_message(pr)
            pr = '총 주문수: ' +  str(order_num)
            self.show_message(pr)

            # 취소 주문 : price == 0
            # 미체결 : 'executed_volume' == 0
            # 체결 :  'executed_volume' == volume
            for order in ordered[0] :
                if order['price'] == 0 : #취소주문
                    pr = '취소  : ' + order['market']
                    self.show_message(pr)
                elif order['executed_volume'] == 0 : #미체결
                    pr = '미체결: ' + order['market']
                    self.show_message(pr)
                elif order['executed_volume'] ==  order['volume']: #미체결
                    pr = '체결  : ' + order['side'] + ' ' + order['market'] + ' 가격: ' + str(order['price']) + ' 수량: ' +str(order['volume'])
                    self.show_message(pr)
                else :
                    pr = 'unknown'
                    self.show_message(pr)
                print(order)

            # orders[1] : 체결에 대한 총괄 정보
            # {'ord_total':ord_total, 'ord_fee':ord_fee, 'ord_tax':ord_tax})
            pr = '--------------------------'
            self.show_message(pr)            
            ord_summary = ordered[1][0]
            pr = '주문총수량 : ' + str(ord_summary['ord_total'])
            self.show_message(pr)
            pr = '주문수수료 : ' + str(ord_summary['ord_fee'])
            self.show_message(pr)
            pr = '주문세금   : ' + str(ord_summary['ord_tax'])
            self.show_message(pr)
            pr = '--------------------------'
            self.show_message(pr)

    # t1825 Q 검색 리스트 받기
    def Q_Query_1825(self) :
        if self.query_list == [] :
            self.show_message('press 1826 first')
            return

        for lst in self.query_list :
            time.sleep(1)
            pr = "\n=== " + lst[1] +  " ==="
            self.show_message(pr)
            res = get_q_query(lst[0])
            if 'error' in res[0] :
                self.show_message (res[0]['error']['message'])
            else :
                if len(res) > 1 :
                    pr = "total : " + str(res[0][0]['total'])                    
                    self.show_message (pr)
                    cnt = 0
                    for itm in res[1] :
                        pr = itm['code'] + ' ' + itm['name'] +' ' + str(itm['price']) + ' ' + str(itm['gubun'])
                        self.show_message(pr)
                        if cnt > 10 :
                            break
                        cnt+=1
                else :
                    self.show_message ("total : 0")

    # t18256 Q 검색 결과 받기
    def Q_List_1826(self) :
        rest = get_q_query_list('0')
        if 'error' in rest[0] :
            self.show_message(rest[0]['error']['message'])
        self.query_list = rest[0] 

        for query in self.query_list :
            pr = query[0] + ' ' + query[1]
            self.show_message(pr)

class XASessionEventHandler:
    login_state = 0

    def OnLogin(self, code, msg):
        print('on login start')
        if code == "0000":
            print("login succ")
            XASessionEventHandler.login_state = 1
        else:
            XASessionEventHandler.login_state = -1
            print("login fail")

def wait_for_event(code) :
    while XAQueryEventHandler.query_state == 0:
        pythoncom.PumpWaitingMessages()

    if XAQueryEventHandler.query_code != code :
        print('diff code : wish(', code,')', XAQueryEventHandler.query_code)
        return 0
    XAQueryEventHandler.query_state = 0
    XAQueryEventHandler.query_code = ''
    return 1

class XAQueryEventHandler:
    query_state = 0
    query_code = ''
    T1102_query_state = 0
    T8413_query_state = 0

    def OnReceiveData(self, code):
#        print('OnRecv', code)
        XAQueryEventHandler.query_code = code
        XAQueryEventHandler.query_state = 1


def login(server, id, pwd, cer_pwd, acc, acc_pwd) :
    instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)

    instXASession.ConnectServer(server, 20001)
    instXASession.Login(id, passwd, cert_passwd, 0, 0)
    while XASessionEventHandler.login_state == 0:
        pythoncom.PumpWaitingMessages()

    login = XASessionEventHandler.login_state
    return login

def get_balance(ticker) :
    time.sleep(0.2)
    tr_code = 't0424'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = XING_PATH+"Res\\"+tr_code+".res"

    query.SetFieldData(INBLOCK, "accno", 0, account_number)#계좌번호)
    query.SetFieldData(INBLOCK, "passwd", 0, account_pwd) #비밀번호)
    query.SetFieldData(INBLOCK, "prcgb", 0, '1')#단가구분)
    query.SetFieldData(INBLOCK, "chegb", 0, '0')#체결구분)
    query.SetFieldData(INBLOCK, "dangb", 0, '0')#단일가구분)
    query.SetFieldData(INBLOCK, "charge", 0, '1')#제비용포함여부)
    query.SetFieldData(INBLOCK, "cts_expcode", 0, '')#CTS_종목번호)
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        err_msg = {'error':{'message':'Not respond msg'}}
        # all이면 리스트로, 아니면 set로
        if ticker == 'all' :
            return -1, [err_msg]
        else :
            return -1, err_msg
    if 0 :
        result = []
        nCount = query.GetBlockCount(OUTBLOCK)
        for i in range(nCount):
            cur_asset = int(query.GetFieldData(OUTBLOCK, "sunamt", i).strip()) #추정순자산
            profit = int(query.GetFieldData(OUTBLOCK, "dtsunik", i).strip()) #실현손익
            org_inv = int(query.GetFieldData(OUTBLOCK, "mamt", i).strip()) #매입금액
            est_amount = int(query.GetFieldData(OUTBLOCK, "tappamt", i).strip()) #평가금액
            est_profit = int(query.GetFieldData(OUTBLOCK, "tdtsunik", i).strip()) #평가손익

            lst = [cur_asset, profit, org_inv, est_amount, est_profit]
            result.append(lst)

    result = []
    nCount = query.GetBlockCount(OUTBLOCK1)
    bal = {'code':ticker, 'total':0, 'orderable':0}
    stock_code = ticker
    for i in range(nCount):
        stock_code = query.GetFieldData(OUTBLOCK1, "expcode", i).strip()
        if stock_code == ticker or ticker == 'all' :
            balance = int(query.GetFieldData(OUTBLOCK1, "janqty", i).strip()) #잔고수량
            orderable = int(query.GetFieldData(OUTBLOCK1, "mdposqt", i).strip()) #잔고수량
            bal = {'code':stock_code, 'total':balance, 'orderable':orderable}
            result.append(bal)

    if len(result) == 0 : # nothing
        result.append({'code':ticker, 'total':0, 'orderable':0})
    # all이면 리스트로, 아니면 set로
    if ticker == 'all' :
        return 1, [result]
    else :
        return 1, bal

# 주식 미체결 결과  t0425  # 미체결 :'2' 체결: '1' 전체:'0' 
def order_status_tr(kind='2', code='all', cmd_cont='') :
    '''
    주식 미체결
    '''
    time.sleep(0.2)        
    tr_code = 't0425'
    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = XING_PATH+"\\Res\\"+tr_code+".res"

    MYNAME = tr_code
    INBLOCK = "%sInBlock" % MYNAME
    INBLOCK1 = "%sInBlock1" % MYNAME
    OUTBLOCK = "%sOutBlock" % MYNAME
    OUTBLOCK1 = "%sOutBlock1" % MYNAME
    OUTBLOCK2 = "%sOutBlock2" % MYNAME

    query.SetFieldData(INBLOCK, "accno", 0, account_number)
    query.SetFieldData(INBLOCK, "passwd", 0, account_pwd)
    code_in = code
    if code == 'all' :
        code_in = ''
    query.SetFieldData(INBLOCK, "expcode", 0, code_in) # 종목번호 or blank(all)
    query.SetFieldData(INBLOCK, "chegb", 0, kind) # 미체결 :'?' 체결: '1' 전체:'0'
    query.SetFieldData(INBLOCK, "medosu", 0, '0') #매매구분)
    query.SetFieldData(INBLOCK, "sortgb", 0, '2') #정렬순서)
    query.SetFieldData(INBLOCK, "cts_ordno", 0, cmd_cont) #주문번호)
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':'Not respond msg'}}]

    result1 = []
    nCount = query.GetBlockCount(OUTBLOCK)
    ord_total = 0
    for i in range(nCount):
        ord_total = int(query.GetFieldData(OUTBLOCK, "tcheqty", i).strip()) #총체결수량
        ord_fee = int(query.GetFieldData(OUTBLOCK, "cmss", i).strip()) #추정수수료
        ord_tax = int(query.GetFieldData(OUTBLOCK, "tax", i).strip()) #추정세금
        result1.append({'ord_total':ord_total, 'ord_fee':ord_fee, 'ord_tax':ord_tax})

    result2 = []
    comp_qty = 0
    ord_no = ''
    last_order_num = ''
    nCount = query.GetBlockCount(OUTBLOCK1)
    for i in range(nCount):
        ord_no = query.GetFieldData(OUTBLOCK1, "ordno", i).strip() #주문번호
        # long 값으로 return하지만 그냥 string으로 사용함
        last_order_num = ord_no
        ord_code = query.GetFieldData(OUTBLOCK1, "expcode", i).strip() #종목번호
        if ord_code != code and code != 'all':
            continue
        ord_name = query.GetFieldData(OUTBLOCK1, "hname", i).strip() #종목명
        ord_side = query.GetFieldData(OUTBLOCK1, "medosu", i).strip() #구분
        org_qty = int(query.GetFieldData(OUTBLOCK1, "qty", i).strip()) #주문수량
        ord_price = int(query.GetFieldData(OUTBLOCK1, "price", i).strip()) #주문가격
        done_price = int(query.GetFieldData(OUTBLOCK1, "cheprice", i).strip()) #주문가격
        done_qty = int(query.GetFieldData(OUTBLOCK1, "cheqty", i).strip()) #체결수량
        comp_qty += 1 
        ord_time = query.GetFieldData(OUTBLOCK1, "ordtime", i).strip() #주문시간
        side_type = 'ask'
        if ord_side == '매수' : # 매수
            side_type = 'bid'

        order = {'time':ord_time, 'market':ord_code, 'uuid':ord_no, 'side':side_type, 'price':ord_price, 'executed_price':done_price, 'volume':org_qty, 'executed_volume':done_qty}
        result2.append(order)

    res = []
    res.append(result2)
    res.append(result1)
    if nCount < 100 :
        last_order_num = ''    # 체결 내역을 모두 다 읽어들였다.
    res.append([{'cont':last_order_num, 'total':nCount}])
    return res

# q_code : 검색하고자하는 q-query 번호
# gubun : 구분 (0(전체), 1(코스피), 2(코스닥))
def get_q_query(q_code, gubun='0') :
    tr_code = 't1825'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = XING_PATH+"\\Res\\"+tr_code+".res"

    query.SetFieldData(INBLOCK, "search_cd", 0, str(q_code))#q code
    query.SetFieldData(INBLOCK, "gubun", 0, str(gubun)) #구분
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':tr_code+' msg error'}}]

    result1 = []
    nCount = query.GetBlockCount(OUTBLOCK)
    res_cnt = 0
    for i in range(nCount):
        res_cnt = int(query.GetFieldData(OUTBLOCK, "JongCnt", i).strip()) #검색종목수

        lst = {'total':res_cnt}
        result1.append(lst)

    result2 = []
    for i in range(res_cnt):
        sh_code = query.GetFieldData(OUTBLOCK1, "shcode", i).strip() #종목코드
        sh_name = query.GetFieldData(OUTBLOCK1, "hname", i).strip() #종목명
        cur_gubun = query.GetFieldData(OUTBLOCK1, "sign", i).strip() #전일대비구분
        consec_bong = int(query.GetFieldData(OUTBLOCK1, "signcnt", i).strip()) #연속봉수
        cur_price = int(query.GetFieldData(OUTBLOCK1, "close", i).strip()) #현재가
        change = int(query.GetFieldData(OUTBLOCK1, "change", i).strip()) # 전일대비
        diff = query.GetFieldData(OUTBLOCK1, "diff", i).strip() # 등락율
        cur_vol = int(query.GetFieldData(OUTBLOCK1, "volume", i).strip()) #거래량
        vol_rate = query.GetFieldData(OUTBLOCK1, "volumerate", i).strip() # 거래량전일대비율
        lst = {'code':sh_code, 'name':sh_name, 'gubun':cur_gubun, 'consec_bong':consec_bong, 'price':cur_price, 'change':change, 'diff':diff, 'qty':cur_vol, 'qty_rate':vol_rate}
        result2.append(lst)

    res = []
    res.append(result1)
    res.append(result2)
    return res

# q_code : 검색하고자하는 q-query 번호
# gubun : 구분 (0(전체), 1(코스피), 2(코스닥))
def get_q_query_list(gubun) :
    tr_code = 't1826'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = XING_PATH+"Res\\"+tr_code+".res"

    query.SetFieldData(INBLOCK, "search_gb", 0, str(gubun)) #구분
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':tr_code+' msg error'}}]

    result1 = []
    nCount = query.GetBlockCount(OUTBLOCK)
    for i in range(nCount):
        res_code = query.GetFieldData(OUTBLOCK, "search_cd", i).strip() #검색코드
        res_name = query.GetFieldData(OUTBLOCK, "search_nm", i).strip() #검색명

        lst = [res_code, res_name]
        result1.append(lst)
    return [result1]

if __name__ == "__main__":
    
    print('\nebest testing')
    server_add = "hts.ebestsec.co.kr"
    
    ret = login(server_add, id, passwd, cert_passwd, account_number, account_pwd)
    if ret != -1 :
        time.sleep(1)

        # ======================================================
        # 수정할 부분 
        # GUI로 확인하고 싶으면 1로 변경
        # ======================================================
        USING_GUI = 1  # GUI로 확인

        # 수정하여야 하는 부분 끝
        # ======================================================

        if USING_GUI : # widget을 사용하는 경우
            app = QtWidgets.QApplication(sys.argv)
            WIDGET = Form()
            WIDGET.show()
            app.exec_()
    else :
        print('fail to login')


    
