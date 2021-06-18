# -*- coding: utf-8 -*-
import win32com.client
import pythoncom
import sys
import time

# ======================================================
# 위치가 틀리다면 수정하여야 하는 부분
# ======================================================
XING_PATH = "C:\\eBEST\\xingAPI"

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
    id = ""                         # 본인의 ID로 수정
    passwd = ""
    account_number = ''
    account_pwd = "0000"   
# ======================================================
# 수정하여야 하는 부분 끝
# ======================================================

TODAY = time.strftime("%Y%m%d")
TODAY_TIME = time.strftime("%H%M%S")
TODAY_S = time.strftime("%Y-%m-%d")


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


def login(server, id, pwd, cer_pwd, acc, acc_pwd) :
    instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)

    instXASession.ConnectServer(server, 20001)
    instXASession.Login(id, passwd, cert_passwd, 0, 0)
    while XASessionEventHandler.login_state == 0:
        pythoncom.PumpWaitingMessages()

    login = XASessionEventHandler.login_state
    return login

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

# 코스피200 지수선물 마스터 조회
def get_8432() :
    tr_code = 't8432'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\"+tr_code+".res"
    query.SetFieldData(INBLOCK, "gubun", 0, "1") # 코스피200 지수선물 마스터 조회
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':tr_code+' msg error'}}]

    result1 = []
    nCount = query.GetBlockCount(OUTBLOCK)

    for i in range(nCount):
        sh_code = query.GetFieldData(OUTBLOCK, "shcode", i).strip() #종목코드
        sh_name = query.GetFieldData(OUTBLOCK, "hname", i).strip() #종목명
        lst = {'code':sh_code, 'name':sh_name}
        result1.append(lst)
    return [result1]

# 선물/옵션 보유종목 조회
def get_0441(accno, passwd) :
    tr_code = 't0441'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\"+tr_code+".res"
    query.SetFieldData(INBLOCK, "accno", 0, accno) # 
    query.SetFieldData(INBLOCK, "passwd", 0, passwd) # 
    query.SetFieldData(INBLOCK, "cts_expcode", 0, '')#CTS_종목번호)
    query.SetFieldData(INBLOCK, "cts_medocd", 0, '')#CTS_매매구분)
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':tr_code+' msg error'}}]

    result = []
    result2 = []

    tr_profit = query.GetFieldData(OUTBLOCK, "tdtsunik", 0).strip() #매매손익합
    total = query.GetFieldData(OUTBLOCK, "tappamt", 0).strip() #평가금액  # 매수/매도 구분이 없이 합임. 실제로는 매수/매도 구분하여 +-필요
    profit = query.GetFieldData(OUTBLOCK, "tsunik", 0).strip() #평가손익
    lst = {'tr_profit':tr_profit, 'total':total, 'profit':profit}
    result.append(lst)

    nCount = query.GetBlockCount(OUTBLOCK1)
    for i in range(nCount):
        code = query.GetFieldData(OUTBLOCK1, "expcode", i).strip() #종목코드
        buy_sell = query.GetFieldData(OUTBLOCK1, "medosu", i).strip() #종목명
        qty = query.GetFieldData(OUTBLOCK1, "jqty", i).strip() #잔고수량
        orderable_qty = query.GetFieldData(OUTBLOCK1, "cqty", i).strip() #청산가능수량
        buy_sell_add = query.GetFieldData(OUTBLOCK1, "medocd", i).strip() #매매구분
        profit = query.GetFieldData(OUTBLOCK1, "dtsunik", i).strip() #매매손익
        price = query.GetFieldData(OUTBLOCK1, "price", i).strip() #현재가
        lst = {'code':code, 'buy_sell':buy_sell, 'qty':float(qty), 'orderable_qty':float(orderable_qty), 'buy_sell_add':buy_sell_add, 'profit':float(profit), 'price':float(price) }
        result2.append(lst)

    result.append(result2)
    return result

# 파생 코드 조회
# gubun : 미니선물(MF), MO(미니옵션), WK(위클리옵션), SF(코스닥150선물)
def get_8435(gubun) :
    tr_code = 't8435'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\"+tr_code+".res"
    query.SetFieldData(INBLOCK, "gubun", 0, gubun) # 코스피200 지수선물 마스터 조회
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':tr_code+' msg error'}}]

    result1 = []
    nCount = query.GetBlockCount(OUTBLOCK)

    for i in range(nCount):
        sh_code = query.GetFieldData(OUTBLOCK, "shcode", i).strip() #종목코드
        exp_code = query.GetFieldData(OUTBLOCK, "expcode", i).strip() #확장코드
        rec_price = query.GetFieldData(OUTBLOCK, "recprice", i).strip() #기준가
        lst = {'code':sh_code, 'expcode':exp_code, 'recprice':rec_price}
        print(lst)
        result1.append(lst)
    return [result1]

# call/put info
def get_2301(yyyymm) :
    tr_code = 't2301'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\"+tr_code+".res"
    query.SetFieldData(INBLOCK, "yyyymm", 0, yyyymm) # 월물
    query.SetFieldData(INBLOCK, "gubun", 0, "G") # 구분 mini(M), 정규(G)
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':tr_code+' msg error'}}]

    result1 = []
    result2 = []
    # for call
    nCount = query.GetBlockCount(OUTBLOCK1)
    for i in range(nCount) :
        act_price = query.GetFieldData(OUTBLOCK1, "actprice", i).strip() #행사가
        code = query.GetFieldData(OUTBLOCK1, "optcode", i).strip() #option code
        delta = query.GetFieldData(OUTBLOCK1, "delt", i).strip() #delta
        gamma = query.GetFieldData(OUTBLOCK1, "gama", i).strip() #gamma
        ceta = query.GetFieldData(OUTBLOCK1, "ceta", i).strip() #ceta
        vega = query.GetFieldData(OUTBLOCK1, "vega", i).strip() #vega

        lst = {'act_price':float(act_price), 'code':code, 'delta':float(delta), 'gamma':float(gamma), 'ceta':float(ceta), 'vega':float(vega)}
        result1.append(lst)
    # for put
    nCount = query.GetBlockCount(OUTBLOCK2)
    for i in range(nCount) :
        act_price = query.GetFieldData(OUTBLOCK2, "actprice", i).strip() #행사가
        code = query.GetFieldData(OUTBLOCK2, "optcode", i).strip() #option code
        delta = query.GetFieldData(OUTBLOCK2, "delt", i).strip() #option code
        gamma = query.GetFieldData(OUTBLOCK2, "gama", i).strip() #gamma
        ceta = query.GetFieldData(OUTBLOCK2, "ceta", i).strip() #ceta
        vega = query.GetFieldData(OUTBLOCK2, "vega", i).strip() #vega

        lst = {'act_price':float(act_price), 'code':code, 'delta':float(delta), 'gamma':float(gamma), 'ceta':float(ceta), 'vega':float(vega)}
        result2.append(lst)

    return [result1, result2]

# 선물/옵션 현재가
def get_t2101(code) :
    tr_code = 't2101'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\"+tr_code+".res"
    query.SetFieldData(INBLOCK, "focode", 0, code)
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return [{'error':{'message':tr_code+' msg error'}}]

    result1 = []
    basis = query.GetFieldData(OUTBLOCK, "basis", 0).strip() #basis
    price = query.GetFieldData(OUTBLOCK, "price", 0).strip() #가격
    sign = query.GetFieldData(OUTBLOCK, "sign", 0).strip() #up/down
    theoryprice = query.GetFieldData(OUTBLOCK, "theoryprice", 0).strip() #theoryprice
    delta = query.GetFieldData(OUTBLOCK, "delt", 0).strip() #delta
    gamma = query.GetFieldData(OUTBLOCK, "gama", 0).strip() #gamma
    ceta = query.GetFieldData(OUTBLOCK, "ceta", 0).strip() #ceta
    vega = query.GetFieldData(OUTBLOCK, "vega", 0).strip() #vega
    lst = {'code':code, 'basis':basis, 'price':float(price), 'sign':float(sign), 'delta':float(delta), 'gamma':float(gamma), 'ceta':float(ceta), 'vega':float(vega)}
    result1.append(lst)
    return [result1]


class aboutQueryEvents():
    count_t0441 = 0
    count_t2101 = 0
    count_t0167 = 0
    count_t8432 = 0

    def OnReceiveData(self, szTrCode):
        global insQuery_t0167, insQuery_t2101, insQuery_t8432, insQuery_CFOAT00100
        global insReal_FC0, insReal_IJ_, insReal_JIF
        global current_price, open_price, futcode, open_interest
        global TODAY, SERVERHHMM
        global insReal_C01, account_number, futcode, acctnm, acctnm_r

        if szTrCode == "t8432":     # 종목마스터 - 최근월물종목코드 수신
            futcode = insQuery_t8432.GetFieldData("t8432OutBlock", "shcode", 0)     # 최근월물
            futcode2 = insQuery_t8432.GetFieldData("t8432OutBlock", "shcode", 1)    # 차근월물
            msg = "%s]futcode 1st:%s 2nd:%s " % (szTrCode, futcode, futcode2)
            print(msg)
            time.sleep(1.1)

        else:
            msg = "로그인 에러 로그인정보, 계좌정보 확인요망 %s %s" % (szCode, szMsg)
            mywindow.te_msg.append(msg)
            print(msg)

if __name__ == "__main__":
    
    print('\nebest testing')


    ret = login(server_add, id, passwd, cert_passwd, account_number, account_pwd)
    if ret == 0 :
        print('fail to login')
        quit(0)
    time.sleep(1)


    # 선물/옵션 계좌 조회
    print('-- account balance -- ')
    cur_hold = get_0441(account_number, account_pwd)
    each = cur_hold[0]
    for each in cur_hold[1] :
        print (each['code'], each['qty'], each['buy_sell'], each['buy_sell_add'], each['price'])
    print ( 'total_profit :', each['profit'])

    # 7월물 전광판, 7월물 옵션 정보 조회
    print('-- 옵션 전광판 --')
    all_option = get_2301('202107') # 전광판
    if 'error' in all_option[0] :  # 오류인 경우에 set 형식으로 돌아옴 'error'
        print (all_option[0]['error']['message'])
    else :
        print(' call ')
        print (all_option[0])
        print('\n put ')
        print (all_option[1])

    # 선물/옵션현재가조회
    print('-- option info --')
    rest = get_t2101('201R7435') # 
    print (rest[0])

    # 미니선물 코드 조회
    print('-- mini future code --')
    mini_future = get_8435('MF') # 미니선물
    if 'error' in mini_future[0] :  # 오류인 경우에 set 형식으로 돌아옴 'error'
        print (mini_future[0]['error']['message'])
    for each in mini_future[0] :
        print (each)

    print('-- future code  --')
    rest = get_8432() # 선물종목마스터조회 (최근월물종목코드 가져오기)
    if 'error' in rest[0] :  # 오류인 경우에 set 형식으로 돌아옴 'error'
        print (rest[0]['error']['message'])
    for each in rest[0] :
        print (each)

