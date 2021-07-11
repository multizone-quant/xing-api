# 선물/옵션 코드에 대하여 풀어서 돌려줌
# 참고 글
# https://agibbyeongari.tistory.com/897
# 2021 : 105R8000 : 1:선물, 05:미니선물 R:2021 8:월 A,B,C(10,11,12)

import json

# 최근 월 선물, mini_future 코드 얻는다.
# ex 101R8000
def get_code_info(code) :
    # 파생 종류
    code_1 = {'1' : 'future', '2' : 'call op', '3' : 'put op', '4' : 'spread'}
    # 파생 대상
    code_2_3 = {'01' : '', '04' : 'vix', '05' : 'mini', '06' : 'kosdak150 future', '07' : 'euro50 future', '08' : 'krx300 future', '09' : 'weekly'}
    # 년
    code_4 = {'Q':2020, 'R':2021, 'S':2022, 'T':2023, 'V':2024, 'W':2025, '6':2026, '7':2027, '8':2028, '9':2029, '0':2030, '1':2031, '2':2032, '3':2033, '4':2034, '5':2035}
    # 월
    code_5 = {'1':1, '2':2, '3':3, '4':4, '5':5, '6':6, '7':7, '8':8, '9':9, 'A':10, 'B':11, 'C':12}

    adding = {'future' : 'future', 'call op':'op', 'put op':'op'}

    info = {}
    val = code[0]
    info['type'] = code_1[val]

    weekly = 0
    val = code[1:3]
    if val == '09' : # weekly option은 구조가 틀리다.
        weekly = 1

    info['gubun'] = code_2_3[val]
    if info['gubun'] != '' :
        info['gubun'] += ' '
    # option, 선물 구분
    if info['type'] in adding :
        info['gubun'] += adding[info['type']]

    val = code[3]
    info['year'] = code_4[val]

    val = code[4]
    info['month'] = code_5[val]

    val = code[5:8]
    info['exe_price'] = int(val)

    # weekly option 구조는 일반 option/future와는 틀리다.
    if weekly :  # 20973442 
        val = code[3] # 3번째 칸이 월 (1-9,A,B,C)
        info['year'] = 0
        info['month'] = code_5[val]
        info['week'] = int(code[4]) # 4번째 칸은 몇 번째 주인지.
    else :
        info['week'] = 0

    return info

if __name__ == '__main__':
    
    # kospi200 선물
    code = '101R9000'
    info = get_code_info(code)
    print(code, info)

    # call option
    code = '201R8440'
    info = get_code_info(code)
    print(code, info)

    # weekly call option
    code = '20973442'
    info = get_code_info(code)
    print(code, info)

    # mini call option
    code = '205R8440'
    info = get_code_info(code)
    print(code, info)

    # put option
    code = '301R8440'
    info = get_code_info(code)
    print(code, info)

    # weekly call option
    code = '30973442'
    info = get_code_info(code)
    print(code, info)

    # mini put option
    code = '305R8440'
    info = get_code_info(code)
    print(code, info)

    # kospi200 미니 선물
    code = '105R8000'
    info = get_code_info(code)
    print(code, info)

    print('Trader')

