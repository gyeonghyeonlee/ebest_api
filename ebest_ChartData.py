import win32com.client
import pythoncom

class XASessionEventHandler:
    login_state = 0

    def OnLogin(self, code, msg):
        if code == "0000":
            print("로그인 성공")
            XASessionEventHandler.login_state = 1
        else:
            print("로그인 실패")

class XAQueryEventHandlerT8413:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT8413.query_state = 1


# ----------------------------------------------------------------------------
# login
# ----------------------------------------------------------------------------
id = "이베스트 투자증권 아이디"
passwd = "이베스트 투자증권 비밀번호"
cert_passwd = "공동인증서 비밀번호"

instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)
instXASession.ConnectServer("hts.ebestsec.co.kr", 20001)
instXASession.Login(id, passwd, cert_passwd, 0, 0)

while XASessionEventHandler.login_state == 0:
    pythoncom.PumpWaitingMessages()


# ----------------------------------------------------------------------------
# T8413
# ----------------------------------------------------------------------------
instXAQueryT8413 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT8413)
instXAQueryT8413.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t8413.res"

# 이베스트 투자증권 종목에 대해 2021-05-06~2021-05-14 기간의 일봉 데이터 받아오기
instXAQueryT8413.SetFieldData("t8413InBlock", "shcode", 0, "078020")
instXAQueryT8413.SetFieldData("t8413InBlock", "gubun", 0, "2")
instXAQueryT8413.SetFieldData("t8413InBlock", "sdate", 0, "20210506")
instXAQueryT8413.SetFieldData("t8413InBlock", "edate", 0, "20210514")
instXAQueryT8413.SetFieldData("t8413InBlock", "comp_yn", 0, "N")

instXAQueryT8413.Request(0)

while XAQueryEventHandlerT8413.query_state == 0:
    pythoncom.PumpWaitingMessages()

# t8413OutBlock1은 Occurs 속성을 가진 반복 데이터
count = instXAQueryT8413.GetBlockCount("t8413OutBlock1")
# 2021-05-06~2021-05-14 기간에는 총 7일의 주식 거래일 존재 -> 데이터의 총 개수는 7개
# 0~6까지의 각 인덱스는 특정 거래일 의미, 해당 거래일의 날짜, 시가, 고가, 저가, 종가 구하기
for i in range(count):
    date = instXAQueryT8413.GetFieldData("t8413OutBlock1", "date", i) # 날짜
    open = instXAQueryT8413.GetFieldData("t8413OutBlock1", "open", i) # 시가
    high = instXAQueryT8413.GetFieldData("t8413OutBlock1", "high", i) # 고가
    low = instXAQueryT8413.GetFieldData("t8413OutBlock1", "low", i) # 저가
    close = instXAQueryT8413.GetFieldData("t8413OutBlock1", "close", i) # 종가
    print(date, open, high, low, close)