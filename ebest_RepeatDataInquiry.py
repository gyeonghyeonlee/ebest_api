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

# 이벤트 처리를 위한 클래스 정의
class XAQueryEventHandlerT8430:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT8430.query_state = 1


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
# T8430
# ----------------------------------------------------------------------------
# XAQuery 클래스에 대한 인스턴스 생성, 해당 인스턴스에 t8430 TR 코드에 해당하는 Res 파일 등록
instXAQueryT8430 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT8430)
instXAQueryT8430.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t8430.res"

# 입력 블록명 : t8430InBlock, 입력에 사용되는 필드명 : gubun, 0은 전체 시장, 1은 코스피(유가증권시장), 2는 코스닥(코스닥시장)
instXAQueryT8430.SetFieldData("t8430InBlock", "gubun", 0, 1)

# 입력 데이터 설정 후 서버에 TR 요청 -> 이벤트가 발생할 때까지 대기
instXAQueryT8430.Request(0)

while XAQueryEventHandlerT8430.query_state == 0:
    pythoncom.PumpWaitingMessages()

# 반복 데이터의 총 개수 구하려면 GetBlockCount 메소드 사용
count = instXAQueryT8430.GetBlockCount("t8430OutBlock")
# 각 인덱스에 해당하는 데이터를 가져올 때는 앞서 살펴본 GetFieldData 메소드 사용
# 단일 데이터와 달리 세 번째 파라미터에 반복 데이터의 인덱스 지정
for i in range(5):
    hname = instXAQueryT8430.GetFieldData("t8430OutBlock", "hname", i) # hname : 종목명
    shcode = instXAQueryT8430.GetFieldData("t8430OutBlock", "shcode", i) # shcode : 단축코드
    expcode = instXAQueryT8430.GetFieldData("t8430OutBlock", "expcode", i) # expcode : 확장코드
    etfgubun = instXAQueryT8430.GetFieldData("t8430OutBlock", "etfgubun", i) # etfgubun : ETF 구분
    print(i, hname, shcode, expcode, etfgubun)