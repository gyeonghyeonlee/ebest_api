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
class XAQueryEventHandlerT1102:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT1102.query_state = 1

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
# t1102
# ----------------------------------------------------------------------------

# XAQuery 클래스의 인스턴스 생성
instXAQueryT1102 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT1102)

# Res 파일 경로
instXAQueryT1102.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1102.res"

# 입력 데이터 설정 : XAQuery 인스턴스를 통해 SetFieldData 메소드 호출 후 적절한 인자 값 지정
# SetFieldData의 첫 번째 파라미터는 블록명, 두 번째 파라미터는 필드명
# 세 번째 파라미터는 단일 데이터 조회 시 0 지정, 네 번째 파라미터에는 필드에 해당하는 입력 값
instXAQueryT1102.SetFieldData("t1102InBlock", "shcode", 0, "078020")

# 요청할 작업에 필요한 데이터 입력이 완료되면 Request 메소드 호출하여 입력 데이터 서버로 전송
instXAQueryT1102.Request(0)

# 서버에 TR 요청 후 해당 작업이 완료됐다는 이벤트를 받을 때까지 프로그램이 종료하지 않고 대기해야 함
while XAQueryEventHandlerT1102.query_state == 0:
    pythoncom.PumpWaitingMessages()

# 작업 완료 이벤트를 받은 후에는 GetFieldData 메소드를 사용해 t1102OutBlock 블록명과 원하는 필드명 입력하여 데이터 가져옴
name = instXAQueryT1102.GetFieldData("t1102OutBlock", "hname", 0) # hname : 한글종목명
price = instXAQueryT1102.GetFieldData("t1102OutBlock", "price", 0) # price : 종목의 현재가
print(name)
print(price)