import json
from datetime import datetime
from flask import Flask, request, make_response
from slack_sdk import WebClient
import win32com.client
import requests
 
token = "token"
app = Flask(__name__)
client = WebClient(token)

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# 현재가 객체 구하기
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
objStockMst.BlockRequest()
 
# 현재가 통신 및 통신 에러 처리 
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()
 
# 현재가 정보 조회
code = objStockMst.GetHeaderValue(0)  #종목코드
name= objStockMst.GetHeaderValue(1)  # 종목명
time= objStockMst.GetHeaderValue(4)  # 시간
cprice= objStockMst.GetHeaderValue(11) # 종가
diff= objStockMst.GetHeaderValue(12)  # 대비
open= objStockMst.GetHeaderValue(13)  # 시가
high= objStockMst.GetHeaderValue(14)  # 고가
low= objStockMst.GetHeaderValue(15)   # 저가
offer = objStockMst.GetHeaderValue(16)  #매도호가
bid = objStockMst.GetHeaderValue(17)   #매수호가
vol= objStockMst.GetHeaderValue(18)   #거래량
vol_value= objStockMst.GetHeaderValue(19)  #거래대금
 
# 예상 체결관련 정보
exFlag = objStockMst.GetHeaderValue(58) #예상체결가 구분 플래그
exPrice = objStockMst.GetHeaderValue(55) #예상체결가
exDiff = objStockMst.GetHeaderValue(56) #예상체결가 전일대비
exVol = objStockMst.GetHeaderValue(57) #예상체결수량
 
 
#print("코드", code)
#print("이름", name)
#print("시간", time)
#print("종가", cprice)
#print("대비", diff)
#print("시가", open)
#print("고가", high)
#print("저가", low)
#print("매도호가", offer)
#print("매수호가", bid)
#print("거래량", vol)
#print("거래대금", vol_value)
 
 
if (exFlag == ord('0')):
    print("장 구분값: 동시호가와 장중 이외의 시간")
elif (exFlag == ord('1')) :
    print("장 구분값: 동시호가 시간")
elif (exFlag == ord('2')):
    print("장 구분값: 장중 또는 장종료")
 
#print("예상체결가 대비 수량")
#print("예상체결가", exPrice)
#print("예상체결가 대비", exDiff)
#print("예상체결수량", exVol)

#def get_day_of_week(): #년/월/일 함수
#    weekday_list = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']
 
#    weekday = weekday_list[datetime.today().weekday()]
#    date = datetime.today().strftime("%Y년 %m월 %d일")
#    result = '{}({})'.format(date, weekday)
#    return result
 
#def get_time(): #시/분/초 함수
#    return datetime.today().strftime("%H시 %M분 %S초")
 
def get_answer(text):
    trim_text = text.replace(" ", "")
 
    answer_dict = {
        '메뉴목록': '종목, 현재가, 거래량, 종가',
        '종목' : str(name),
        '현재가': '{} 삼성전자 현재가 : {}'.format(time,offer),
        '종가': '{} 삼성전자 종가 : {}'.format(time,cprice),
        '거래량': '{} 삼성전자 거래량 : {}'.format(time,vol),
    }
 
    if trim_text == '' or None:
        return "알 수 없는 질의입니다. 답변을 드릴 수 없습니다."
    elif trim_text in answer_dict.keys():
        return answer_dict[trim_text]
    else:
        for key in answer_dict.keys():
            if key.find(trim_text) != -1:
                return "연관 단어 [" + key + "]에 대한 답변입니다.\n" + answer_dict[key]
 
        for key in answer_dict.keys():
            if answer_dict[key].find(text[1:]) != -1:
                return "질문과 가장 유사한 질문 [" + key + "]에 대한 답변이에요.\n"+ answer_dict[key]
 
    return text + "은(는) 없는 질문입니다."
 
 
def event_handler(event_type, slack_event):
    channel = slack_event["event"]["channel"]
    string_slack_event = str(slack_event)
 
    if string_slack_event.find("{'type': 'user', 'user_id': ") != -1:
        try:
            if event_type == 'app_mention':
                user_query = slack_event['event']['blocks'][0]['elements'][0]['elements'][1]['text']
                answer = get_answer(user_query)
                result = client.chat_postMessage(channel=channel,
                                                 text=answer)
            return make_response("ok", 200, )
        except IndexError:
            pass
 
    message = "[%s] cannot find event handler" % event_type
 
    return make_response(message, 200, {"X-Slack-No-Retry": 1})
 
 
@app.route('/', methods=['POST'])
def hello_there():
    slack_event = json.loads(request.data)
    if "challenge" in slack_event:
        return make_response(slack_event["challenge"], 200, {"content_type": "application/json"})
 
    if "event" in slack_event:
        event_type = slack_event["event"]["type"]
        return event_handler(event_type, slack_event)
    return make_response("There are no slack request events", 404, {"X-Slack-No-Retry": 1})
 
 
if __name__ == '__main__':
    app.run(debug=True, port=5000)
