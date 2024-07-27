import time, win32con, win32api, win32gui
from io import BytesIO
import win32clipboard 
from PIL import Image
from datetime import datetime, timedelta
from pynput.keyboard import Key, Controller
import schedule
from bs4 import BeautifulSoup as bs
from pprint import pprint
import requests
import ssl
import datetime
import urllib
from bs4 import BeautifulSoup
import urllib.request as req
from datetime import date
from urllib.request import urlopen
import json



now = datetime.datetime.now()
today = date.today()
nowDate = now.strftime('%Y년 %m월 %d일 %H시 %M분 입니다.')


# 카톡창 이름
kakao_opentalk_name = '형섭'

# 기상청 API
url = "http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getUltraSrtFcst" + \
    "?serviceKey=Hr7ptUjmxsg2qquNmNN4c5nKyF%2FiQ6na7Y9lu3py3INO9mUeOlSR7hzQTN8DlUEssHGrMWySZcYH4apzSsylWtw%3D%3D" + \
    "&numOfRows=60&pageNo=1" + \
    "&base_date=2024" + today.strftime("%m%d") + "&base_time=0700" + "&nx=61&ny=129&dataType=JSON"
print(url)



# # 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None)
    # hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    #pressPaste() 오류코드
    time.sleep(0.01)
    returnPaste()

    print("sent") 

    returnPaste()

# 붙여넣기
def pressPaste():
    keyboard = Controller()

    keyboard.press(Key.ctrl)
    keyboard.press('v')
    time.sleep(0.01)
    keyboard.release(Key.ctrl)
    keyboard.release('v')

# 검색용 엔터
def returnPaste():
    keyboard = Controller()
    keyboard.press(Key.enter)
    time.sleep(0.01)
    keyboard.release(Key.enter)

# 채팅 보내는 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# 채팅방 열기
def open_chatroom(chatroom_name):
    # 채팅방 목록 검색
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)

    # Edit에 검색
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)
    SendReturn(hwndkakao_edit3)
    time.sleep(1)

# 클립보드에 저장하기
def send_to_clipboard(clip_type, data): 
    win32clipboard.OpenClipboard() 
    win32clipboard.EmptyClipboard() 
    win32clipboard.SetClipboardData(clip_type, data) 
    win32clipboard.CloseClipboard() 

ad = []
def weather():
    global ad
    answer = urlopen(url).read()
    data = json.loads(answer)
    rain = dict()
    for item in data["response"]["body"]["items"]["item"]:
        if item["category"] == "RN1":
            rain[item["fcstTime"]] = item["fcstValue"]
    for k, v in rain.items():
        print("{}시에 예상 강수량 {}".format(k, v))
        ad.append("{}시에 예상 강수량 {}".format(k, v))



def main():
    global ad
    open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    weather()
    ing_li = [word for word in ad if 'ing' in word]
    text = str(ing_li) #리스트 깔끔하게 출력
    print(text) 
    #text = "1"
    kakao_sendtext(kakao_opentalk_name, text)    # 메시지 전송


if __name__ == '__main__':
    main()
