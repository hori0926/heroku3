from flask import Flask, request, abort

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
)
import pandas as pd
import datetime
import openpyxl

app = Flask(__name__)

line_bot_api = LineBotApi('WXxwsnhHb8dZ9UCl7mgV1vx4js8887TMjYwiDZh+jXLapUOi/7isvtcOmZP9m+prCkf1A4hpCMlBfNaRPTAiCzid+mPvNMhcpnfHPHz98cxR2bTTJCjghADZnyk7JOjWFYJVC14reFW39QdF6p607gdB04t89/1O/w1cDnyilFU=')
handler = WebhookHandler('38e51470f5e59138220e544832393d63')

datenum = {}

@app.route("/")
def test():
    return "ok"

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    d_today = datetime.date.today()
    try:
        print(datenum[d_today])
    except:
        datenum[d_today]=0
        print(datenum[d_today])
    data = pd.read_excel('ttldatabase.xlsx',index_col=0)
    data.loc[d_today.strftime("%Y-%m-%d"),datenum[d_today]]=event.message.text
    data.to_excel('ttldatabase.xlsx')
    datenum[d_today] +=1
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=event.message.text + "を本日の入場者に記録しました"))



if __name__ == "__main__":
    app.run("0.0.0.0", debug=True)