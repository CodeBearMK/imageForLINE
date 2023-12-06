# 引用flask(網頁處理及網頁伺服器相關套件)
from flask import Flask, request, abort
# 引用linebot sdk
from linebot import (
    LineBotApi, WebhookHandler
)
# 引用linebot.exceptions類的InvalidSignatureError
from linebot.exceptions import (
    InvalidSignatureError
)
# 引用linebot.models類所提供的Message Type
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage, ImageMessage, ImageSendMessage
)
# 引用內建套件
import tempfile, os
# 引用imgur api套件(上傳圖檔用)
from imgurpython import ImgurClient
# 引用opencv影像處理套件
import cv2
# 環境變數套件
from dotenv import load_dotenv 
# 引用列印套件
import win32print
# 引用windows ui套件
import win32ui
# 引用圖片處理套件
from PIL import Image,ImageWin

load_dotenv('config.env')

# 啟用flask
app = Flask(__name__)

# 建立linebot物件
line_bot_api = LineBotApi(os.getenv('LINE_TOKEN'))

# line事件處理器
handler = WebhookHandler(os.getenv('LINE_CHANNEL_SECRET'))

# 圖檔暫存路徑
static_tmp_path = os.path.join(os.path.dirname(__file__),'static','tmp')

# 驗證簽章(Webhook名稱:lineImgnPrint)
@app.route("/lineImgnPrint",methods=['POST'])
def lineImgnPrint():
    # 取得 X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # 取得 request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body,signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)
    return 'OK'

# 新增MessageEvent handle事件,對象為圖片檔
@handler.add(MessageEvent,message=(TextMessage,ImageMessage))
def handle_message(event):
    # 如果對方傳的是圖片的話
    if isinstance(event.message,ImageMessage):
        # 取得該筆訊息ID
        message_content = line_bot_api.get_message_content(event.message.id)
        # 暫存原始圖檔(只是把圖片暫存到指定位置,之後影像處理用)
        with tempfile.NamedTemporaryFile(dir=static_tmp_path,prefix='jpg-',delete=False) as tf:
            for chunk in message_content.iter_content():
                tf.write(chunk)
            tempfile_path = tf.name

        dist_path = tempfile_path+'.jpg'
        dist_name = os.path.basename(dist_path)
        os.rename(tempfile_path,dist_path)
        path = os.path.join('static','tmp',dist_name)
        # 二值化影像處理
        path2, th = imgProcess(path)
        # 移除原圖
        os.remove(path)
        try:
            # 建立Imgur Client以上傳圖片之用(LINE回覆圖片用)
            client = ImgurClient(os.getenv('IMGUR_CLIENT_ID'),os.getenv('IMGUR_CLIENT_SECRET'),os.getenv('IMGUR_ACCESS_TOKEN'),os.getenv('IMGUR_REFRESH_TOKEN'))
            # 存影像處理完的圖檔
            cv2.imwrite(path2, th)
            # 上傳圖檔到Imgur
            response = client.upload_from_path(path2,anon=False)
            # 建立圖片訊息物件
            messages = ImageSendMessage(response['link'],response['link'])
            # 回覆訊息給傳訊者
            test = line_bot_api.reply_message(event.reply_token,messages)
            #client.delete_image(response['id'])
            # 列印處理完的圖檔
            print(path2)
            # 移除處理完的圖檔
            os.remove(path2)
            return '200 OK'
        except:
            # 傳送失敗時回覆傳送失敗給傳訊者
            line_bot_api.reply_message(event.reply_token,TextSendMessage(text='上傳失敗'))
        return '400 Bad Request'
    
# 二值化影像處理
def imgProcess(path):
    # 檔名設定為jpg
    img = cv2.imread(path,0)
    # 影像處理
    th = cv2.adaptiveThreshold(img,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,67,13)
    # 影像處理完存檔
    with tempfile.NamedTemporaryFile(dir=static_tmp_path,prefix='jpg-',delete=False) as tf:
        tf.write(th)
        tempfile_path2 = tf.name

    dist_path2 = tempfile_path2+'.jpg'
    dist_name2 = os.path.basename(dist_path2)
    os.rename(tempfile_path2,dist_path2)
    path2 = os.path.join('static','tmp',dist_name2)
    # 這邊只有回傳存檔路徑跟圖檔(這邊圖檔不能當作LINE傳檔格式)
    return path2, th


# 列印服務
def print(path2):
    #
    # Constants for GetDeviceCaps
    #
    #
    # HORZRES / VERTRES = printable area
    #
    HORZRES = 8
    VERTRES = 10
    #
    # LOGPIXELS = dots per inch
    #
    LOGPIXELSX = 88
    LOGPIXELSY = 90
    #
    # PHYSICALWIDTH/HEIGHT = total area
    #
    PHYSICALWIDTH = 110
    PHYSICALHEIGHT = 111
    #
    # PHYSICALOFFSETX/Y = left / top margin
    #
    PHYSICALOFFSETX = 112
    PHYSICALOFFSETY = 113
    
    printer_name = win32print.GetDefaultPrinter()
    file_name = path2
    
    #
    # You can only write a Device-independent bitmap
    # directly to a Windows device context; therefore
    # we need (for ease) to use the Python Imaging
    # Library to manipulate the image.
    #
    # Create a device context from a named printer
    # and assess the printable size of the paper.
    #
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)
    printable_area = hDC.GetDeviceCaps(HORZRES),hDC.GetDeviceCaps(VERTRES)
    printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH),hDC.GetDeviceCaps(PHYSICALHEIGHT)
    printer_margins = hDC.GetDeviceCaps(PHYSICALOFFSETX),hDC.GetDeviceCaps(PHYSICALOFFSETY)
    
    #
    # Open the image,rotate it if it's wider than
    # it is high,and work out how much to multiply
    # each pixel by to get it as big as possible on
    # the page without distorting.
    #
    bmp = Image.open(file_name)
    if bmp.size[0] > bmp.size[1]:
        bmp = bmp.rotate(90)
    
    ratios = [1.0 * printable_area[0] / bmp.size[0],1.0 * printable_area[1] / bmp.size[1]]
    scale = min (ratios)
    
    #
    # Start the print job,and draw the bitmap to
    # the printer device at the scaled size.
    #
    hDC.StartDoc(file_name)
    hDC.StartPage()
    
    dib = ImageWin.Dib(bmp)
    scaled_width,scaled_height = [int (scale * i) for i in bmp.size]
    x1 = int ((printer_size[0] - scaled_width) / 2)
    y1 = int ((printer_size[1] - scaled_height) / 2)
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height
    dib.draw(hDC.GetHandleOutput (),(x1,y1,x2,y2))
    
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()


if __name__ == "__main__":
    app.run()