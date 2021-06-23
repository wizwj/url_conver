import xlwt
from wechatsogou import *
import xlrd
from pywinauto import backend
from pywinauto.application import Application
import time
import win32clipboard
import xlutils
from xlutils.copy import copy

# 采集数据开始
def search_hot_article(type):
    wechats = WechatSogouAPI(captcha_break_time=1)
    relate_article = wechats.search_article(type)
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row0 = [u'标题', u'长网址', u'短网址']
    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i])
    j = 1
    for k in relate_article:
        print("来源公众号:{1}\t\t文章名称:{0}".format(k['article']['title'], k['article']['url']))
        sheet1.write(j, 0, k['article']['title'])
        sheet1.write(j, 1, k['article']['url'])
        j = j + 1
    f.save('data.xls')


def search_gzh_article(name):
    wechats = WechatSogouAPI(captcha_break_time=1)
    relate_article = wechats.search_article(name)
    for i in relate_article:
        print("来源公众号:{1}\t\t文章名称:{0}".format(i['article']['title'], i['gzh']['wechat_name']))

# 采集数据结束

#################################################################################################################################
# 转换网址开始##################################
# 读取excel数据
def read_excel():
    wb = xlrd.open_workbook(r'data.xls')
    print(wb.sheet_names())
    sheet1 = wb.sheet_by_index(0)
    rowNum = sheet1.nrows
    colNum = sheet1.ncols
    f = copy(wb)
    sheet2 = f.get_sheet('sheet1')
    for i in range(rowNum):
        s = sheet1.cell(i, 1).value
        if i != 0:
            d_cli = wechat_fuc(s)
            sheet2.write(i, 2, d_cli)
            f.save('data.xls')
            time.sleep(2)


# 微信操作
def wechat_fuc(text):
    app = Application(backend="uia").connect(title="文件传输助手", class_name="ChatWnd")
    win_main_Dialog = app.window(title="文件传输助手")
    win_main_Dialog.draw_outline(colour="red")
    chat_list = win_main_Dialog.child_window(title="输入", control_type="Edit")
    chat_list.type_keys(text)
    win_main_Dialog.child_window(title="发送(S)", control_type="Button").click_input()
    a = win_main_Dialog.child_window(title="消息", control_type="List")[-1]
    a.click_input(coords=(int(float(a.rectangle().width()) / 2), int(float(a.rectangle().height()) / 5)))
    time.sleep(2)
    app2 = Application(backend="uia").connect(title="微信", class_name="CefWebViewWnd")
    win_main_Dialog2 = app2.window(title="微信", class_name="CefWebViewWnd")
    try:
        win_main_Dialog2.child_window(title="复制链接地址", control_type="Button").wait('ready', timeout=15)
    except:
        win_main_Dialog2.child_window(title="关闭", control_type="Button").click_input()
        # text = "http://mp.weixin.qq.com/s?src=11&timestamp=1624360686&ver=3146&signature=4FAel0O93chG59nkIF64Ps8*swE7OJp7wQCg0dN7TzdKpH0EzkxWO320liP3FQWqL5yPc7eOtKYuEBkydPSYVa5AOzrvy0IslUcJG10gyrfwIH*b8inKBP7yKONAsN*E&new=1"
        wechat_fuc(text)
    else:
        win_main_Dialog2.child_window(title="复制链接地址", control_type="Button").click_input()
        win_main_Dialog2.child_window(title="关闭", control_type="Button").click_input()
    win32clipboard.OpenClipboard()
    data_cli = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return data_cli
if __name__ == '__main__':
    flag = 0
    if flag == 0:
        read_excel()
    else:
        search_hot_article(WechatSogouConst.hot_index.technology)