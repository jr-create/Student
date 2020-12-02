#
# author: jr
#
import time

import pikepdf
import wx
import wx.grid
import os
from PyPDF2 import PdfFileMerger, PdfFileReader
import GUI_view.word_pdf as wp

def nowtime():  # 显示当前时间
    return time.strftime('北京时间%Y/%m/%d %A %H:%M:%S', time.localtime(time.time()))
def getFileName(filedir):  # 获取文件夹中后缀为PDF的文件
    file_list = [os.path.join(root, filespath) \
                 for root, dirs, files in os.walk(filedir) \
                 for filespath in files \
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []

def pdf_Crack(startFile,endFile):
    print(startFile)
    print(endFile)
    with pikepdf.open(startFile) as pdf:
        pdf.save(endFile)
        return True
def pdf_merge(listBox, file, verdict_head, verdict_tail, verdict_REPEAT):  # pdf合并 参数  文件列表和文件路径
    try:
        outputPage = 0
        merger = PdfFileMerger()
        for i in listBox:
            start_page = 0
            input1 = open(i, 'rb')
            print('文件打开成功')
            input = PdfFileReader(input1)   # 读取源PDF文件
            print('文件读取成功')
            if input.isEncrypted:
                message = i + '有密码'+'是否破解'
                # wx.MessageBox(message)    # 弹出提示框
                dlg = wx.MessageDialog(None,message, "pdf_Crack", wx.YES_NO | wx.ICON_QUESTION)
                if dlg.ShowModal() == wx.ID_YES:
                    pdf_Crack(i,file)       #PDF解密
                    return True
                else:
                    return False
            else:
                pageCount = input.getNumPages()  # 获得源PDF文件中页面总数
                print(pageCount)
                outputPage += pageCount
                if verdict_head == True:    #去头
                    start_page = 1
                if verdict_tail == True:    #去尾
                    pageCount -= 1
                if verdict_REPEAT == False: #去重复
                    verdict_head = False
                    verdict_tail = False
                merger.append(fileobj=input1, pages=(start_page, pageCount))
                print('写入成功')
        output = open(file, 'wb')
        merger.write(output)
        print("总页数{}".format(outputPage))
        return True
    except:
        message = '文件合并失败'
        wx.MessageBox(message)  # 弹出提示框


def method1(self, panel):
    self.sampleList = []
    #列表
    self.listBox = wx.ListBox(panel, -1, (100, 20), (300, 120), self.sampleList, wx.LB_MULTIPLE)
    #文本
    self.label_list = wx.StaticText(panel, label="文件列表:", pos=(100, 0))
    self.label_line = wx.StaticText(panel, label="最终路径:", pos=(45, 153))
    self.label_time = wx.StaticText(panel, label=nowtime(), pos=(200, 0))
    #文本框
    self.file = wx.TextCtrl(panel, value='', pos=(100, 150), size=(230, 25), style=wx.TE_LEFT)
    #按钮
    bt_confirm = wx.Button(panel, label='确定', pos=(105, 200))  # 创建“确定”按钮
    bt_cancel = wx.Button(panel, label='取消', pos=(195, 200))  # 创建“取消”按钮

    self.checkbox1 = wx.CheckBox(panel, -1, "去头", (402, 20), (200, 20))  # 复选框
    self.checkbox2 = wx.CheckBox(panel, -1, "去尾", (402, 50), (200, 20))  # 复选框
    self.checkbox3 = wx.CheckBox(panel, -1, "重复", (402, 80), (200, 20))  # 复选框

    bt_confirm.Bind(wx.EVT_BUTTON, self.OnButton_confirm)       # “确定”按钮事件
    bt_cancel.Bind(wx.EVT_BUTTON, self.OnButton_cancel)         # “取消”按钮事件
    self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.listCtrlSelectFunc)  # 列表双击取消
    self.checkbox1.Bind(wx.EVT_CHECKBOX, self.One_Play)         # 复选框事件
    self.checkbox2.Bind(wx.EVT_CHECKBOX, self.Two_Play)         # 复选框事件

def frame_config(frame,self,parent,id, tit):
    frame.__init__(self, parent, id, title=tit, size=(500, 350)
                      , style=wx.SYSTEM_MENU | wx.MINIMIZE_BOX | wx.CLOSE_BOX | wx.CAPTION
                      )
    self.CreateStatusBar()  # 创建位于窗口的底部的状态栏
    self.SetBackgroundColour('white')
    self.Center()  # 在显示屏中心显示
class MyFrame(wx.Frame):
    def __init__(self, parent, id):
        frame_config(wx.Frame,self, parent,id, "PDF合并")
        # 创建一个分隔窗
        # splitter = wx.SplitterWindow(self, -1)
        # 创建子面板
        # self.leftpanel = wx.Panel(self)
        # rightpanel = wx.Panel(splitter)
        self.panel = wx.Panel(self)  # 创建面板
        method1(self, self.panel)
        # 分隔窗口
        # splitter.SplitVertically(self.panel,self.leftpanel,  500)
        # splitter.SetMinimumPaneSize(80)  # 设置最小窗口尺寸，这里指的是左窗口尺寸

        self.InitUI()  # 菜单

    def InitUI(self):
        menuBar = wx.MenuBar()  # 创建一个菜单栏
        fileMenu = wx.Menu()  # 创建一个菜单 1
        # wx.ID_OPEN 和 wx.ID_EXIT是wxWidgets提供的标准ID
        menuOpenfile = fileMenu.Append(wx.ID_FILE, "&Open_file", "Open a 文件")
        menuOpenfolder = fileMenu.Append(wx.ID_FILE1, "&Open_folder", "Open a 文件夹")
        menuAbout = fileMenu.Append(wx.ID_ABOUT, "&About", "Information about this program")  # (ID, 项目名称, 状态栏信息)
        newItem = fileMenu.Append(wx.ID_ADD, "&WORD->PDF", "WORD->PDF")
        fileMenu.AppendSeparator()  # 分割线
        menuExit = fileMenu.Append(wx.ID_EXIT, "&Exit", "Terminate the program")  # (ID, 项目名称, 状态栏信息)

        menuBar.Append(fileMenu, title='File')  # 在菜单栏中添加filemenu菜单
        self.SetMenuBar(menuBar)  # 在frame中添加菜单栏
        # 绑定事件处理
        self.Bind(wx.EVT_MENU, self.OnButton_selct2, menuOpenfile)
        self.Bind(wx.EVT_MENU, self.OnButton_selct1, menuOpenfolder)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.Bind(wx.EVT_MENU, self.menuHandler, newItem)
        self.Show(True)

    def OnAbout(self, e):
        # 创建一个带"OK"按钮的对话框。wx.OK是wxWidgets提供的标准ID
        dlg = wx.MessageDialog(self, "A Small Editor.", "About Sample Editor", wx.OK)  # 语法是(self, 内容, 标题, ID)
        dlg.ShowModal()  # 显示对话框
        dlg.Destroy()  # 当结束之后关闭对话框

    def OnExit(self, event):        #列表Exit事件处理
        self.Close(True)            # 关闭整个frame
    def menuHandler(self, event):   #列表WORD->PDF事件处理
        self.Close(True)            # 关闭整个frame
        frame = wp.windows2(parent=None, id=-1)  # 实例MyFrame类，并传递参数
        # frame.ShowFullScreen(False)
        frame.Show()                # 显示WORD->PDF窗口

    def One_Play(self, event):
        print("本次选择了吗：", self.checkbox1.GetValue())

    def Two_Play(self, event):
        print("本次选择了吗：", self.checkbox2.GetValue())

    def listCtrlSelectFunc(self, event):  # 列表点击事件
        i = self.listBox.GetSelections()
        print(int(i[0]))
        self.listBox.Delete(int(i[0]))
        # print(self.listBox.GetSelections())

    def OnButton_confirm(self, event):  # 确定点击事件
        verdict_REPEAT = self.checkbox3.GetValue()
        verdict_head = self.checkbox1.GetValue()
        verdict_tail = self.checkbox2.GetValue()
        if pdf_merge(self.listBox.GetStrings(), self.file.GetValue(), verdict_head, verdict_tail, verdict_REPEAT):
            message = '文件合并成功'
            wx.MessageBox(message)  # 弹出提示框

    def OnButton_cancel(self, event):  # 关闭点击事件
        self.listBox.Clear()

    def OnButton_selct1(self, event):  # 选择文件夹点击事件
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            print(dlg.GetPath())  # 文件夹路径
            self.file.SetValue(getFileName(dlg.GetPath())[0])
            for i in getFileName(dlg.GetPath()):
                self.listBox.Append(i)

        dlg.Destroy()

    def OnButton_selct2(self, event):  # 选择文件点击事件
        # wildcard = u"Python 文件 (*.py)|*.py|" \
        #            u"文本文件 (*.txt)|*.txt|" \
        #            u"pdf文件 (*.pdf)|*.pdf|" \
        #            "Egg file (*.egg)|*.egg|" \
        #            "All files (*.*)|*.*"
        wildcard = u"pdf文件 (*.pdf)|*.pdf|" \
                   "All files (*.*)|*.*"
        dlg = wx.FileDialog(self, message=u"选择文件",
                            defaultDir=os.getcwd(),
                            defaultFile="",
                            wildcard=wildcard,
                            style=wx.FD_OPEN)

        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()  # 返回一个list，如[u'E:\\test_python\\Demo\\ColourDialog.py', u'E:\\test_python\\Demo\\DirDialog.py']
            self.listBox.Append(paths)
            self.file.SetValue(dlg.GetPaths()[0])
        else:
            print('打开失败')
        dlg.Destroy()
def word_pdf():
    print('dsf')

if __name__ == '__main__':
    app = wx.App()  # 初始化
    frame = MyFrame(parent=None, id=-1)  # 实例MyFrame类，并传递参数
    # frame.ShowFullScreen(False)
    frame.Show()  # 显示窗口
    app.MainLoop()  # 调用主循环方法
