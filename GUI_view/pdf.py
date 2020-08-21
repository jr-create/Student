import wx
import os
from PyPDF2 import PdfFileMerger,PdfFileReader
def getFileName(filedir):#获取文件夹中后缀为PDF的文件
    file_list = [os.path.join(root, filespath) \
                 for root, dirs, files in os.walk(filedir) \
                 for filespath in files \
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []
class MyFrame(wx.Frame):
    def __init__(self, parent, id):
        wx.Frame.__init__(self,  parent, id,  title="PDF合并", size=(500,  300))
        panel = wx.Panel(self)  # 创建面板
        self.sampleList = []
        self.listBox = wx.ListBox(panel, -1, (100, 20), (300, 120), self.sampleList,wx.LB_SINGLE)
        self.label_user = wx.StaticText(panel, label="文件列表:", pos=(100, 0))
        self.label_user = wx.StaticText(panel, label="最终路径:", pos=(45, 153))
        self.file = wx.TextCtrl(panel,value='', pos=(100, 150), size=(230, 25), style=wx.TE_LEFT)
        bt_selct1 = wx.Button(panel, label='选择文件夹', pos=(10, 60))  # 创建“确定”按钮
        bt_selct2 = wx.Button(panel, label='选择文件', pos=(10, 20))  # 创建“确定”按钮
        bt_confirm = wx.Button(panel, label='确定', pos=(105, 200))   # 创建“确定”按钮
        bt_cancel = wx.Button(panel, label='取消', pos=(195, 200))    # 创建“取消”按钮
        bt_confirm.Bind(wx.EVT_BUTTON, self.OnButton_confirm)
        bt_selct1.Bind(wx.EVT_BUTTON,self.OnButton_selct1)
        bt_selct2.Bind(wx.EVT_BUTTON, self.OnButton_selct2)
        bt_cancel.Bind(wx.EVT_BUTTON,self.OnButton_cancel)
        self.listBox.Bind(wx.EVT_LISTBOX,self.listCtrlSelectFunc)
    def listCtrlSelectFunc(self,event):     #列表点击事件
        i=self.listBox.GetSelections()
        print(int(i[0]))
        self.listBox.Delete(int(i[0]))
        # print(self.listBox.GetSelections())
    def OnButton_confirm(self,event):       #确定点击事件
        try:
            outputPage = 0
            merger = PdfFileMerger()
            for i in self.listBox.GetStrings():
                input1 = open(i, 'rb')
                input = PdfFileReader(input1)  # 读取源PDF文件
                if input.isEncrypted:
                    message = i+'有密码'
                    wx.MessageBox(message)  # 弹出提示框
                    raise Exception("抛出一个密码异常")
                else:
                    pageCount = input.getNumPages()  # 获得源PDF文件中页面总数
                    outputPage += pageCount
                    merger.append(fileobj=input1, pages=(0, pageCount - 1))
            print(outputPage)
            output = open(self.file.GetValue(), 'wb')
            merger.write(output)
            message = '文件合并成功'
            wx.MessageBox(message)  # 弹出提示框
        except:
            message='文件合并失败'
            wx.MessageBox(message)  # 弹出提示框
    def OnButton_cancel(self, event):           #关闭点击事件
       self.listBox.Clear()
    def OnButton_selct1(self, event):           #选择文件夹点击事件
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            print(dlg.GetPath())  # 文件夹路径
            self.file.SetValue(dlg.GetPath())
            for i in getFileName(dlg.GetPath()):
                self.listBox.Append(i)
        dlg.Destroy()
    def OnButton_selct2(self, event):           #选择文件点击事件
        # wildcard = u"Python 文件 (*.py)|*.py|" \
        #            u"文本文件 (*.txt)|*.txt|" \
        #            u"pdf文件 (*.pdf)|*.pdf|" \
        #            "Egg file (*.egg)|*.egg|" \
        #            "All files (*.*)|*.*"
        wildcard=u"pdf文件 (*.pdf)|*.pdf|"\
                   "All files (*.*)|*.*"
        dlg = wx.FileDialog(self, message=u"选择文件",
                            defaultDir=os.getcwd(),
                            defaultFile="",
                            wildcard=wildcard,
                            style=wx.FD_OPEN)

        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()  # 返回一个list，如[u'E:\\test_python\\Demo\\ColourDialog.py', u'E:\\test_python\\Demo\\DirDialog.py']
            self.listBox.Append(paths)
            self.file.SetValue(dlg.GetPath())
        else:
            print('打开失败')
        dlg.Destroy()
if __name__ == '__main__':
    app = wx.App()                      # 初始化
    frame = MyFrame(parent=None, id=-1)  # 实例MyFrame类，并传递参数
    frame.Show()                        # 显示窗口
    app.MainLoop()                      # 调用主循环方法