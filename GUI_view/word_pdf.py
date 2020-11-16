import os

import wx
import GUI_view.pdf as pdf
from win32com.client import constants, gencache
def createPdf(wordPaths, pdfPath):
    """
    word转pdf
    :param wordPath: word文件路径
    :param pdfPath:  生成pdf文件路径
    """
    for wordPath in wordPaths:
        word = gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(wordPath, ReadOnly=1)
        pdfPath=wordPath[:-5]+'.pdf'
        doc.ExportAsFixedFormat(pdfPath,
                                constants.wdExportFormatPDF,
                                Item=constants.wdExportDocumentWithMarkup,
                                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
        word.Quit(constants.wdDoNotSaveChanges)
    message = 'word->pdf成功'
    wx.MessageBox(message)  # 弹出提示框
class windows2 (wx.Frame) :
    def __init__(self, parent, id):
        pdf.frame_config(wx.Frame,self,parent,id,'WORD转PDF')
        panel = wx.Panel(self)
        self.sampleList = []
        self.listBox = wx.ListBox(panel, -1, (100, 20), (300, 120), self.sampleList, wx.LB_MULTIPLE)
        # listBox = wx.ListBox(panel, -1, (100, 20), (300, 120), self.sampleList, wx.LB_MULTIPLE)
        label_list = wx.StaticText(panel, label="文件列表:", pos=(100, 0))
        label_line = wx.StaticText(panel, label="最终路径:", pos=(45, 153))
        bt_confirm = wx.Button(panel, label='确定', pos=(105, 200))  # 创建“确定”按钮
        bt_cancel = wx.Button(panel, label='取消', pos=(195, 200))  # 创建“取消”按钮
        self.file = wx.TextCtrl(panel, value='', pos=(100, 150), size=(300, 25), style=wx.TE_LEFT)
        bt_confirm.Bind(wx.EVT_BUTTON, self.OnButton_confirm)
        bt_cancel.Bind(wx.EVT_BUTTON, self.OnButton_cancel)
        self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.listCtrlSelectFunc)  # 双击取消
        pdf.MyFrame.InitUI(self)
    def OnAbout(self, e):
        # 创建一个带"OK"按钮的对话框。wx.OK是wxWidgets提供的标准ID
        dlg = wx.MessageDialog(self, "A Small Editor.", "About Sample Editor", wx.OK)  # 语法是(self, 内容, 标题, ID)
        dlg.ShowModal()  # 显示对话框
        dlg.Destroy()  # 当结束之后关闭对话框

    def OnExit(self, event):
        self.Close(True)  # 关闭整个frame
    def menuHandler(self, event):
        print('djf')

    def listCtrlSelectFunc(self, event):  # 列表点击事件
        i = self.listBox.GetSelections()
        print(int(i[0]))
        self.listBox.Delete(int(i[0]))
        # print(self.listBox.GetSelections())
    def OnButton_confirm(self, event):  # 确定点击事件
        if self.listBox.GetCount()==0:
            message = '列表为空'
            wx.MessageBox(message)  # 弹出提示框
        print('列表'+self.listBox.GetStrings()[0])
        print(self.file.GetValue())
        #起始文件路径
        createPdf(self.listBox.GetStrings(), self.file.GetValue())
    def OnButton_cancel(self, event):  # 关闭点击事件
        self.listBox.Clear()

    def OnButton_selct1(self, event):  # 选择文件夹点击事件
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            print(dlg.GetPath())  # 文件夹路径
            dir=os.listdir(dlg.GetPath())
            print(dir)
            for i in dir:
                a, b = os.path.splitext(i)
                if b=='.docx':
                    self.listBox.Append(dlg.GetPath() + '\\' + a + '.docx')
                    self.file.SetValue(dlg.GetPath())
        else :
            raise Exception("抛出一个密码异常")
            message = '文件打开失败'
            wx.MessageBox(message)  # 弹出提示框
        dlg.Destroy()

    def OnButton_selct2(self, event):  # 选择文件点击事件
        # wildcard = u"Python 文件 (*.py)|*.py|" \
        #            u"文本文件 (*.txt)|*.txt|" \
        #            u"pdf文件 (*.pdf)|*.pdf|" \
        #            "Egg file (*.egg)|*.egg|" \
        #            "All files (*.*)|*.*"
        wildcard = u"word文件 (*.docx)|*.docx|" \
                   "All files (*.*)|*.*"
        dlg = wx.FileDialog(self, message=u"选择文件",
                            defaultDir=os.getcwd(),
                            defaultFile="",
                            wildcard=wildcard,
                            style=wx.FD_OPEN)

        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()  # 返回一个list，如[u'E:\\test_python\\Demo\\ColourDialog.py', u'E:\\test_python\\Demo\\DirDialog.py']
            self.listBox.Append(paths)
            print(paths[0].split(".")[0])
            path1=paths[0].split(".")[0]
            self.file.SetValue(path1+'.pdf')
        else:
            print('打开失败')
        dlg.Destroy()
if __name__ == '__main__':
    app = wx.App()  # 初始化
    frame = windows2(parent=None, id=-1)  # 实例MyFrame类，并传递参数
    # frame.ShowFullScreen(False)
    frame.Show()  # 显示窗口
    app.MainLoop()  # 调用主循环方法