#coding:utf-8

import Tkinter
import os
import tkFileDialog
import xlrd
import xlwt


APP_WINDOW = '600x400+200+200'
APP_TITLE = '创源微致考勤记录转换程序'
APP_LOGO = 'cy-logo.ico'

class MyAPP(Tkinter.Frame):
    @classmethod
    def logout(cls, logstr=''):
        cls.loghandler.insert(Tkinter.END, '\n{}'.format(logstr))
        # 此处可以增加判断鼠标焦点是否在文档的最末尾，如果在最末尾或者无焦点则将执行yview_moveto
        # 否则不执行，这样做是为了便于阅读log
        cls.loghandler.yview_moveto(1)

    def _choose_input(self):
        path_ = tkFileDialog.askopenfilename(filetypes=[('xlsx','xlsx')])
        self.entry_input_val.set(path_)

        if self.entry_input.get() and self.entry_output.get():
            self.btn_convert.config(state=Tkinter.NORMAL)

        # self.btn_start.config(state=Tkinter.DISABLED)
        # self.btn_end.config(state=Tkinter.NORMAL)
        # self.entry_server.config(state=Tkinter.DISABLED)
        # self.entry_port.config(state=Tkinter.DISABLED)
        # self.entry_api.config(state=Tkinter.DISABLED)
        # self.__class__.logout('Start Web server at {}:{}, api is {}'.format(server, port, api))

    def _choose_output(self):
        path_ = tkFileDialog.asksaveasfilename(filetypes=[('xls', 'xls')], initialfile='output.xls')
        self.entry_output_val.set(path_)

        if self.entry_input.get() and self.entry_output.get():
            self.btn_convert.config(state=Tkinter.NORMAL)

        # self.btn_start.config(state=Tkinter.NORMAL)
        # self.btn_end.config(state=Tkinter.DISABLED)
        # self.entry_server.config(state=Tkinter.NORMAL)
        # self.entry_port.config(state=Tkinter.NORMAL)
        # self.entry_api.config(state=Tkinter.NORMAL)

    def _convert(self):
        curpath = os.path.dirname(__file__)
        filepath = self.entry_input.get()
        excel_handle = xlrd.open_workbook(filepath)
        sheet0 = excel_handle.sheet_by_index(0)
        rows = sheet0.nrows
        columns = sheet0.ncols
        excel_output = xlwt.Workbook(encoding='utf-8')
        excel_output_sheet = excel_output.add_sheet('MySheet1')
        out_idx = 1
        for i in range(0, rows):
            if i == 0:
                for j in range(0, columns):
                    value = sheet0.row_values(i)[j].encode('utf-8')
                    excel_output_sheet.write(i, j, label=value)
            else:
                for j in range(0, columns):
                    value = sheet0.row_values(i)[j].encode('utf-8')
                    # 标题行直接拷贝
                    if j < 4:
                        excel_output_sheet.write(out_idx, j, label=value)
                        excel_output_sheet.write(out_idx + 1, j, label=value)
                    else:
                        time_list = value.split('\n')
                        if len(time_list) > 1:
                            excel_output_sheet.write(out_idx, j, label=time_list[0])
                            excel_output_sheet.write(out_idx + 1, j, label=time_list[-1])
                        elif len(time_list) > 0:
                            excel_output_sheet.write(out_idx, j, label=time_list[0])
                        else :
                            pass
                out_idx += 2

        target = self.entry_output.get()
        try:
            if os.path.exists(target):
                os.remove(target)
            excel_output.save(target)
        except Exception as e:
            self.__class__.logout(e)
            self.__class__.logout('转换失败')
        else:
            self.__class__.logout('转换完成')



    def __init__(self):
        self.app = Tkinter.Tk()
        Tkinter.Frame.__init__(self, master=self.app)
        self.app.geometry(APP_WINDOW)
        self.app.title(APP_TITLE)
        self.app.iconbitmap(APP_LOGO)
        self.web_t = None

        # 标签
        self.title_server = Tkinter.Label(self.app, text='输入')
        self.title_server.config(font='Helvetica -15 bold', fg='blue')
        self.title_server.place(x=50, y=20, anchor="center")

        self.title_api = Tkinter.Label(self.app, text='输出')
        self.title_api.config(font='Helvetica -15 bold', fg='blue')
        self.title_api.place(x=50, y=60, anchor="center")

        # 输入框
        self.entry_input_val = Tkinter.StringVar()
        self.entry_input = Tkinter.Entry(self.app, textvariable=self.entry_input_val)
        self.entry_input.place(x=90, y=10, width=300)

        self.entry_output_val = Tkinter.StringVar()
        self.entry_output = Tkinter.Entry(self.app, textvariable=self.entry_output_val)
        self.entry_output.place(x=90, y=50, width=300)

        # 按钮
        self.btn_input = Tkinter.Button(self.app, text='选择输入', command=self._choose_input)
        self.btn_input.place(x=400, y = 10, height=20)
        self.btn_input.config(state=Tkinter.NORMAL)

        self.btn_output = Tkinter.Button(self.app, text='选择输出', command=self._choose_output)
        self.btn_output.place(x=400, y=50, height=20)
        self.btn_output.config(state=Tkinter.NORMAL)

        self.btn_convert = Tkinter.Button(self.app, text='转换', command=self._convert)
        self.btn_convert.place(x=275, y=100, width=80)
        self.btn_convert.config(state=Tkinter.DISABLED)


        # 消息
        self.message_log0 = Tkinter.Text(self.app, background='gray', borderwidth=1)
        self.message_log0.place(x=30, y=150, width=558, height=220)

        self.message_log = Tkinter.Text(self.app, background='gray', borderwidth=1)
        self.message_log.place(x=30, y= 150, width=540, height=220)
        setattr(self.__class__, 'loghandler', self.message_log)

        # 消息框滚动条
        self.scrollbar_msg = Tkinter.Scrollbar(self.message_log0)
        self.scrollbar_msg.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)
        # 绑定消息框和scrollbar
        self.message_log['yscrollcommand'] = self.scrollbar_msg.set
        self.scrollbar_msg['command'] = self.message_log.yview

if __name__ == '__main__':
    ap = MyAPP()
    ap.mainloop()