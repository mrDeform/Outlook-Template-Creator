from re import findall
from pandas import read_excel
import win32com.client
from tkinter import *
from tkinter.messagebox import showinfo, askyesno
# pandas.set_option('display.max_columns', None)
# pandas.set_option('display.width', 600)


class App:
    def __init__(self, window):
        self.window = window
        self.window.title("Link down")
        self.found_channel = None

        self.point1_lbl = Label(window, text='Router1')
        self.point1_lbl.grid(row=0, column=0, padx=5, pady=5)
        self.point1 = Entry(window)
        self.point1.grid(row=0, column=1, padx=5, pady=5)

        self.point2_lbl = Label(window, text='Router2')
        self.point2_lbl.grid(row=1, column=0, padx=5, pady=5)
        self.point2 = Entry(window)
        self.point2.grid(row=1, column=1, padx=5, pady=5)

        self.region_id_lbl = Label(window, text='Region ID')
        self.region_id = Entry(window)

        self.tt_lbl = Label(window, text='Trouble Ticket')
        self.tt = Entry(window)

        self.start_time_lbl = Label(window, text='Start Time')
        self.start_time = Entry(window)

        self.btn_send_mail = Button(window, text="Send the mail", width=20, command=self.send_mail)
        self.btn_search = Button(window, text="Search Channel", width=20, command=self.channel_search)
        self.exit_button = Button(window, text='Exit Program', width=20, command=self.window.destroy)

        self.btn_search.grid(row=3, column=1, padx=5, pady=5)
        self.exit_button.grid(row=4, column=1, padx=5, pady=5)

        self.window.mainloop()

    def send_mail(self):
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.Display()
        name_channel = self.found_channel['ID channel'].values[0]
        mail_to = findall("([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z]+)",
                          self.found_channel['Контакты для связи'].values[0])
        message.To = '{};'.format("; ".join(mail_to))
        message.Cc = 'Monitoring.{}@gmail.com'.format(self.region_id.get())
        message.Subject = '{} Авария на канале {} {}'.format(self.region_id.get(), name_channel, self.tt.get())
        html_body = """
        <div>
        <body>Добрый день.
        <br>Авария на канале {0}
        <br>Просьба взять в работу
        <br>{1}
        <br>
        <br>Начало(МСК): {2}
        <br>{3}
        <br>
        <br>Подпись
        </body>
        </div>
        """
        message.HTMLBody = html_body.format(name_channel, self.found_channel.to_html(index=False),
                                            self.start_time.get(), self.tt.get())

    def channel_search(self):
        df = read_excel('test.xlsx', sheet_name=0)
        if df.empty:
            showinfo(title='ERROR', message="DataFrame is empty!")
        else:
            self.found_channel = df.loc[(df.Router1 == self.point1.get()) & (df.Router2 == self.point2.get()) |
                                        (df.Router1 == self.point2.get()) & (df.Router2 == self.point1.get())]
            if self.found_channel.empty:
                showinfo(title='INFO', message="Channel not found")
            else:
                if askyesno(title="Channel found!", message="Do you want to send the mail?"):
                    self.point1_lbl.grid_forget()
                    self.point1.grid_forget()
                    self.point2_lbl.grid_forget()
                    self.point2.grid_forget()
                    self.btn_search.grid_forget()

                    self.region_id_lbl.grid(row=0, column=0, padx=5, pady=5)
                    self.region_id.grid(row=0, column=1, padx=5, pady=5)
                    self.tt_lbl.grid(row=1, column=0, padx=5, pady=5)
                    self.tt.grid(row=1, column=1, padx=5, pady=5)
                    self.start_time_lbl.grid(row=2, column=0, padx=5, pady=5)
                    self.start_time.grid(row=2, column=1, padx=5, pady=5)
                    self.btn_send_mail.grid(row=3, column=1, padx=5, pady=5)


def main():
    root = Tk()
    root.geometry("250x170")
    x = (root.winfo_screenwidth() - root.winfo_reqwidth()) / 2
    y = (root.winfo_screenheight() - root.winfo_reqheight()) / 2
    root.wm_geometry("+%d+%d" % (x, y))
    App(root)


if __name__ == '__main__':
    main()
