import win32com.client
from re import findall
from pandas import read_excel, read_csv
from tkinter import *
from tkinter.messagebox import showinfo, askyesno
from os.path import isfile


def copy_paste_for_ru_layout(event):
    if event.keycode == 86 and ((event.state & 0x4) != 0) and event.keysym.lower() != "v":
        event.widget.event_generate("<<Paste>>")

    elif event.keycode == 67 and ((event.state & 0x4) != 0) and event.keysym.lower() != "c":
        event.widget.event_generate("<<Copy>>")


class App:
    def __init__(self, window):
        self.window = window
        self.window.title("Link down")
        self.found_channel = None
        self.main_file = True

        self.point1_lbl = Label(window, text='Router1')
        self.point1 = Entry(window)
        self.point1_lbl.grid(row=0, column=0, padx=5, pady=5)
        self.point1.grid(row=0, column=1, padx=5, pady=5)

        self.point2_lbl = Label(window, text='Router2')
        self.point2 = Entry(window)
        self.point2_lbl.grid(row=1, column=0, padx=5, pady=5)
        self.point2.grid(row=1, column=1, padx=5, pady=5)

        self.search_btn = Button(window, text="Search Channel", width=20, command=self.search_channel)
        self.exit_btn = Button(window, text='Exit Program', width=20, command=self.window.destroy)
        self.search_btn.grid(row=3, column=1, padx=5, pady=5)
        self.exit_btn.grid(row=4, column=1, padx=5, pady=5)

        self.tt_lbl = Label(window, text='Trouble Ticket')
        self.tt = Entry(window)

        self.start_time_lbl = Label(window, text='Start Time')
        self.start_time = Entry(window)

        self.send_mail_btn = Button(window, text="Send the mail", width=20, command=self.send_mail)
        self.restart_btn = Button(window, text='Restart the program', width=20, command=self.restart_program)

        self.window.bind("<Key>", copy_paste_for_ru_layout)
        self.window.mainloop()

    def restart_program(self):
        self.window.destroy()
        main()

    def search_channel(self):
        p1 = self.point1.get()
        p2 = self.point2.get()
        if p1 == "" or p2 == "":
            showinfo(title='INFO', message="Enter 'Router1' and 'Router2'")
        else:
            df = read_excel('File_Channel_1.xlsx', sheet_name=0, dtype=str).fillna('')
            if df.empty:
                showinfo(title='ERROR', message="DataFrame is empty!\n'File_Channel_1.xlsx'")
            else:
                self.found_channel = df.loc[(df.Router1.str.contains(p1)) & (df.Router2.str.contains(p2)) |
                                            (df.Router1.str.contains(p2)) & (df.Router2.str.contains(p1))]
                if self.found_channel.empty:
                    self.search_channel_in_another_file(p1, p2)
                else:
                    if askyesno(title="Channel found!", message="Do you want to send the mail?"):
                        self.change_window()

    def search_channel_in_another_file(self, p1, p2):
        self.main_file = False
        df = read_excel('File_Channel_2.xlsx', sheet_name=0, dtype=str).fillna('')
        if df.empty:
            showinfo(title='ERROR', message="DataFrame is empty!\n'File_Channel_2.xlsx'")
        else:
            self.found_channel = df.loc[(df['Point Router 1'].str.contains(p1)) &
                                        (df['Point Router 2'].str.contains(p2)) |
                                        (df['Point Router 1'].str.contains(p2)) &
                                        (df['Point Router 2'].str.contains(p1))]
            if self.found_channel.empty:
                showinfo(title='INFO', message="Channel not found")
            else:
                important_information = self.found_channel['important information'].values[0]
                if askyesno(title='Channel found!',
                            message='Important: "{}".\nDo you want to send the mail?'.format(important_information)):
                    self.change_window()

    def change_window(self):
        self.point1_lbl.destroy()
        self.point1.destroy()
        self.point2_lbl.destroy()
        self.point2.destroy()
        self.search_btn.destroy()
        self.exit_btn.destroy()

        self.tt_lbl.grid(row=1, column=0, padx=5, pady=5)
        self.tt.grid(row=1, column=1, padx=5, pady=5)
        self.start_time_lbl.grid(row=2, column=0, padx=5, pady=5)
        self.start_time.grid(row=2, column=1, padx=5, pady=5)
        self.send_mail_btn.grid(row=3, column=1, padx=5, pady=5)
        self.restart_btn.grid(row=4, column=1, padx=5, pady=5)

    def send_mail(self):
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.Display()
        reg_dict = read_csv("dict.csv", encoding='cp1251', header=0)

        self.found_channel = self.found_channel.replace(r'\n', ' ', regex=True)

        if self.main_file:
            name_channel = self.found_channel['ID channel'].values[0]
            contact = self.found_channel['Контакты для связи'].values[0]

        else:
            name_channel = self.found_channel['Name channel'].values[0]
            contact = self.found_channel['Email for message'].values[0]

        region_name = self.found_channel['region_reduction'].values[0]
        if region_name != '':
            region_id = reg_dict.loc[reg_dict.Regions.str.contains(region_name)]['id'].values[0]
        else:
            region_id = 'Region_reduction'
            showinfo(title='INFO', message="Error in the 'region_reduction' field!"
                                           "\nEnter information manually instead of 'Region_reduction'.")

        mail_to = findall("([a-zA-Z0-9_.+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]+)", contact)

        if self.start_time.get() != '':
            if not mail_to:
                showinfo(title='INFO', message="Missing email")
            else:
                message.To = '{};'.format("; ".join(mail_to))
            message.Cc = 'Monitoring.{}@gmail.com'.format(region_id)

        message.Subject = '{} Проблема на канале {} {}'.format(region_id, name_channel, self.tt.get())

        html_body = """
        <div>
        <body>Добрый день.
        <br>Проблема на канале {0}
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


def main():
    root = Tk()
    root.geometry("250x170")
    x = (root.winfo_screenwidth() - root.winfo_reqwidth()) / 2
    y = (root.winfo_screenheight() - root.winfo_reqheight()) / 2
    root.wm_geometry("+%d+%d" % (x, y))
    App(root)


if __name__ == '__main__':
    if isfile('dict.csv') & isfile('File_Channel_1.xlsx') & isfile('File_Channel_2.xlsx'):
        main()
    else:
        showinfo(title='INFO', message='Download files to directory:\n- "dict.csv"\n- "File_Channel_1.xlsx"'
                                       '\n- "File_Channel_2.xlsx"')
