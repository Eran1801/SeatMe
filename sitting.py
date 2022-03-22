import tkinter
from tkinter import *
from tkinter import messagebox
import xlrd


class Sitting:

    def __init__(self):

        global win, icon

        win = tkinter.Tk()
        win.title("Wedding")

        # Set the screen in the middle
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()

        app_width = 600
        app_height = 500

        x = (screen_width / 2) - (app_width / 2)
        y = (screen_height / 2) - (app_height / 2)

        win.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')

        # set the icon of the app
        icon = PhotoImage(file="files/icon/ring_icon.png")
        win.iconphoto(False, icon)

        # set a background image on all of the window
        # Add image file
        bg = PhotoImage(file="files/background/background_win.png")

        # Create Canvas
        canvas1 = Canvas(win, width=app_width, height=app_height)

        canvas1.pack(fill="both", expand=True)

        # Display image
        canvas1.create_image(0, 0, image=bg, anchor="nw")

        # function that handle all the widgets in the first window
        self.create_main_window(win)

        win.mainloop()

    def create_main_window(self, win):
        main_label = Label(win, text="ברוכים הבאים לחתונה של קארין ואופק", font="SuezOne 25 bold")
        main_label.place(relx=0.1, rely=0.3)
        main_label.config(bg="#FFFFFF")

        date_label = Label(win, text="6.7.22", font="SuezOne 15 bold")
        date_label.place(relx=0.0, rely=1.0, anchor='sw')
        date_label.config(bg="#FFFFFF")

        place_name_label = Label(win, text="אולמי קאסטלו", font="SuezOne 15 bold")
        place_name_label.place(relx=1, rely=1, anchor='se')
        place_name_label.config(bg="#FFFFFF")

        tele_label = Label(win, text="הכנס את מספר הטלפון שלך ולחץ על הכפתור למטה ",
                           font="SuezOne 13 bold", bg="#FFFFFF")
        tele_label.place(relx=0.5, rely=0.43, anchor='center')

        input_telephone = Entry(win, bg="white", width=20, borderwidth=5, font="bold")
        input_telephone.place(relx=0.5, rely=0.5, anchor='center')
        input_telephone.focus()

        message_to_guest = Label(win, text="", font="SuezOne 15 bold", bg="gray51", fg="red")
        message_to_guest.place(relx=0.5, rely=0.8, anchor='center')
        message_to_guest.config(bg="#FFFFFF")

        # a function that extract the data from excel file
        data_dict = self.extract_data()

        button_start = Button(win, text="מצא את השולחן שלי", font="SuezOne 20 bold",
                              command=lambda: self.find_table_new(input_telephone.get(), input_telephone, data_dict,
                                                                  message_to_guest))

        # If you press enter or press the button the function will work
        win.bind('<Return>',
                 lambda event: self.find_table_new(input_telephone.get(), input_telephone, data_dict,
                                                                  message_to_guest))

        button_start.place(relx=0.5, rely=0.62, anchor='center')

    def find_table_new(self, input_tele_str, input_telephone, data_dict, message_to_guest):
        global full_name, number_approve, table_number

        input_telephone.delete(0, END)
        telephone_guest = input_tele_str
        # check input for numbers only and in the correct length - 10
        if not telephone_guest.isdigit():
            messagebox.showwarning("weeding", "הטלפון לא מורכב רק ממספרים, נסה שוב")
        elif not len(telephone_guest) == 10:
            messagebox.showwarning("weeding", "מספר הטלפון קצר מדי, נסה שוב")

        full_name = data_dict[telephone_guest][0]
        number_of_invites = data_dict[telephone_guest][2]
        number_approve = data_dict[telephone_guest][3]
        table_number = data_dict[telephone_guest][4]

        if number_of_invites == 1:
            message_1 = f" .שלום {full_name}, מספר השולחן שלך הוא {table_number}\n" \
                        f".תעשה חיים ושמור על עצמך, קארין ואופק"
            message_to_guest.config(text=message_1)
            # After 7.5 sec the text will disappear
            win.after(10000, lambda: message_to_guest.config(text=""))
        elif number_of_invites > 1:
            message_1 = f" .שלום {full_name}, את/ה ו+{number_approve - 1} המוזמנים שאיתך יושבים בשולחן {table_number}\n" \
                        f".תעשו חיים ושמרו על עצמכם, קארין ואופק"
            message_to_guest.config(text=message_1)
            win.after(10000, lambda: message_to_guest.config(text=""))

    def extract_data(self) -> dict:
        '''
            first we need to crate a dict when the key is the phone number and the value is a list that:
            list[0] = The full name of the person
            list[1] = The name of the partner ( if he exist )
            list[2] = The number of invites
            list[3] = The number of approval
            list[4] = The number of the table
        '''

        dict_guest = {}

        # dealing with the excel file
        path_excel = "files/excel/weeding.xlsx"
        excel_workbook = xlrd.open_workbook(path_excel)
        excel_worksheet = excel_workbook.sheet_by_index(0)

        rows = excel_worksheet.nrows

        t_rows = rows - 1
        while t_rows != 0:
            phone_number = excel_worksheet.cell_value(t_rows, 6)
            round_phone = round(phone_number)
            final_phone = "0" + str(round_phone)

            full_name = excel_worksheet.cell_value(t_rows, 0) + " " + excel_worksheet.cell_value(t_rows, 1)
            partner = excel_worksheet.cell_value(t_rows, 2)
            invites = round(excel_worksheet.cell_value(t_rows, 3))
            approval = round(excel_worksheet.cell_value(t_rows, 4))
            table = round(excel_worksheet.cell_value(t_rows, 5))

            list_guest = [full_name, partner, invites, approval, table]
            dict_guest[final_phone] = list_guest

            t_rows -= 1

        return dict_guest


guest = Sitting()
