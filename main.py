import tkinter
from tkinter import *
from tkinter import messagebox
import xlrd


def main_function():
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
    icon = PhotoImage(file="ring_icon.png")
    win.iconphoto(False, icon)

    # set a background image on all of the window
    # Add image file
    bg = PhotoImage(file="background_win.png")

    # Create Canvas
    canvas1 = Canvas(win, width=app_width, height=app_height)

    canvas1.pack(fill="both", expand=True)

    # Display image
    canvas1.create_image(0, 0, image=bg, anchor="nw")

    # calls the function that handle all the widgets in the first window
    main_window(win)

    win.mainloop()


def main_window(win):
    main_label = Label(win, text="ברוכים הבאים לחתונה של קארין ואופק", font="SuezOne 25 bold")
    main_label.place(relx=0.1, rely=0.3)
    main_label.config(bg="#FFFFFF")

    data_label = Label(win, text="6.7.22", font="SuezOne 15 bold")
    data_label.place(relx=0.0, rely=1.0, anchor='sw')

    place_name_label = Label(win, text="אולמי קאסטלו", font="SuezOne 15 bold")
    place_name_label.place(relx=1, rely=1, anchor='se')

    message_telephone = Label(win, text="הכנס את מספר הטלפון שלך ולחץ על הכפתור למטה ",
                              font="SuezOne 13 bold", bg="#FFFFFF")
    message_telephone.place(relx=0.5, rely=0.5, anchor='center')

    input_telephone = Entry(win, bg="white", width=20, borderwidth=5, font="bold")
    input_telephone.place(relx=0.5, rely=0.6, anchor='center')
    input_telephone.focus()

    # a function that extract the data from excel file
    data_dict = extract_data()

    button_start = Button(win, text="מצא את השולחן שלי", font="SuezOne 20 bold",
                          command=lambda: find_table(input_telephone.get(), input_telephone, data_dict))
    button_start.place(relx=0.5, rely=0.8, anchor='center')


def find_table(input_tele_str, input_telephone, data_dict):
    input_telephone.delete(0, END)
    telephone_guest = input_tele_str
    # check input for numbers only and in the right length
    if not telephone_guest.isdigit():
        messagebox.showwarning("weeding", "הטלפון לא מורכב רק ממספרים, נסה שנית")
    elif not len(telephone_guest) == 10:
        messagebox.showwarning("weeding", "מספר הטלפון קצר מדי, נסה שוב")

    full_name = data_dict[telephone_guest][0]
    partner_name = data_dict[telephone_guest][1]
    number_of_invites = data_dict[telephone_guest][2]
    number_approve = data_dict[telephone_guest][3]
    table_number = data_dict[telephone_guest][4]

    if number_of_invites == 1:
        # TODO: https://www.youtube.com/watch?v=tpwu5Zb64lQ CUSTOM MessageBox
        messagebox.showinfo("Wedding", f"שלום {full_name} מספר השולחן שלך הוא: {table_number}.\n")


def extract_data() -> dict:
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
    path_excel = "weeding.xlsx"
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


if __name__ == '__main__':
    main_function()
