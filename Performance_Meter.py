######################## GUI And SQL Modules  #######################################
import _tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import pymysql

##########################   Modules For Scraping The Data  ##########################
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.support.ui import Select
import time


class Student:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Management System")
        self.root.geometry("1350x700+0+0")

        ###################### All Variables  #########################

        self.Roll_No_var = StringVar()
        self.name_var = StringVar()
        self.email_var = StringVar()
        self.gender_var = StringVar()
        self.contact_var = StringVar()
        self.section = StringVar()
        self.sem_var = IntVar()
        self.fetch_sem_var = IntVar()
        self.search_by = StringVar()
        self.sort_by = StringVar()
        self.search_txt = StringVar()
        self.sgpa_var = DoubleVar()

        self.row1 = []
        self.avg = 0
        self.sgpa=0

        ##################### Title Frame #############################

        Title_Frame = Frame(self.root, bd=8, relief=GROOVE, bg="#7303fc")
        Title_Frame.pack(side=TOP, fill=X)

        lbl_Section = Label(Title_Frame, text="Select Section ", font=("times new roman", 15, "bold"), bg="#7303fc",
                            fg="white")
        lbl_Section.grid(row=0, column=0, pady=10, padx=10, sticky="w")

        combo_section = ttk.Combobox(Title_Frame, textvariable=self.section, font=("times new roman", 12, "bold"),
                                     width="5", state="readonly")
        combo_section['values'] = ("cse1", "cse2")
        combo_section.current(0)
        combo_section.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        Show_all_btn = Button(Title_Frame, text="Show All", width=8, pady=2, command=self.fetch_data).grid(row=0,
                                                                                                          column=2,
                                                                                                          padx=10,
                                                                                                          pady=10)

        title = Label(Title_Frame, text="CSE_PER_METER", bd=1, relief=GROOVE, font=("times new roman", 25, "bold"),
                      bg="#7303fc", fg="#ffffff")
        title.grid(row=0, column=3, padx=200, pady=10)

        lbl_fetch_sem = Label(Title_Frame, text="Select Sem ", font=("times new roman", 15, "bold"), bg="#7303fc",
                              fg="white")
        lbl_fetch_sem.grid(row=0, column=4, pady=10, padx=10, sticky="w")

        self.combo_fetch_sem = ttk.Combobox(Title_Frame, textvariable=self.fetch_sem_var,
                                            font=("times new roman", 12, "bold"),
                                            width="5", state="readonly")
        self.combo_fetch_sem['values'] = (1, 2, 3, 4, 5, 6, 7, 8)
        self.combo_fetch_sem.current("0")
        self.combo_fetch_sem.grid(row=0, column=5, padx=10, pady=10, sticky="w")

        Fetch_btn = Button(Title_Frame, text="Fetch", width=8, pady=2, command=self.fetch_sem_data).grid(row=0,
                                                                                                         column=6,
                                                                                                         padx=10,
                                                                                                         pady=10)

        #####################    Manage Frame     ######################
        Manage_Frame = Frame(self.root, bd=4, relief=RIDGE, bg="#7303fc")
        Manage_Frame.place(x=10, y=85, width=450, height=600)

        m_title = Label(Manage_Frame, text="Manage Students", font=("times new roman", 25, "bold"), bg="#7303fc",
                        fg="white")
        m_title.grid(row=0, columnspan=2, pady=20, padx=90)
        # Roll Number Field
        lbl_roll = Label(Manage_Frame, text="Roll No*", font=("times new roman", 15, "bold"), bg="#7303fc", fg="white")
        lbl_roll.grid(row=1, column=0, pady=10, padx=30, sticky="w")

        txt_roll = Entry(Manage_Frame, textvariable=self.Roll_No_var, font=("times new roman", 10, "bold"), bd=3,
                         relief=GROOVE)
        txt_roll.grid(row=1, column=1, pady=10, padx=10, sticky="w")
        # Name Field
        lbl_name = Label(Manage_Frame, text="Name*", font=("times new roman", 15, "bold"), bg="#7303fc", fg="white")
        lbl_name.grid(row=2, column=0, pady=10, padx=30, sticky="w")

        txt_name = Entry(Manage_Frame, textvariable=self.name_var, font=("times new roman", 10, "bold"), bd=3,
                         relief=GROOVE)
        txt_name.grid(row=2, column=1, pady=10, padx=10, sticky="w")
        # E-mail Field
        lbl_Email = Label(Manage_Frame, text="Email*", font=("times new roman", 15, "bold"), bg="#7303fc", fg="white")
        lbl_Email.grid(row=3, column=0, pady=10, padx=30, sticky="w")

        txt_Email = Entry(Manage_Frame, textvariable=self.email_var, font=("times new roman", 10, "bold"), bd=3,
                          relief=GROOVE)
        txt_Email.grid(row=3, column=1, pady=10, padx=10, sticky="w")

        # Gender   Field
        lbl_Gender = Label(Manage_Frame, text="Gender*", font=("times new roman", 15, "bold"), bg="#7303fc", fg="white")
        lbl_Gender.grid(row=4, column=0, pady=10, padx=30, sticky="w")

        combo_gender = ttk.Combobox(Manage_Frame, textvariable=self.gender_var, font=("times new roman", 12, "bold"),
                                    width="16", state="readonly")
        combo_gender['values'] = ("male", "female", "other")
        combo_gender.grid(row=4, column=1, padx=10, pady=10, sticky="w")

        # Contact Field
        lbl_Contact = Label(Manage_Frame, text="Contact", font=("times new roman", 15, "bold"), bg="#7303fc",
                            fg="white")
        lbl_Contact.grid(row=5, column=0, pady=10, padx=30, sticky="w")

        txt_Contact = Entry(Manage_Frame, textvariable=self.contact_var, font=("times new roman", 10, "bold"), bd=3,
                            relief=GROOVE)
        txt_Contact.grid(row=5, column=1, pady=10, padx=10, sticky="w")

        # Select Semester Field
        lbl_Sem = Label(Manage_Frame, text="Semester", font=("times new roman", 15, "bold"), bg="#7303fc", fg="white")
        lbl_Sem.grid(row=6, column=0, pady=10, padx=30, sticky="w")

        self.combo_sem = ttk.Combobox(Manage_Frame, textvariable=self.sem_var, font=("times new roman", 12, "bold"),
                                      width="16", state="readonly")
        self.combo_sem.bind("<<ComboboxSelected>>", self.fetch_selected_sem_sgpa)

        self.combo_sem['values'] = (1, 2, 3, 4, 5, 6, 7, 8)
        self.combo_sem.current("0")
        self.combo_sem.grid(row=6, column=1, padx=10, pady=10, sticky="w")

        # Enter SGPA Field
        lbl_Sgpa = Label(Manage_Frame, text="SGPA", font=("times new roman", 15, "bold"), bg="#7303fc", fg="white")
        lbl_Sgpa.grid(row=7, column=0, pady=10, padx=30, sticky="w")

        txt_Sgpa = Entry(Manage_Frame, text="", textvariable=self.sgpa_var, font=("times new roman", 10, "bold"), bd=3,
                         relief=GROOVE)
        self.sgpa_var.set("")
        txt_Sgpa.grid(row=7, column=1, pady=10, padx=10, sticky="w")

        ################    Button Frame     #################

        Button_Frame = Frame(Manage_Frame, bd=4, relief=RIDGE, bg="#7303fc")
        Button_Frame.place(x=10, y=516, width=430)

        Add_btn = ttk.Button(Button_Frame, text="Add", width=10, command=self.add_students).grid(row=0, column=0, padx=10,
                                                                                             pady=10)
        Update_btn = ttk.Button(Button_Frame, text="Update", width=10, command=self.update_data).grid(row=0, column=1,
                                                                                                  padx=10, pady=10)
        Delete_btn = ttk.Button(Button_Frame, text="Delete", width=10, command=self.delete_data).grid(row=0, column=2,
                                                                                                  padx=10, pady=10)
        Clear_btn = ttk.Button(Button_Frame, text="Clear", width=10, command=self.clear).grid(row=0, column=3, padx=10,
                                                                                          pady=10)

        #####################    Details Frame     ######################

        Details_Frame = Frame(self.root, bd=4, relief=RIDGE, bg="#7303fc")
        Details_Frame.place(x=470, y=85, width=870, height=600)

        lbl_Search = Label(Details_Frame, text="Search By", font=("times new roman", 15, "bold"), bg="#7303fc",
                           fg="white")
        lbl_Search.grid(row=0, column=0, pady=10, padx=10, sticky="w")

        combo_search = ttk.Combobox(Details_Frame, textvariable=self.search_by, font=("times new roman", 12, "bold"),
                                    width="10", state="readonly")
        combo_search['values'] = ("roll_no", "name", "contact")
        combo_search.current(0)
        combo_search.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        txt_Search = Entry(Details_Frame, textvariable=self.search_txt, width=15, font=("times new roman", 10, "bold"),
                           bd=3, relief=GROOVE)
        txt_Search.grid(row=0, column=2, pady=10, padx=10, sticky="w")

        Search_btn = Button(Details_Frame, text="Search", width=8, pady=2, command=self.search_data).grid(row=0,
                                                                                                          column=3,
                                                                                                          padx=10,
                                                                                                          pady=10)
        lbl_Sort_by = Label(Details_Frame, text="Sort By Sem", font=("times new roman", 15, "bold"), bg="#7303fc",
                            fg="white")
        lbl_Sort_by.grid(row=0, column=5, pady=10, padx=10, sticky="w")

        combo_sort_by = ttk.Combobox(Details_Frame, textvariable=self.sort_by, font=("times new roman", 12, "bold"),
                                     width="7", state="readonly")
        combo_sort_by['values'] = ("sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8", "avg")
        combo_sort_by.current(0)
        combo_sort_by.grid(row=0, column=6, padx=10, pady=10, sticky="w")

        Sort_btn = Button(Details_Frame, text="Sort", width=8, pady=2, command=self.sort_data).grid(row=0, column=7,
                                                                                                    padx=10, pady=10)
        Sort_overall_btn = Button(Details_Frame, text="Overall", width=8, pady=2, command=self.sort_overall).grid(row=0, column=8,
                                                                                                    padx=10, pady=10)

        #########################   Table Frame    ###############################

        Table_Frame = Frame(Details_Frame, bd=4, relief=RIDGE, bg="#7303fc")
        Table_Frame.place(x=10, y=70, width=840, height=500)
        scroll_x = Scrollbar(Table_Frame, orient=HORIZONTAL)
        scroll_y = Scrollbar(Table_Frame, orient=VERTICAL)
        self.Student_table = ttk.Treeview(Table_Frame, columns=(
        "roll", "name", "email", "gender", "contact", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8",
        "avg"), xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_x.config(command=self.Student_table.xview)
        scroll_y.config(command=self.Student_table.yview)
        self.Student_table.heading("roll", text="Roll No:")
        self.Student_table.heading("name", text="Name")
        self.Student_table.heading("email", text="Email")
        self.Student_table.heading("gender", text="Gender")
        self.Student_table.heading("contact", text="Contact")
        self.Student_table.heading("sem1", text="Sem 1")
        self.Student_table.heading("sem2", text="Sem 2")
        self.Student_table.heading("sem3", text="Sem 3")
        self.Student_table.heading("sem4", text="Sem 4")
        self.Student_table.heading("sem5", text="Sem 5")
        self.Student_table.heading("sem6", text="Sem 6")
        self.Student_table.heading("sem7", text="Sem 7")
        self.Student_table.heading("sem8", text="Sem 8")
        self.Student_table.heading("avg", text="Average")
        self.Student_table['show'] = "headings"
        self.Student_table.column("roll", width=125)
        self.Student_table.column("name", width=130)
        self.Student_table.column("email", width=220)
        self.Student_table.column("gender", width=100)
        self.Student_table.column("contact", width=100)
        self.Student_table.column("sem1", width=100)
        self.Student_table.column("sem2", width=100)
        self.Student_table.column("sem3", width=100)
        self.Student_table.column("sem4", width=100)
        self.Student_table.column("sem5", width=100)
        self.Student_table.column("sem6", width=100)
        self.Student_table.column("sem7", width=100)
        self.Student_table.column("sem8", width=100)
        self.Student_table.column("avg", width=100)
        self.Student_table.pack(fill=BOTH, expand=1)
        self.Student_table.bind("<ButtonRelease-1>", self.get_cursor)
        self.fetch_data()

#######################   All Functions  ##########################
    def add_students(self):
        if self.Roll_No_var.get() == "" or self.name_var.get() == "" or self.email_var.get() == "" or self.gender_var.get() == "":
            messagebox.showerror("Error", "Fields marked with a * are mandatory!!")
        else:
            con = pymysql.connect(host="localhost", user="root", password="", database="stm")
            cur = con.cursor()
            cur.execute("select contact from " + str(self.section.get()))
            co = cur.fetchall()
            check_contact_list = []
            co1 = list(co)
            if len(co1) != 0:
                co1.pop(0)
                for i in co1:
                    check_contact_list.append(i[0])
            if self.contact_var.get() != "" and self.contact_var.get() in check_contact_list:
                messagebox.showerror("Error", "A Person With This Phone Number Already Exists!!")
            else:
                if len(self.Roll_No_var.get()) != 4:
                    messagebox.showwarning("Warning", "Please enter a 4 digit roll number!")
                else:
                    try:
                        cur.execute(
                            "insert into " + str(self.section.get()) + " (roll_no,name,email,gender,contact,sem" + str(
                                self.sem_var) + ") value"
                                                "s(%s,%s"
                                                ",%s,%s,"
                                                "%s,%s)",
                            ("1200011" + self.Roll_No_var.get(),
                             self.name_var.get(),
                             self.email_var.get() + "@gmail.com",
                             self.gender_var.get(),
                             self.contact_var.get().strip(),
                             self.sgpa_var.get(),
                             ))
                    except pymysql.err.IntegrityError:
                        messagebox.showerror("Error", "This Person Already Exists!")
                    except _tkinter.TclError:
                        messagebox.showerror("Error", "SGPA takes only numeric values!")
                    else:
                        cur.execute(
                            "select sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8 from " + str(
                                self.section.get()) + " where roll_no=1200011" + str(
                                self.Roll_No_var.get()))
                        sem_rows = cur.fetchall()
                        sem_rows_set = set(sem_rows[0])
                        sem_rows_set.remove(0.0)
                        if len(sem_rows_set) == 0:
                            self.avg = 0
                        else:
                            self.avg = sum(sem_rows_set) / len(sem_rows_set)
                        cur.execute(
                            "update " + str(self.section.get()) + " set avg=" + str(
                                self.avg.__round__(2)) + " where roll_no=1200011" + str(
                                self.Roll_No_var.get()))
                        self.fetch_data()
                        messagebox.showinfo("Success", "Record has been inserted")
                    finally:
                        con.commit()
                        con.close()
                        self.clear()

    def fetch_data(self):
        con = pymysql.connect(host="localhost", user="root", password="", database="stm")
        cur = con.cursor()
        cur.execute("select * from " + str(self.section.get()))
        rows = cur.fetchall()
        if len(rows) != 0:
            self.Student_table.delete(*self.Student_table.get_children())
            for row in rows:
                self.Student_table.insert('', END, values=row)
            con.commit()
        con.close()

    def clear(self):
        self.Roll_No_var.set("")
        self.name_var.set("")
        self.email_var.set("")
        self.gender_var.set("")
        self.contact_var.set("")
        self.sem_var = 1
        self.combo_sem.set("1")
        self.sgpa_var.set("")
        self.row1 = []

    def fetch_selected_sem_sgpa(self, e):
        self.sem_var = int(self.combo_sem.get())
        if len(self.row1) != 0:
            if self.sem_var == 1:
                self.sgpa_var.set(self.row1[5])
            elif self.sem_var == 2:
                self.sgpa_var.set(self.row1[6])
            elif self.sem_var == 3:
                self.sgpa_var.set(self.row1[7])
            elif self.sem_var == 4:
                self.sgpa_var.set(self.row1[8])
            elif self.sem_var == 5:
                self.sgpa_var.set(self.row1[9])
            elif self.sem_var == 6:
                self.sgpa_var.set(self.row1[10])
            elif self.sem_var == 7:
                self.sgpa_var.set(self.row1[11])
            elif self.sem_var == 8:
                self.sgpa_var.set(self.row1[12])
            else:
                self.sgpa_var.set("")

    def get_cursor(self, ev):
        cursor_row = self.Student_table.focus()
        contents = self.Student_table.item(cursor_row)
        row = contents['values']
        self.row1 = row
        ch = str(row[0]).split(" ")
        if len(ch) == 3:
            self.Roll_No_var.set(ch[0])
        else:
            self.Roll_No_var.set(row[0])
        self.name_var.set(row[1])
        self.email_var.set(row[2])
        self.gender_var.set(row[3])
        self.contact_var.set(row[4])
        self.sem_var = 1
        self.combo_sem.set("1")
        self.sgpa_var.set(row[5])

    ######################  Function For Fetching The Results From The Website #######################
    def fetch_sem_data(self):
        con = pymysql.connect(host="localhost", user="root", password="", database="stm")
        cur = con.cursor()
        cur.execute("select roll_no from " + str(self.section.get()))
        co = cur.fetchall()
        list_roll = []
        counter = 0
        for item in co:
            list_roll.append(item[0])
        index = self.fetch_sem_var.get()
        driver = webdriver.Chrome(executable_path="C:\\Drivers\\chromedriver.exe")  # Initializing the web driver
        driver.get("https://makaut1.ucanapply.com/smartexam/public/result-details")
        try:
            for i in list_roll:
                driver.find_element_by_id("username").send_keys(i)
                drp = Select(driver.find_element_by_id("semester"))
                drp.select_by_index(index)
                driver.find_element_by_tag_name('button').click()
                try:
                    if index % 2 == 0:
                        line = driver.find_element_by_xpath("//strong[contains(text(),'EVEN')]").text
                    else:
                        line = driver.find_element_by_xpath("//strong[contains(text(),'ODD')]").text
                except NoSuchElementException:
                    self.sgpa = 0.0
                    counter = counter + 1
                else:
                    li = line.split(":")
                    if li[1].strip() == "--":
                        self.sgpa = 0.0
                        counter = counter + 1
                    else:
                        self.sgpa = float(li[1])
                finally:
                    cur.execute(
                        "update " + str(self.section.get()) + " set sem" + str(
                            index) + "=%s where roll_no=%s", (
                             self.sgpa,
                             str(i),
                         ))
                    cur.execute(
                        "select sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8 from " + str(
                            self.section.get()) + " where roll_no=" + str(i))
                    sem_rows = cur.fetchall()
                    sem_rows_set = set(sem_rows[0])
                    sem_rows_set.remove(0.0)
                    if len(sem_rows_set) == 0:
                        self.avg = 0
                    else:
                        self.avg = sum(sem_rows_set) / len(sem_rows_set)
                    cur.execute("update " + str(self.section.get()) + " set avg=" + str(
                            self.avg.__round__(2)) + " where roll_no=" + str(i))

                    driver.find_element_by_link_text("Reset").click()
                    time.sleep(1)
        except NoSuchElementException:
            messagebox.showerror("Error", "Please check your internet connectivity!")
        except ElementNotInteractableException:
            messagebox.showerror("Error", "You have poor net connectivity!. Please try again later")
        else:
            driver.close()
            messagebox.showinfo("Success", "Data imported successfully!")
            if counter == 1:
                messagebox.showinfo("Info", "Result of {} student is still pending! ".format(counter))
            else:
                messagebox.showinfo("Info", "Result of {} students are still pending! ".format(counter))
        finally:
            con.commit()
            con.close()
            self.fetch_data()

    def update_data(self):
        con = pymysql.connect(host="localhost", user="root", password="", database="stm")
        cur = con.cursor()
        if len(self.Roll_No_var.get()) != 11:
            messagebox.showwarning("Warning", "Please check the roll_number!")
        else:
            try:
                cur.execute(
                    "update " + str(self.section.get()) + " set name=%s,email=%s,gender=%s,contact=%s,sem" + str(
                        self.sem_var) + "=%s where roll_no=%s", (
                        self.name_var.get(),
                        self.email_var.get(),
                        self.gender_var.get(),
                        self.contact_var.get(),
                        self.sgpa_var.get(),
                        self.Roll_No_var.get(),
                    ))
            except _tkinter.TclError:
                messagebox.showerror("Error", "SGPA takes only numeric values!")
            else:
                messagebox.showinfo("Success", "All fields updated successfully.")
            finally:
                cur.execute(
                    "select sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8 from " + str(
                        self.section.get()) + " where roll_no=" + str(
                        self.Roll_No_var.get()))
                sem_rows = cur.fetchall()
                sem_rows_set = set(sem_rows[0])
                sem_rows_set.remove(0.0)
                if len(sem_rows_set) == 0:
                    self.avg = 0
                else:
                    self.avg = sum(sem_rows_set) / len(sem_rows_set)
                cur.execute("update " + str(self.section.get()) + " set avg=" + str(
                        self.avg.__round__(2)) + " where roll_no=" + str(self.
                                                                         Roll_No_var.
                                                                         get()))
            con.commit()
            con.close()
            self.fetch_data()
            self.clear()

    def delete_data(self):
        con = pymysql.connect(host="localhost", user="root", password="", database="stm")
        cur = con.cursor()
        delete_confirmation = messagebox.askquestion('Delete Record', 'Are you sure you want to delete this record?',
                                                     icon='warning')
        if delete_confirmation == 'yes':
            cur.execute("delete from " + str(self.section.get()) + " where roll_no=%s", self.Roll_No_var.get())
        con.commit()
        con.close()
        self.fetch_data()
        self.clear()

    def search_data(self):
        con = pymysql.connect(host="localhost", user="root", password="", database="stm")
        cur = con.cursor()
        cur.execute(
            "select * from " + str(self.section.get()) + " where " + str(self.search_by.get()) + " LIKE '%" + str(
                self.search_txt.get()) + "%'")
        rows = cur.fetchall()
        if len(rows) != 0:
            self.Student_table.delete(*self.Student_table.get_children())
            for row in rows:
                self.Student_table.insert('', END, values=row)
        else:
            messagebox.showinfo("Search Complete", "No Records Found For " + self.search_txt.get().capitalize() + " !")
        self.search_by.set("roll_no")
        self.search_txt.set("")
        con.commit()
        con.close()

    def sort_data(self):
        con = pymysql.connect(host="localhost", user="root", password="", database="stm")
        cur = con.cursor()
        cur.execute("select * from " + str(self.section.get()) + " order by " + str(self.sort_by.get()) + " desc")
        rows = cur.fetchall()
        if len(rows) != 0:
            self.Student_table.delete(*self.Student_table.get_children())
            i = 1
            for row in rows:
                row_list = list(row)
                row_list[0] = row_list[0] + " (rank " + str(i) + ")"
                self.Student_table.insert('', END, values=row_list)
                i = i + 1
        con.commit()
        con.close()

    def sort_overall(self):
        con = pymysql.connect(host="localhost", user="root", password="", database="stm")
        cur = con.cursor()
        cur.execute("select * from cse1 union select * from cse2 order by " + str(self.sort_by.get()) + " desc")
        rows = cur.fetchall()
        if len(rows) != 0:
            self.Student_table.delete(*self.Student_table.get_children())
            i = 1
            for row in rows:
                row_list = list(row)
                row_list[0] = row_list[0] + " (rank " + str(i) + ")"
                self.Student_table.insert('', END, values=row_list)
                i = i + 1
        con.commit()
        con.close()


root = Tk()
ob = Student(root)
root.configure(background='#fff')
root.iconbitmap(r'C:\Users\Aditya\PycharmProjects\Student_Management\icon.ico')
root.mainloop()
