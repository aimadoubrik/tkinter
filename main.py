import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from ttkbootstrap import *
import sqlite3
import openpyxl

#________________________________________________________________________________________

class Student:
    #create main window
    def __init__(self, root):
        #create main window _____________________________________________________
        self.root = root
        self.root.title("My App")
        # Create a SQLite database and table _____________________________________
        self.conn = sqlite3.connect('student_data.db')
        self.cursor = self.conn.cursor()
        self.create_table()

        # Create Notebook___________________________________________________________
        notebook = ttk.Notebook(root)
        notebook.pack(pady=10, padx=10, fill='both', expand=True)

        # First Notebook - Student information________________________________________
        student_notebook = ttk.Frame(notebook)
        notebook.add(student_notebook, text='Student Information')

       
        # LabelFrames_______________________________________________________
        label_frame = ttk.LabelFrame(student_notebook, text='Les informations')
        label_frame.grid(row=0, column=0, pady=10, padx=10, sticky='nsew')

        # Labels ___________________________________________________________
        label_id = ttk.Label(label_frame, text='Student ID :')
        label_id.grid(row=0, column=0, padx=15, pady=5, sticky='w')
        label_first_name = ttk.Label(label_frame, text='First Name :')
        label_first_name.grid(row=1, column=0, padx=15, pady=5, sticky='w')
        label_middle_name = ttk.Label(label_frame, text='Middle name :')
        label_middle_name.grid(row=2, column=0, padx=15, pady=5, sticky='w')
        label_last_name = ttk.Label(label_frame, text='Last name :')
        label_last_name.grid(row=0, column=2, padx=15, pady=5, sticky='e')
        label_course = ttk.Label(label_frame, text='Course :')
        label_course.grid(row=1, column=2, padx=15, pady=5, sticky='e')

        # Entries ______________________________________________________________
        self.entry_id = ttk.Entry(label_frame) 
        self.entry_first_name = ttk.Entry(label_frame)
        self.entry_middle_name = ttk.Entry(label_frame)
        self.entry_last_name = ttk.Entry(label_frame)
        self.entry_course = ttk.Entry(label_frame)

        self.entry_id.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.entry_first_name.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.entry_middle_name.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        self.entry_last_name.grid(row=0, column=3, padx=5, pady=5, sticky='e')
        self.entry_course.grid(row=1, column=3, padx=5, pady=5, sticky='e')


        # Treeview ________________________________________________________________

        columns = ('Student ID', 'First Name', 'Middle Name', 'Last Name', 'Course')
        self.treeview = ttk.Treeview(student_notebook, columns=columns, show='headings')
        self.treeview.grid(row=2, column=0, sticky="news")

        # Set up column headings ____________________________________________________
        for col in columns:
            self.treeview.heading(col, text=col)

        # Buttons __________________________________________________________________
        button_frame = ttk.Frame(student_notebook)
        button_frame.grid(row=3, column=0, pady=10, padx=10, sticky='nsew')

        add_button = ttk.Button(button_frame, text='Add', command=self.add_data_to_treeview, style='Success.TButton outline')
        add_button.grid(row=0, column=0, padx=5, pady=5)

        delete_button = ttk.Button(button_frame, text='Delete', command=self.delete_data, style='Danger.TButton outline')
        delete_button.grid(row=0, column=1, padx=5, pady=5)

        update_button = ttk.Button(button_frame, text='Update', command=self.update_data, style='Toolbutton outline')
        update_button.grid(row=0, column=2, padx=5, pady=5)

        clear_button = ttk.Button(button_frame, text='Clear', command=self.confirm_clear, style='TButton outline')
        clear_button.grid(row=0, column=3, padx=5, pady=5)

        send_button = ttk.Button(button_frame, text='Save', command=self.save_to_excel, style='Link.TButton outline')
        send_button.grid(row=0, column=4, padx=5, pady=5)

        # Second Notebook - ToDo List ____________________________________________________
        to_do_list = ttk.Frame(notebook)
        notebook.add(to_do_list, text='To Do List')

        # Frame _____________________________________________________________________
        main_frame = ttk.Frame(to_do_list, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # LabelFrames _______________________________________________________________
        task_frame = ttk.LabelFrame(main_frame, text="Task")
        task_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        list_frame = ttk.LabelFrame(main_frame, text="Task List")
        list_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Widgets __________________________________________________________________
        self.task_entry = ttk.Entry(task_frame, width=80)
        self.task_entry.pack(padx=5, pady=25, ipady=12)

        add_button = ttk.Button(task_frame, text="Add", width=60, style="Success.TButton outline", command=self.add_task)
        add_button.pack(padx=5, pady=15)

        delete_button = ttk.Button(task_frame, text="Delete", width=60, style="Danger.TButton outline", command=self.delete_task)
        delete_button.pack(padx=5, pady=5)

        # Listbox _________________________________________________________________
        self.task_listbox = tk.Listbox(list_frame, width=40)
        self.task_listbox.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        # Configure weight for notebook rows and columns ___________________________
        student_notebook.grid_rowconfigure(0, weight=1)
        student_notebook.grid_rowconfigure(1, weight=1)
        student_notebook.grid_rowconfigure(2, weight=1)
        student_notebook.grid_columnconfigure(0, weight=1)

        # Configure weight for todolist _____________________________________________
        to_do_list.grid_rowconfigure(0, weight=1)
        to_do_list.grid_columnconfigure(0, weight=1)
        to_do_list.grid_rowconfigure(2, weight=1)
        to_do_list.grid_columnconfigure(0, weight=1)

# F O N C T I O N S __________________________________________________________________
    def create_table(self):
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY,
                student_id TEXT,
                first_name TEXT,
                middle_name TEXT,
                last_name TEXT,
                course TEXT
            )
        ''')
        self.conn.commit()
    #_______________________________________________________________________________________
    def add_data_to_treeview(self):

        # Get entries _________________________________________________
        student_id = self.entry_id.get()
        first_name = self.entry_first_name.get()
        middle_name = self.entry_middle_name.get()
        last_name = self.entry_last_name.get()
        course = self.entry_course.get()

        # Insert data database ___________________________________________
        self.cursor.execute('''
            INSERT INTO students (student_id, first_name, middle_name, last_name, course)
            VALUES (?, ?, ?, ?, ?)
        ''', (student_id, first_name, middle_name, last_name, course))
        self.conn.commit()

        # Insert data into Treeview ____________________________________________________
        self.treeview.insert('', 'end', values=(student_id, first_name, middle_name, last_name, course))

        # Clear entry fields after adding data _________________________________________
        self.entry_id.delete(0, tk.END)
        self.entry_first_name.delete(0, tk.END)
        self.entry_middle_name.delete(0, tk.END)
        self.entry_last_name.delete(0, tk.END)
        self.entry_course.delete(0, tk.END)
    
    #_______________________________________________________________________________________
    def delete_data(self):
        
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a row to delete.")
            return

        
        confirmation = messagebox.askyesno("Confirme", "Are you sure you want to delete this data?")
        if confirmation:
            
            self.treeview.delete(selected_item)
            selected_id = self.treeview.item(selected_item)['values'][0]
            self.cursor.execute('DELETE FROM students WHERE student_id = ?', (selected_id,))
            self.conn.commit()
    #_______________________________________________________________________________________
    def update_data(self):
        # Get entries _________________________________________________________________
        student_id = self.entry_id.get()
        first_name = self.entry_first_name.get()
        middle_name = self.entry_middle_name.get()
        last_name = self.entry_last_name.get()
        course = self.entry_course.get()

        selected_item = self.treeview.focus()
        if selected_item:
            self.treeview.item(selected_item, values=(student_id, first_name, middle_name, last_name, course))

            messagebox.showinfo("Update", "Data updated successfully!!!!!!")
        else:
            messagebox.showerror("Error", "Please select a one to update.")
    #_______________________________________________________________________________________
    def confirm_clear(self):
        
        confirm = messagebox.askyesno("Confirme", "Are you sure you want to clear?")
        if confirm:
            
            self.entry_id.delete(0, tk.END)
            self.entry_first_name.delete(0, tk.END)
            self.entry_middle_name.delete(0, tk.END)
            self.entry_last_name.delete(0, tk.END)
            self.entry_course.delete(0, tk.END)
    #_______________________________________________________________________________________
    def save_to_excel(self):
        file_name = "student_data.xlsx"
        data = []
        
        for child in self.treeview.get_children():
            item = self.treeview.item(child)['values']
            data.append(item)
        
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        headers = ['Student ID', 'First Name', 'Middle Name', 'Last Name', 'Course']
        sheet.append(headers) 

        for item in data:
            sheet.append(item)

        workbook.save(file_name)
    #_______________________________________________________________________________________
    def add_task(self):

        task = self.task_entry.get()
        self.task_listbox.insert(tk.END, task)

    #_______________________________________________________________________________________
    def delete_task(self):

        selected_index = self.task_listbox.curselection()
        if selected_index:
            self.task_listbox.delete(selected_index)
            
        else:
            pass

# display main window ______________________________________________________________________

if __name__ == "__main__":
    root = tk.Tk()
    style = Style(theme='flatly')
    app = Student(root)
    root.mainloop()
