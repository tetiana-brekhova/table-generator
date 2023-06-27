import tkinter
from tkinter import ttk, messagebox, filedialog
from docx import Document


class TableCreator(tkinter.Tk):
    def __init__(self):
        super().__init__()

        self.grid_rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.title("Генератор таблиць")
        self.geometry("400x400")

        self.frame_main = tkinter.Frame(self, bg="gray")
        self.frame_main.grid(sticky='news')

        self.add_row_button = ttk.Button(self.frame_main, text="Додати рядок", command=self.add_new_row)
        self.add_row_button.grid(row=0, column=0, pady=3)

        self.add_column_button = ttk.Button(self.frame_main, text="Додати стовпець", command=self.add_new_column)
        self.add_column_button.grid(row=1, column=0, pady=3)

        self.save_button = ttk.Button(self.frame_main, text="Зберегти", command=self.save_table)
        self.save_button.grid(row=2, column=0, pady=3)

        self.frame_canvas = tkinter.Frame(self.frame_main, width=500, height=500)

        self.frame_canvas.grid(row=3, column=0, pady=(5, 0), sticky='nw')
        self.frame_canvas.grid_rowconfigure(0, weight=1)
        self.frame_canvas.grid_columnconfigure(0, weight=1)
        self.frame_canvas.grid_propagate(True)

        self.canvas = tkinter.Canvas(self.frame_canvas)
        self.canvas.grid(row=0, column=0, sticky="news")

        self.vsb = tkinter.Scrollbar(self.frame_canvas, orient=tkinter.VERTICAL, command=self.canvas.yview)
        self.vsb.grid(row=0, column=1, sticky='ns')
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.hsb = tkinter.Scrollbar(self.frame_canvas, orient=tkinter.HORIZONTAL, command=self.canvas.xview)
        self.hsb.grid(row=1, column=0, sticky='we')
        self.canvas.configure(xscrollcommand=self.hsb.set)

        self.frame_table = tkinter.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.frame_table, anchor='nw')

        self.e = tkinter.Entry(self.frame_table, width=15, fg='black', font=('Arial', 14, 'bold'))
        self.e.grid(row=0, column=0)

        self.data_lst = [[self.e]]

        self.total_rows = len(self.data_lst)
        self.total_columns = len(self.data_lst[0])

        self.frame_table.update_idletasks()
        # self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.canvas.config(scrollregion=self.canvas.bbox("rect"))
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def add_new_row(self):
        row_entries = []

        for colum in range(self.total_columns):
            entry = tkinter.Entry(self.frame_table, width=15, fg='black', font=('Arial', 14, 'bold'))
            entry.grid(row=(self.total_rows), column=colum, padx=5, pady=5)
            row_entries.append(entry)

        self.total_rows += 1
        self.data_lst.append(row_entries)

    def add_new_column(self):
        for i in range(self.total_rows):
            entry = tkinter.Entry(self.frame_table, width=15, fg='black', font=('Arial', 14, 'bold'))
            entry.grid(row=i, column=self.total_columns, padx=5, pady=5)
            self.data_lst[i].append(entry)
        self.total_columns += 1

    def save_table(self):
        document = Document()
        table = document.add_table(rows=0, cols=self.total_columns)

        for item_row in self.data_lst:
            row = table.add_row().cells
            for item in range(len(item_row)):
                row[item].text = str(item_row[item].get())

        file_path = filedialog.asksaveasfilename(defaultextension="table.csv")
        if file_path:
            document.save(file_path)
            # or:
            # document.save(table.docx)
            messagebox.showinfo("Успіх", "Таблицю успішно збережено.")
        else:
            messagebox.showwarning("Помилка", "Не вибрано файл для збереження.")


app = TableCreator()
app.mainloop()

