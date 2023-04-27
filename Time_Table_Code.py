import customtkinter as ctk
import tkinter as tk
import math
import tabula
import math
from tkinter import ttk
import csv
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from tkinter import messagebox
import datetime
from genetictabler import GenerateTimeTable
csv_file = '3d_array.csv'

classes = [" SE1 ", " SE2 ", " SE3 ", " SE4 ", " TE1 ", " TE2 ", " TE3 ", " TE4 ", " BE1 ", " BE2 ", " BE3 ", " BE4 "]
timings = ["8:00-9:00", "9:00-10:00", "10:15-11:15", "11:15-12:15", "13:00-14:00", "14:00-15:00"]

def load(pdf):
    tables = tabula.read_pdf(pdf, pages='all')
    subjects = set()
    teachers = set()
    lab_rooms = set()
    df = tables[2]
    subs = df.iloc[:, 0].values
    for i in subs:
        if isinstance(i, str):
            subjects.add(i)
    teach = df.iloc[:, 1].values
    for i in teach:
        if isinstance(i, str):
            teachers.add(i)
    sub = df.iloc[:, 2].values
    for i in sub:
        if isinstance(i, str):
            if i != 'SUBJECT':
                if len(i) > 2:
                    subjects.add(i)
    teach = df.iloc[:, 3].values
    for i in teach:
        if isinstance(i, str) and i != 'STAFF':
            teachers.add(i)
    labs = df.iloc[:, 4].values
    for i in labs:
        if isinstance(i, str) and i != 'LAB':
            lab_rooms.add(i)
    subjects.add(tables[4].columns[3])
    df = tables[4]
    teach = df.iloc[:, 5].values
    for i in teach:
        if isinstance(i, str):
            teachers.add(i)
    rooms = df.iloc[:, 6].values
    for i in rooms:
        if isinstance(i, str):
            lab_rooms.add(i)
    df = tables[1]
    return subjects, lab_rooms, teachers


ctk.set_default_color_theme("dark-blue")
ctk.set_appearance_mode("dark")


def Accumulate_Cells(table):  # Accumulating Cells Data into timetableCells
    timetableCells = table
    # Filling the clash dictionary
    clash = {}
    for r in range(1, len(timetableCells) + 1):
        clash[r] = {}

    for r in range(1, len(timetableCells) + 1):  # Slot-wise saving in the dictionary
        for c in range(len(timetableCells[r])):
            w = timetableCells[r][c]
            teacherName = str(w.getTeacherName())
            roomNumber = str(w.getRoomNumber())

            if teacherName not in clash[r]:
                if (teacherName != ""):
                    clash[r][teacherName] = [c]
            else:
                if (teacherName != ""):
                    clash[r][teacherName].append(c)

            if roomNumber not in clash[r]:
                if (roomNumber != ""):
                    clash[r][roomNumber] = [c]
            else:
                if (roomNumber != ""):
                    clash[r][roomNumber].append(c)

    return clash


def Validation_Algorithm(clash):
    invalidCells = []
    for r in range(1, len(clash) + 1):
        for c in clash[r]:
            if len(clash[r][c]) > 1:
                invalidCells.append([r, clash[r][c]])
    return invalidCells


class MyWidget(tk.Frame):
    def __init__(self, pos,subject, master=None):
        super().__init__(master, background="white")
        self.label1 = ctk.CTkLabel(master, text=subject, fg_color="#4bcffa", text_color="black", padx=10, pady=5,
                                   corner_radius=8, font=("Britannic Bold", 12))  # Subject Name
        self.label2 = ctk.CTkLabel(master, text="", fg_color="#B9E9FC", text_color="black", padx=10, pady=5,
                                   corner_radius=8, font=("Britannic Bold", 12))  # Teacher Name
        self.label3 = ctk.CTkLabel(master, text="", fg_color="#95BDFF", text_color="black", padx=10, pady=5,
                                   corner_radius=8, font=("Britannic Bold", 12))  # Room Number
        self.postn = pos
        self.label1.grid(row=pos[0] + 0, column=pos[1] + 0)
        self.label2.grid(row=pos[0] + 1, column=pos[1] + 0)
        self.label3.grid(row=pos[0] + 2, column=pos[1] + 0)

    def setdrop(self, ttobj, master=None):
        self.label1.bind("<Button-1>", ttobj.on_cell_click)
        self.label2.bind("<Button-1>", ttobj.on_cell_click)
        self.label3.bind("<Button-1>", ttobj.on_cell_click)

    def myconfigure(self, value, coord):
        if coord[0] == self.postn[0] + 0 and coord[1] == self.postn[1] + 0:
            self.label1.configure(text=value)
        if coord[0] == self.postn[0] + 1 and coord[1] == self.postn[1] + 0:
            self.label2.configure(text=value)
        if coord[0] == self.postn[0] + 2 and coord[1] == self.postn[1] + 0:
            self.label3.configure(text=value)

    def getTeacherName(self):
        return self.label2._text

    def getRoomNumber(self):
        return self.label3._text

    def getSubject(self):
        return self.label1._text
    
    def showErrorCell(self):
        self.config(bg="red")

    def resetColor(self):
        self.config(bg="white")


class TimetableApp():
    def __init__(self, master): 
        self.master = master
        
        # Create a Canvas widget and a Frame widget to hold your app
        self.canvas = tk.Canvas(master)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        
        xscrollbar = tk.Scrollbar(self.master, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=xscrollbar.set)
        xscrollbar.grid(row=1, column=0, sticky='ew')

        # Create a vertical scrollbar and associate it with the canvas
        yscrollbar = tk.Scrollbar(self.master, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=yscrollbar.set)
        yscrollbar.grid(row=0, column=1, sticky='ns')

        self.app_frame = tk.Frame(self.canvas)
        
        self.canvas.create_window((0,0), window=self.app_frame, anchor='nw')
        self.app_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
        self.app_frame.bind('<MouseWheel>', self._on_mousewheel)
        self.app_frame.bind('<Configure>', lambda event: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

        tup=load('TECOMP AllClass TT-1.pdf')
        self.subjects = list(tup[0])
        self.subject_widgets = []
        self.teachers = list(tup[2])
        self.teacher_widgets = []
        self.rooms = list(tup[1])
        self.room_widgets = []
        total_classes = len(classes)
        no_courses = len(self.subjects)
        slots = 6
        total_days = 1
        daily_repetition = 3
        table = GenerateTimeTable(total_classes, no_courses, slots, total_days, daily_repetition)
        self.passedSubjects = []
        for row in table.run():
            for col in row:
                self.passedSubjects.append(col)

        Menu_Day = ctk.StringVar(value="Select Day!")
        drop = ctk.CTkOptionMenu(master=self.app_frame, variable=Menu_Day,
                                 values=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"])
        drop.grid(row=0, column=2, sticky="ew", padx=5, pady=5)

        self.Class_select = ctk.StringVar(value="Select Class to export!")
        drop = ctk.CTkOptionMenu(master=self.app_frame, variable=self.Class_select,
                                 values=classes)
        drop.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        text_var = tk.StringVar(value="Teachers and Rooms")
        label1 = ctk.CTkLabel(master=self.app_frame,textvariable=text_var,width=50,height=25,fg_color="transparent",text_color="#3A58FF",corner_radius=8,font=("Britannic Bold", 14))
        label1.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

        subject_var = tk.StringVar(value="Subjects")
        label2 = ctk.CTkLabel(master=self.app_frame,textvariable=subject_var,width=50,height=25,fg_color="transparent",text_color="#3A58FF",corner_radius=8,font=("Britannic Bold", 14))
        label2.grid(row=2, column=1, sticky="ew", padx=5, pady=5)


        ValiateButton = ctk.CTkButton(master=self.app_frame, text="Validate", border_width=2, fg_color="#C7E8CA", hover_color="#C7FFCA",
                                      text_color="black", border_color="#5D9C59", command=self.on_button_validation,
                                      height=28, width=100, font=('Arial Bold', 15), corner_radius=8)
        ExportButton = ctk.CTkButton(master=self.app_frame, text="Export", border_width=2, fg_color="#C7E8CA",hover_color="#C7FFCA",
                                      text_color="black", border_color="#5D9C59", command=self.exportPDF,
                                      height=28, width=100, font=('Arial Bold', 15), corner_radius=8)
        '''ImportButton = ctk.CTkButton(master=self.app_frame, text="Import", border_width=2, fg_color="#C7E8CA",hover_color="#C7FFCA",
                                      text_color="black", border_color="#5D9C59",
                                      height=28, width=100, font=('Arial Bold', 15), corner_radius=8)'''
        for i, subject in enumerate(self.subjects):
            widget = ctk.CTkLabel(self.app_frame, text=subject, fg_color="#4bcffa", text_color="black", padx=10, pady=5,
                                  corner_radius=8, font=("Britannic Bold", 12))
            
            widget.grid(row=i + 3, column=1, sticky="ew", padx=2, pady=2)
            widget.bind("<Button-1>", self.on_subject_click)
            self.subject_widgets.append(widget)

        for i, teacher in enumerate(self.teachers):
            widget = ctk.CTkLabel(self.app_frame, text=teacher, fg_color="#B9E9FC", text_color="black", padx=10, pady=5,
                                  corner_radius=8, font=("Britannic Bold", 12))
            widget.grid(row=3+i, column=0, sticky="ew", padx=2, pady=2)
            widget.bind("<Button-1>", self.on_subject_click)
            self.teacher_widgets.append(widget)

        for i, room in enumerate(self.rooms):
            widget = ctk.CTkLabel(self.app_frame, text=room, fg_color="#95BDFF", text_color="black", padx=10, pady=5,
                                  corner_radius=8, font=("Britannic Bold", 12))
            widget.grid(row=i + len(self.teachers)  + 2, column=0, sticky="ew", padx=2, pady=2)
            widget.bind("<Button-1>", self.on_subject_click)
            self.room_widgets.append(widget)

        ValiateButton.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        #ImportButton.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        ExportButton.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.subject_widgets.append(ValiateButton)

        for i, day in enumerate(classes):
            widget = ctk.CTkLabel(self.app_frame, text=day, fg_color="#3A98B9", padx=10, pady=5, corner_radius=8,
                                  text_color="white", font=("Britannic Bold", 14))
            widget.grid(row=0, column=i + 4, sticky="ew", padx=5, pady=5)

        
        for i, section in enumerate(timings):
            widget = ctk.CTkLabel(self.app_frame, text=section, fg_color="#3A98B9", padx=10, pady=5, corner_radius=8,
                                  text_color="white", font=("Britannic Bold", 14))
            widget.grid(row=3 * i + 2, column=2, sticky="ew", padx=5, pady=5)

        self.table = {}
        for row in range(int(len(timings))):
            table_row = []
            for col in range(int(len(classes))):
                widget = MyWidget([3 * row + 1, col + 4], self.subjects[self.passedSubjects[col][row]-1], master=self.app_frame,)
                widget.grid(row=3 * row + 1, column=col + 4,
                            sticky="nsew", padx=5, pady=5, rowspan=3)
                widget.setdrop(self, master=self.app_frame)
                widget.bind("<Button-1>", self.on_cell_click)
                table_row.append(widget)
            self.table[row + 1] = table_row

    def _on_mousewheel(self, event):
         self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    def on_subject_click(self, event):
        subject = event.widget.cget("text")
        event.widget.configure(bg="white")
        self.dragged_subject = subject

    def on_cell_click(self, event):
        row = event.widget.master.grid_info()["row"]
        col = event.widget.master.grid_info()["column"]
        if not hasattr(self, "dragged_subject"):
            self.table[math.ceil(row / 3)][col - 4].myconfigure(value="", coord=[row, col])
            return
        if (row % 3 == 1 and self.dragged_subject not in self.subjects):
            return
        if (row % 3 == 2 and self.dragged_subject not in self.teachers):
            return
        if (row % 3 == 0 and self.dragged_subject not in self.rooms):
            return
        self.table[math.ceil(row / 3)][col - 4].myconfigure(value=self.dragged_subject, coord=[row, col])
        delattr(self, "dragged_subject")

    def on_button_validation(self):
        clash = Accumulate_Cells(self.table)
        invalidCells = Validation_Algorithm(clash)

        for r in self.table:
            for widget in self.table[r]:
                widget.resetColor()

        for r in invalidCells:
            for c in r[1]:
                self.table[r[0]][c].showErrorCell()
                
    def exportPDF(self):
        class_name=str(self.Class_select.get())
        column=classes.index(class_name)
        flag =True
        export_data=[]
        
        for row in range(1,len(self.table)+1):
            if(row==3 or row==5):
                export_data.append(["Break"])
            l1=[timings[((row-1)%6)],str(self.table[row][column].getSubject())]
            l2=['',str(self.table[row][column].getTeacherName())]
            l3=['',str(self.table[row][column].getRoomNumber())]
            export_data.append(l1)
            export_data.append(l2)
            export_data.append(l3)
        export_data.insert(0,['Timing','Monday','Tuesday','Wednesday','Thursday','Friday'])
        with open(csv_file, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            for sublist in export_data:
                writer.writerow(sublist)
        data = []
        with open(csv_file, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                data.append(row)
        pdf_file = 'Time_Table_'+datetime.datetime.now().strftime("%Y%m%d%H%M%S")+'_.pdf'
        doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
        table = Table(data)
        table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), '#4bcffa'),  # Header row background color
        ('TEXTCOLOR', (0, 0), (-1, 0), '#000000'),  # Header row text color
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header row font
        ('FONTSIZE', (0, 0), (-1, 0), 12),  # Header row font size
        ('BOTTOMPADDING', (0, 0), (-1, 0), 15),  # Header row bottom padding
        ('BACKGROUND', (0, 1), (-1, -1), '#EAEAEA'),  # Other row background color
        ('GRID', (0, 0), (-1, -1), 1, '#000000'),  # Grid color
        ('GRID', (0, 0), (-1, 0), 1, '#000000'),  # Header row grid color
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alignment
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Other row font
        ('FONTSIZE', (0, 1), (-1, -1), 10),  # Other row font size
        ('LEFTPADDING', (0, 0), (-1, -1), 3),  # Cell left padding
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),  # Cell right padding
        ('TOPPADDING', (0, 0), (-1, -1), 3),  # Cell top padding
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),  # Cell bottom padding
        ('SPAN', (0, 7), (5, 7)),
        ('SPAN', (0, 14), (5, 14)),
        ('BACKGROUND', (0, 7), (5, 7), '#B9E9FC'),  # Break Color
        ('BACKGROUND', (0, 14), (5, 14), '#B9E9FC'),  # Break Color
        ('FONTNAME', (0, 7), (5, 7), 'Helvetica-Bold'),  # Break
        ('FONTNAME', (0, 14), (5, 14), 'Helvetica-Bold'),  
        ]))
        
        sample_style_sheet = getSampleStyleSheet()
        paragraph_1 = Paragraph("----------------------------------------"+class_name+" Time-Table"+"----------------------------------------", sample_style_sheet['Heading1'])
        doc.title='Time Table'
        doc.build([paragraph_1,table])
        messagebox.showinfo("Alert", "The timetable document has been generated , kindly check the folder where this app resides !")

    def run(self):
        self.master.mainloop()


root = ctk.CTk()
app = TimetableApp(root)
width= root.winfo_screenwidth()
height= root.winfo_screenheight()
root.geometry("%dx%d" % (width, height))
root.title("ClashFree")
app.run()

