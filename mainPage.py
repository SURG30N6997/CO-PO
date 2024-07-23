import customtkinter as ctk 
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog
import openpyxl
from openpyxl.styles import Alignment 
from openpyxl import Workbook

class User_mode:
    def __init__(self):
        self.app = None
        self.cursor = None
        self.open_main_page()
    
    def open_main_page(self):
        ctk.set_appearance_mode("dark")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
        
        
        self.app = ctk.CTk()  # creating custom tkinter window
        #self.app.geometry("700x540")
        self.app.title('CO-PO')
         # Get screen width and height
        # screen_width = self.app.winfo_screenwidth()
        # screen_height = self.app.winfo_screenheight()
       
        screen_width=1500
        screen_height=800
       
        # Calculate the coordinates for centering the window
        x_position = 0
        y_position = 0
        
        # Set the window position and size
        self.app.geometry(f"{screen_width}x{screen_height}+{x_position}+{y_position}")
        
        self.main_frame = ctk.CTkFrame(master=self.app)
        self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=1)

        def switch():
            tabview.set(" CO Mapping ")
        
        def download():
            # Get the values from the Entry widgets
            values = [entry1.get(), entry2.get(), entry3.get(), entry4.get(), entry5.get(),
                    a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),
                    a2aT.get(), a2bT.get(), a3aT.get(), a3bT.get(),entry7.get(),entry8.get(),entry10.get(),entry11.get(),entry12.get()]
            
            coValues =[entry13.get(),entry14.get(),entry15.get(),q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
            coCAs=[entry13.get(),entry14.get(),entry15.get()]
            coQuizs=[q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
            
            # Check if any entry is empty
            if any(value == "" for value in values):
                CTkMessagebox(title="Error", message="Please fill in all fields.", icon="cancel")
            elif entry10.get()=="Quiz-Only":
                if any(coQuiz == "" for coQuiz in coQuizs):
                    CTkMessagebox(title="Error", message="Please fill in all fields.", icon="cancel")
                else :
                    import template_generator
                    template_generator.template_gen(values,coValues)
                    CTkMessagebox(message="Excel template downloaded successfully .",
                    icon="check", option_1="OK")
            elif entry10.get()=="CA-Only":
                if any(coCA == "" for coCA in coCAs):
                    CTkMessagebox(title="Error", message="Please fill in all fields.", icon="cancel")
                else :
                    import template_generator
                    template_generator.template_gen(values,coValues)
                    CTkMessagebox(message="Excel template downloaded successfully .",
                    icon="check", option_1="OK")
            else :
                if any(coValue == "" for coValue in coValues):
                    CTkMessagebox(title="Error", message="Please fill in all fields.", icon="cancel")
                else :
                    import template_generator
                    template_generator.template_gen(values,coValues)
                    CTkMessagebox(message="Excel template downloaded successfully .",
                    icon="check", option_1="OK")
            # else :
            #     import template_generator
            #     template_generator.template_gen(values)
            #     CTkMessagebox(message="Excel template downloaded successfully .",
            #       icon="check", option_1="OK")
                
        def ca1(option):
            if option == "NPTEL Course" or option == "Presentation":
                q1TCA1.configure(state="disabled", fg_color="gray")
                q2TCA1.configure(state="disabled", fg_color="gray")
                q3TCA1.configure(state="disabled", fg_color="gray")
                q4TCA1.configure(state="disabled", fg_color="gray")
                q5TCA1.configure(state="disabled", fg_color="gray")
                q6TCA1.configure(state="disabled", fg_color="gray")
                q7TCA1.configure(state="disabled", fg_color="gray")
                q8TCA1.configure(state="disabled", fg_color="gray")
                q9TCA1.configure(state="disabled", fg_color="gray")
                q10TCA1.configure(state="disabled", fg_color="gray")

                nptelCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

            else:
                nptelCA1Text.configure(state="disabled", fg_color="gray") 

                q1TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q2TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q3TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q4TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q5TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q6TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q7TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q8TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q9TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q10TCA1.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
               
        tabview = ctk.CTkTabview(self.main_frame,corner_radius=20)
        tabview.pack(expand=True, fill="both", padx=10, pady=5)

        tabview.add(" Instructions ") 
        tabview.add(" Basic Information ")  # add tab at the end
       
        tabview.add(" CO Mapping ")
        tabview.add(" Upload Excel File ")  # add tab at the end
        tabview.set(" Basic Information ")  # set currently visible tab

        button = ctk.CTkButton(master=tabview.tab(" Basic Information "),text=" Next ",font=("Arial",20),command=switch)
        button.place(x=725,y=600)
        
        label0=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Basic Details",font=("Arial",20))
        label0.place(x=725,y=5)
        
        label1=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="No of Students :",font=("Arial",15))
        label1.place(x=200,y=55)
        
        entry1=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="Enter no of students",font=("Arial",15),width=300)
        entry1.place(x=400,y=55)
        
        newLabel= ctk.CTkLabel(master= tabview.tab(" Basic Information "), text="Select year :", font=("Arial",15))
        newLabel.place(x = 200, y = 105)

        def semester(option):
            if option == "I":
                entry2.configure(values=["I","II"])
            elif option == "II":
                entry2.configure(values=["III","IV"])
            elif option == "III":
                entry2.configure(values=["V","VI"])
            elif option == "IV":
                entry2.configure(values=["VII","VIII"])

        def subject(option):
            if option == "I":
                entry3.configure(values=["Universal Human Values - 1","Fundamentals of Vedic Mathematics (Indian Knowledge System)", "Basic Electrical Engineering", "Engineering Drawing", "Engineering Mechanics", "Engineering Physics", "Matrices and Differential Calculus", "Python Programming"])
            elif option == "II":
                entry3.configure(values=["Universal Human Values - 2","Basic Workshop Practice", "Computer Programming", "Integral Calculus and Complex Numbers", "Biology for Engineers", "Engineering Chemistry", "Professional Communication and Ethics - 1"])
            elif option == "III":
                entry3.configure(values=["Engineering Mathematics III", "Data Structures and Analysis", "Database Management System", "Principle of Communications", "Paradigm and computer programming fundamentals"])
            elif option == "IV":
                entry3.configure(values=["Engineering Mathematics IV", "Computer Network and Network Design", "Operating System", "Automata Theory", "Computer Organization and Architecture"])
            elif option == "V":
                entry3.configure(values=["Internet Programming", "Computer Network Security", "Entrepreneurship and E- business", "Software Engineering", "Advance Data Management Technologies", "Advanced Data structure and Analysis"])
            elif option == "VI":
                entry3.configure(values=["Data Mining & Business Intelligence", "Web X.0", "Wireless Technology", "AI and DS – 1", "Optional Course 2"])
            elif option == "VII":
                entry3.configure(values=["AI and DS – II", "Internet of Everything", "Department Optional Course – 3", "Department Optional Course – 4", "Institute Optional Course – 1"])
            elif option == "VIII":
                entry3.configure(values=["Blockchain and DLT", "Department Optional Course – 5", "Department Optional Course – 6", "Institute Optional Course – 2"])

        yearDropDown = ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["I","II","III","IV"],font=("Arial",15),width=300,command=semester)
        yearDropDown.place(x=400,y=105)

        label2=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Semester :",font=("Arial",15))
        label2.place(x=200,y=155)
        
        entry2=ctk.CTkOptionMenu(master= tabview.tab(" Basic Information "), values=["I","II"],font=("Arial", 15), width=300, command=subject)
        entry2.place(x=400,y=155)
        
        label3=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Subject :",font=("Arial",15))
        label3.place(x=200,y=205)
        
        entry3=ctk.CTkOptionMenu(master= tabview.tab(" Basic Information "), values=["Universal Human Values - 1","Fundamentals of Vedic Mathematics (Indian Knowledge System)", "Basic Electrical Engineering", "Engineering Drawing", "Engineering Mechanics", "Engineering Physics", "Matrices and Differential Calculus", "Python Programming"],font=("Arial", 15), width=300)
        entry3.place(x=400,y=205)
        
        label4=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Academic Year :",font=("Arial",15))
        label4.place(x=200,y=255)
    
        entry4=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="YYYY-YYYY",font=("Arial",15),width=300)
        entry4.place(x=400,y=255)
        
        label5=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Subject Teacher :",font=("Arial",15))
        label5.place(x=200,y=305)
    
        entry5=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="Subject Teacher",font=("Arial",15),width=300)
        entry5.place(x=400,y=305)
        
        label7=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Class :",font=("Arial",15))
        label7.place(x=200,y=355)
    
        entry7=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="Eg.D10 C",font=("Arial",15),width=300)
        entry7.place(x=400,y=355)
        
        label8=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Department :",font=("Arial",15))
        label8.place(x=875,y=55)
    
        entry8=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="Information Technology",font=("Arial",15),width=300)
        entry8.place(x=1075,y=55)
        
        label11=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="EndSem CO's :",font=("Arial",15))
        label11.place(x=875,y=105)
    
        entry11=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=300)
        entry11.place(x=1075,y=105)
        
        label12=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Attainment Target :",font=("Arial",15))
        label12.place(x=875,y=155)
    
        entry12=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="52.5",font=("Arial",15),width=300)
        entry12.place(x=1075,y=155)
        
        label13=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="CA1 type :",font=("Arial",15))
        label13.place(x=875,y=255)
    
        entry13=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Quiz", "NPTEL Course", "Presentation"],font=("Arial",15),width=300,command=ca1)
        entry13.place(x=1075,y=255)
        
        label14=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="CA2 type :",font=("Arial",15))
        label14.place(x=875,y=305)
    
        entry14=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Quiz", "NPTEL Course", "Presentation"],font=("Arial",15),width=300)
        entry14.place(x=1075,y=305)
        
        label15=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="CA3 type :",font=("Arial",15))
        label15.place(x=875,y=355)
    
        entry15=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Quiz", "NPTEL Course", "Presentation"],font=("Arial",15),width=300, state="disabled", fg_color='gray')
        entry15.place(x=1075,y=355)

        nptel = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CO's for NPTEL/Presentation (CA)", font=("Arial", 20))
        nptel.place(x=650, y=425)

        nptelCA1Label = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CA1: ", font=("Arial", 15))
        nptelCA1Label.place(x=100, y=500)

        nptelCA1Text = ctk.CTkEntry(master=tabview.tab(" Basic Information "), placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=300)
        nptelCA1Text.place(x=200, y=500)

        nptelCA2Label = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CA2: ", font=("Arial", 15))
        nptelCA2Label.place(x=550, y=500)

        nptelCA2Text = ctk.CTkEntry(master=tabview.tab(" Basic Information "), placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=300)
        nptelCA2Text.place(x=650, y=500)

        nptelCA3Label = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CA3: ", font=("Arial", 15))
        nptelCA3Label.place(x=1000, y=500)

        nptelCA3Text = ctk.CTkEntry(master=tabview.tab(" Basic Information "), placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=300)
        nptelCA3Text.place(x=1100, y=500)
        
        label6=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="COs for Midterm",font=("Arial",20))
        label6.place(x=375,y=20)
    
        # entry6=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="Title",font=("Arial",20),width=300)
        # entry6.place(x=750,y=200)
        
        a1L=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="1a :",font=("Arial",15))
        a1L.place(x=200,y=60)
    
        a1T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a1T.place(x=250,y=60)
        
        a2L=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="1b :",font=("Arial",15))
        a2L.place(x=200,y=110)
    
        a2T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a2T.place(x=250,y=110)
        
        a3L=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="1c :",font=("Arial",15))
        a3L.place(x=200,y=160)
    
        a3T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a3T.place(x=250,y=160)
        
        a4L=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="1d :",font=("Arial",15))
        a4L.place(x=200,y=210)
    
        a4T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a4T.place(x=250,y=210)
    
        a5L=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="1e :",font=("Arial",15))
        a5L.place(x=200,y=260)

        a5T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a5T.place(x=250,y=260)
        
        a6L=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="1f :",font=("Arial",15))
        a6L.place(x=500,y=60)

        a6T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a6T.place(x=550,y=60)
        
        a2aL=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="2a :",font=("Arial",15))
        a2aL.place(x=500,y=110)
    
        a2aT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a2aT.place(x=550,y=110)
        
        a2bL=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="2b :",font=("Arial",15))
        a2bL.place(x=500,y=160)
    
        a2bT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a2bT.place(x=550,y=160)
        
        a3aL=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="3a :",font=("Arial",15))
        a3aL.place(x=500,y=210)
    
        a3aT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a3aT.place(x=550,y=210)
        
        a3bL=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="3b :",font=("Arial",15))
        a3bL.place(x=500,y=260)
    
        a3bT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        a3bT.place(x=550,y=260)
        
        #For Quiz

        label9=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="COs for CA1 Quiz",font=("Arial",20))
        label9.place(x=375,y=340)

        q1LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q1 :",font=("Arial",15))
        q1LCA1.place(x=200,y=380)
    
        q1TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q1TCA1.place(x=250,y=380)
        
        q2LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q2 :",font=("Arial",15))
        q2LCA1.place(x=200,y=430)
    
        q2TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q2TCA1.place(x=250,y=430)
        
        q3LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q3 :",font=("Arial",15))
        q3LCA1.place(x=200,y=480)
    
        q3TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q3TCA1.place(x=250,y=480)
        
        q4LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q4 :",font=("Arial",15))
        q4LCA1.place(x=200,y=530)
    
        q4TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q4TCA1.place(x=250,y=530)
    
        q5LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q5 :",font=("Arial",15))
        q5LCA1.place(x=200,y=580)

        q5TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q5TCA1.place(x=250,y=580)
        
        q6LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q6 :",font=("Arial",15))
        q6LCA1.place(x=500,y=380)

        q6TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q6TCA1.place(x=550,y=380)
        
        q7LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q7 :",font=("Arial",15))
        q7LCA1.place(x=500,y=430)
    
        q7TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q7TCA1.place(x=550,y=430)
        
        q8LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q8 :",font=("Arial",15))
        q8LCA1.place(x=500,y=480)
    
        q8TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q8TCA1.place(x=550,y=480)
        
        q9LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q9 :",font=("Arial",15))
        q9LCA1.place(x=500,y=530)
    
        q9TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q9TCA1.place(x=550,y=530)
        
        q10LCA1=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q10 :",font=("Arial",15))
        q10LCA1.place(x=500,y=580)
    
        q10TCA1=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q10TCA1.place(x=550,y=580)

        label18=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="COs for CA2 Quiz",font=("Arial",20))
        label18.place(x=1025,y=20)
        
        q1LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q1 :",font=("Arial",15))
        q1LCA2.place(x=850,y=60)
    
        q1TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q1TCA2.place(x=900,y=60)
        
        q2LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q2 :",font=("Arial",15))
        q2LCA2.place(x=850,y=110)
    
        q2TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q2TCA2.place(x=900,y=110)
        
        q3LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q3 :",font=("Arial",15))
        q3LCA2.place(x=850,y=160)
    
        q3TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q3TCA2.place(x=900,y=160)
        
        q4LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q4 :",font=("Arial",15))
        q4LCA2.place(x=850,y=210)
    
        q4TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q4TCA2.place(x=900,y=210)
    
        q5LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q5 :",font=("Arial",15))
        q5LCA2.place(x=850,y=260)

        q5TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q5TCA2.place(x=900,y=260)
        
        q6LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q6 :",font=("Arial",15))
        q6LCA2.place(x=1150,y=60)

        q6TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q6TCA2.place(x=1200,y=60)
        
        q7LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q7 :",font=("Arial",15))
        q7LCA2.place(x=1150,y=110)
    
        q7TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q7TCA2.place(x=1200,y=110)
        
        q8LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q8 :",font=("Arial",15))
        q8LCA2.place(x=1150,y=160)
    
        q8TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q8TCA2.place(x=1200,y=160)
        
        q9LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q9 :",font=("Arial",15))
        q9LCA2.place(x=1150,y=210)
    
        q9TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q9TCA2.place(x=1200,y=210)
        
        q10LCA2=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q10 :",font=("Arial",15))
        q10LCA2.place(x=1150,y=260)
    
        q10TCA2=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q10TCA2.place(x=1200,y=260)

        label21=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="COs for CA3 Quiz",font=("Arial",20))
        label21.place(x=1025,y=340)

        q1LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q1 :",font=("Arial",15))
        q1LCA3.place(x=850,y=380)
    
        q1TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q1TCA3.place(x=900,y=380)
        
        q2LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q2 :",font=("Arial",15))
        q2LCA3.place(x=850,y=430)
    
        q2TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q2TCA3.place(x=900,y=430)
        
        q3LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q3 :",font=("Arial",15))
        q3LCA3.place(x=850,y=480)
    
        q3TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q3TCA3.place(x=900,y=480)
        
        q4LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q4 :",font=("Arial",15))
        q4LCA3.place(x=850,y=530)
    
        q4TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q4TCA3.place(x=900,y=530)
    
        q5LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q5 :",font=("Arial",15))
        q5LCA3.place(x=850,y=580)

        q5TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q5TCA3.place(x=900,y=580)
        
        q6LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q6 :",font=("Arial",15))
        q6LCA3.place(x=1150,y=380)

        q6TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q6TCA3.place(x=1200,y=380)
        
        q7LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q7 :",font=("Arial",15))
        q7LCA3.place(x=1150,y=430)
    
        q7TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q7TCA3.place(x=1200,y=430)
        
        q8LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q8 :",font=("Arial",15))
        q8LCA3.place(x=1150,y=480)
    
        q8TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q8TCA3.place(x=1200,y=480)
        
        q9LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q9 :",font=("Arial",15))
        q9LCA3.place(x=1150,y=530)
    
        q9TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q9TCA3.place(x=1200,y=530)
        
        q10LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q10 :",font=("Arial",15))
        q10LCA3.place(x=1150,y=580)
    
        q10TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q10TCA3.place(x=1200,y=580)
        

        
        # def disable(option):
        #     if option=="Quiz-Only":
        #         q1T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q2T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q3T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q4T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q5T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q6T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q2aT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q2bT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q3aT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q3bT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                
        #         entry13.configure(state="disabled",fg_color='gray')
        #         entry14.configure(state="disabled",fg_color='gray')
        #         entry15.configure(state="disabled",fg_color='gray')
                
        #     elif option=="CA-Only":
        #         entry13.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         entry14.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         entry15.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                
        #         q1T.configure(state="disabled",fg_color='gray')
        #         q2T.configure(state="disabled",fg_color='gray')
        #         q3T.configure(state="disabled",fg_color='gray')
        #         q4T.configure(state="disabled",fg_color='gray')
        #         q5T.configure(state="disabled",fg_color='gray')
        #         q6T.configure(state="disabled",fg_color='gray')
        #         q2aT.configure(state="disabled",fg_color='gray')
        #         q2bT.configure(state="disabled",fg_color='gray')
        #         q3aT.configure(state="disabled",fg_color='gray')
        #         q3bT.configure(state="disabled",fg_color='gray')
        #     else :
        #         entry13.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         entry14.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         entry15.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                
        #         q1T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q2T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q3T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q4T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q5T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q6T.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q2aT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q2bT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q3aT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         q3bT.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        
        def disable(option):
            if option == "3":
                entry15.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                q1TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q2TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q3TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q4TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q5TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q6TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q7TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q8TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q9TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                q10TCA3.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                nptelCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            else:
                entry15.configure(state="disabled", fg_color='gray')
                q1TCA3.configure(state="disabled", fg_color="gray")
                q2TCA3.configure(state="disabled", fg_color="gray")
                q3TCA3.configure(state="disabled", fg_color="gray")
                q4TCA3.configure(state="disabled", fg_color="gray")
                q5TCA3.configure(state="disabled", fg_color="gray")
                q6TCA3.configure(state="disabled", fg_color="gray")
                q7TCA3.configure(state="disabled", fg_color="gray")
                q8TCA3.configure(state="disabled", fg_color="gray")
                q9TCA3.configure(state="disabled", fg_color="gray")
                q10TCA3.configure(state="disabled", fg_color="gray")

                nptelCA3Text.configure(state="disabled", fg_color="gray") 

                
                
        label10=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="No. of CA's :",font=("Arial",15))
        label10.place(x=875,y=205)
    
        entry10=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["2", "3"],font=("Arial",15),width=300,command=disable)
        entry10.place(x=1075,y=205)
        
        def upload_file():
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            path_entry.delete(0, ctk.END)  # Clear any existing text in the entry widget
            path_entry.insert(0, file_path)  # Insert the file path into the entry widget

         
        def process_file():
            file_path = path_entry.get()
            import Cal
            Cal.cal_sheet(file_path)
            # print(file_path) 
            # workbook=openpyxl.load_workbook(file_path)
            # sheet=workbook['Sheet1']

            # for col in range (2,16):
            #     column_letter = openpyxl.utils.get_column_letter(col)
            #     sheet[f'{column_letter}71']=f'=COUNT({column_letter}7:{column_letter}70)'

            # for row in range (7,71):    
            #     sheet[f'H{row}']=f'=ROUND(SUM(B{row}:G{row}),0)'

            # for row in range (7,71):    
            #     sheet[f'K{row}']=f'=ROUND(SUM(I{row}:J{row}),0)'
                
            # for row in range (7,71):    
            #     sheet[f'N{row}']=f'=ROUND(SUM(L{row}:M{row}),0)'
                    
            # for row in range (7,71):    
            #     sheet[f'O{row}']=f'=ROUND(SUM(H{row},K{row},N{row}),0)'

            # for col in range (2,16):
            #     column_letter = openpyxl.utils.get_column_letter(col)
            #     sheet[f'{column_letter}72']=f'=ROUND(AVERAGE({column_letter}7:{column_letter}70),0)'
            
            # for col in range (2,8):
            #     column_letter = openpyxl.utils.get_column_letter(col)
            #     sheet[f'{column_letter}73']=f'=COUNTIF({column_letter}7:{column_letter}70,">=1.05")'

            # for col in range (8,16):
            #     column_letter = openpyxl.utils.get_column_letter(col)
            #     sheet[f'{column_letter}73']=f'=COUNTIF({column_letter}7:{column_letter}70,">=2.625")'

            # for col in range (2,16):
            #     column_letter = openpyxl.utils.get_column_letter(col)
            #     sheet[f'{column_letter}74']=f'=ROUND(({column_letter}73/{column_letter}71)*100,1)'


            # for col in range (2,16):
            #     column_letter = openpyxl.utils.get_column_letter(col)
            #     sheet[f'{column_letter}75']=f'=COUNTIF({column_letter}7:{column_letter}70,">="&{column_letter}72)'


            # for col in range (2,16):
            #     column_letter = openpyxl.utils.get_column_letter(col)
            #     sheet[f'{column_letter}76']=f'=IF({column_letter}74<60,1,IF(AND({column_letter}74>59,{column_letter}74<70),2,IF(AND({column_letter}74>69,{column_letter}74<80),3,4)))'
                


            # columns_with_1 = []
            # columns_with_2 = []
            # columns_with_3 = []
            # columns_with_4 = []
            # columns_with_5 = []

            # column_range = ['B', 'C', 'D', 'E', 'F', 'G', 'I', 'J', 'L', 'M']

            # # Iterate through cells in the specified column range
            # for column_letter in column_range:
            #     # Get the cell in row 2 corresponding to the column letter
            #     cell = sheet[f"{column_letter}6"]
            #     # Check if the cell has a value
            #     if cell.value :
            #         # values = [int(val.strip()) for val in str(cell.value).split(',') if val.strip().isdigit()]  #This for without CO like 1,2,3
            #         values = [int(val.strip()) for val in str(cell.value)[2:].split(',') if val.strip().isdigit()]   #This for with CO like CO1,2,3
            #         # Check if '1' is present in the list of values then 2 and 3,4and 5
            #         # print(values)
            #         if 1 in values:
            #             # Add the cell value to the list for calculation
            #             columns_with_1.append(column_letter)
            #         if 2 in values:
            #             # Add the cell value to the list for calculation
            #             columns_with_2.append(column_letter)
            #         if 3 in values:
            #             # Add the cell value to the list for calculation
            #             columns_with_3.append(column_letter)
            #         if 4 in values:
            #             # Add the cell value to the list for calculation
            #             columns_with_4.append(column_letter)
            #         if 5 in values:
            #             # Add the cell value to the list for calculation
            #             columns_with_5.append(column_letter)

            # # Calculate the average using Excel formula
            # if columns_with_1:
            #     average_formula = f"=ROUND(AVERAGE({','.join([f'{col_letter}76' for col_letter in columns_with_1])}),1)"
            #     sheet['G80'] = average_formula
            # else:
            #     sheet['G80'] = '-'
                
            # if columns_with_2:
            #     average_formula = f"=ROUND(AVERAGE({','.join([f'{col_letter}76' for col_letter in columns_with_2])}),1)"
            #     sheet['G81'] = average_formula
            # else:
            #     sheet['G81'] = '-'
                
            # if columns_with_3:
            #     average_formula = f"=ROUND(AVERAGE({','.join([f'{col_letter}76' for col_letter in columns_with_3])}),1)"
            #     sheet['G82'] = average_formula
            # else:
            #     sheet['G82'] = '-'
                
            # if columns_with_4:
            #     average_formula = f"=ROUND(AVERAGE({','.join([f'{col_letter}76' for col_letter in columns_with_4])}),1)"
            #     sheet['G83'] = average_formula
            # else:
            #     sheet['G83'] = '-'
                
            # if columns_with_5:
            #     average_formula = f"=ROUND(AVERAGE({','.join([f'{col_letter}76' for col_letter in columns_with_5])}),1)"
            #     sheet['G84'] = average_formula
            # else:
            #     sheet['G84'] = '-'
                
            # print(sheet['G84'].internal_value)
            # workbook.save(file_path)

        button_upload=ctk.CTkButton(tabview.tab(" Upload Excel File "),text="Upload",width=100,height=30,command=upload_file)
        button_upload.place(x=100,y=100)
        
        path_entry=ctk.CTkEntry(tabview.tab(" Upload Excel File "))
        
        button_process=ctk.CTkButton(tabview.tab(" Upload Excel File "),text="Process",width=100,height=30,command=process_file)
        button_process.place(x=500,y=100)
        self.app.mainloop()
        
if __name__ == "__main__":
    User_mode()
