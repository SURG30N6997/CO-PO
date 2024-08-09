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
        def switch():
            tabview.set(" CO Mapping ")
        
        def download():
            # Get the values from the Entry widgets
            # values = [entry1.get(), entry2.get(), entry3.get(), entry4.get(), entry5.get(),
            #         a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),
            #         a2aT.get(), a2bT.get(), a3aT.get(), a3bT.get(),entry7.get(),entry8.get(),entry10.get(),entry11.get(),entry12.get()]
            
            # coValues =[entry13.get(),entry14.get(),entry15.get(),q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
            # coCAs=[entry13.get(),entry14.get(),entry15.get()]
            # coQuizs=[q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
            
            values=[entry1.get(),entry8.get(),yearDropDown.get(),entry2.get(),entry3.get(),entry4.get(),entry5.get(),entry7.get(),entry11.get(),entry12.get(),

                    entry10.get(),entry13.get(),entry14.get(),entry15.get(),

                    noCA1Entry.get(),noCA2Entry.get(),noCA3Entry.get(),
                    nptelCA1Text.get(),nptelCA2Text.get(),nptelCA3Text.get(),


                    q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get(),
                    q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get(),q10TCA2.get(),
                    q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get(),q9TCA3.get(),q10TCA3.get(),

                    a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),a2aT.get(),a2bT.get(), a3aT.get(), a3bT.get()]
            
            basic_values=[entry1.get(),entry8.get(),yearDropDown.get(),entry2.get(),entry3.get(),entry4.get(),entry5.get(),entry7.get(),entry11.get(),entry12.get(),entry10.get(),entry13.get(),entry14.get(),entry15.get()]
            midSem_Co_values=[a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),a2aT.get(),a2bT.get(), a3aT.get(), a3bT.get()]
            
            if entry10.get()=="2":
                if entry13.get()=="Select Type":
                    CTkMessagebox(title="Error", message="Please Select Type of CA 1.", icon="cancel")
                elif entry13.get()=="Quiz":
                    if noCA1Entry.get()=="Select No" :
                        CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                    elif noCA1Entry.get()=="1" :
                        CA1_Co_arr=[q1TCA1.get()]
                    elif noCA1Entry.get()=="2" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get()]
                    elif noCA1Entry.get()=="3" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get()]
                    elif noCA1Entry.get()=="4" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get()]
                    elif noCA1Entry.get()=="5" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get()]
                    elif noCA1Entry.get()=="6" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get()]
                    elif noCA1Entry.get()=="7" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get()]
                    elif noCA1Entry.get()=="8" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get()]
                    elif noCA1Entry.get()=="9" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get()]
                    else:
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get()]       
                else:
                    CA1_Co_arr=[nptelCA1Text.get()]
                 
                        
                if entry14.get()=="Select Type":
                    CTkMessagebox(title="Error", message="Please Select Type of CA 2.", icon="cancel")
                elif entry14.get()=="Quiz":
                    if noCA2Entry.get()=="Select No" :
                        CTkMessagebox(title="Error", message="Please Select No Of Question in CA 2.", icon="cancel")
                    elif noCA2Entry.get()=="1" :
                        CA2_Co_arr=[q1TCA2.get()]
                    elif noCA2Entry.get()=="2" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get()]
                    elif noCA2Entry.get()=="3" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get()]
                    elif noCA2Entry.get()=="4" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get()]
                    elif noCA2Entry.get()=="5" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get()]
                    elif noCA2Entry.get()=="6" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get()]
                    elif noCA2Entry.get()=="7" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get()]
                    elif noCA2Entry.get()=="8" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get()]
                    elif noCA2Entry.get()=="9" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get()]
                    else :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get(),q10TCA2.get()]
                else:
                    CA2_Co_arr=[nptelCA2Text.get()]
                    
            else:
                
                if entry13.get()=="Select Type":
                    CTkMessagebox(title="Error", message="Please Select Type of CA 1.", icon="cancel")
                elif entry13.get()=="Quiz":
                    if noCA1Entry.get()=="Select No" :
                        CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                    elif noCA1Entry.get()=="1" :
                        CA1_Co_arr=[q1TCA1.get()]
                    elif noCA1Entry.get()=="2" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get()]
                    elif noCA1Entry.get()=="3" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get()]
                    elif noCA1Entry.get()=="4" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get()]
                    elif noCA1Entry.get()=="5" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get()]
                    elif noCA1Entry.get()=="6" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get()]
                    elif noCA1Entry.get()=="7" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get()]
                    elif noCA1Entry.get()=="8" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get()]
                    elif noCA1Entry.get()=="9" :
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get()]
                    else:
                        CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get()]
                else:
                    CA1_Co_arr=[nptelCA1Text.get()]
                
                    
                if entry14.get()=="Select Type":
                    CTkMessagebox(title="Error", message="Please Select Type of CA 2.", icon="cancel")
                elif entry14.get()=="Quiz":
                    if noCA2Entry.get()=="Select No" :
                        CTkMessagebox(title="Error", message="Please Select No Of Question in CA 2.", icon="cancel")
                    elif noCA2Entry.get()=="1" :
                        CA2_Co_arr=[q1TCA2.get()]
                    elif noCA2Entry.get()=="2" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get()]
                    elif noCA2Entry.get()=="3" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get()]
                    elif noCA2Entry.get()=="4" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get()]
                    elif noCA2Entry.get()=="5" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get()]
                    elif noCA2Entry.get()=="6" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get()]
                    elif noCA2Entry.get()=="7" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get()]
                    elif noCA2Entry.get()=="8" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get()]
                    elif noCA2Entry.get()=="9" :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get()]
                    else :
                        CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get(),q10TCA2.get()]
                else:
                    CA2_Co_arr=[nptelCA2Text.get()]
                    
                if entry15.get()=="Select Type":
                    CTkMessagebox(title="Error", message="Please Select Type of CA 3.", icon="cancel")
                elif entry15.get()=="Quiz":
                    if noCA3Entry.get()=="Select No" :
                        CTkMessagebox(title="Error", message="Please Select No Of Question in CA 3.", icon="cancel")
                    elif noCA3Entry.get()=="1" :
                        CA3_Co_arr=[q1TCA3.get()]
                    elif noCA3Entry.get()=="2" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get()]
                    elif noCA3Entry.get()=="3" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get()]
                    elif noCA3Entry.get()=="4" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get()]
                    elif noCA3Entry.get()=="5" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get()]
                    elif noCA3Entry.get()=="6" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get()]
                    elif noCA3Entry.get()=="7" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get()]
                    elif noCA3Entry.get()=="8" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get()]
                    elif noCA3Entry.get()=="9" :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get(),q9TCA3.get()]
                    else :
                        CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get(),q9TCA3.get(),q10TCA3.get()]
                else:
                    CA3_Co_arr=[nptelCA3Text.get()]
                    
            if basic_values[1]=="Select Department":
                CTkMessagebox(title="Error", message="Please Select Department.", icon="cancel")
            elif basic_values[2]=="Select Year":
                CTkMessagebox(title="Error", message="Please Select Year.", icon="cancel")
            elif basic_values[3]=="Select Sem":
                CTkMessagebox(title="Error", message="Please Select Semester.", icon="cancel")
            elif basic_values[4]=="Select Subject":
                CTkMessagebox(title="Error", message="Please Select Subject.", icon="cancel")
            elif any(value == "" for value in basic_values):
                CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
            elif any(midCo == "" for midCo in midSem_Co_values):
                CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
            elif any(ca1 == "" for ca1 in CA1_Co_arr):
                CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
            elif any(ca2 == "" for ca2 in CA2_Co_arr):
                CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
            elif entry10.get()=="3":
                if any(ca3 == "" for ca3 in CA3_Co_arr):
                    CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
                else :
                    import template_generator
                    template_generator.template_gen(basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr)
                    CTkMessagebox(message="Excel template downloaded successfully .",icon="check", option_1="OK")
            else:
                # import template_generator
                # print(entry10.get())
                if entry10.get()=="2": 
                    import template_generator
                    print("Hi v1",basic_values[10]) 
                    template_generator.template_gen(basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,[])
                    CTkMessagebox(message="Excel template downloaded successfully .",icon="check", option_1="OK")
                # elif entry10.get()=="3":
                #     print("Hi v2",basic_values[10]) 
                #     print("Hi v2",CA3_Co_arr) 
                #     template_generator.template_gen(basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr)
                #     CTkMessagebox(message="Excel template downloaded successfully .",icon="check", option_1="OK")
            
                                        
                
            
        def ca1(option):
            if  option == "Select Type":
                for disca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,noCA1Entry,nptelCA1Text]:
                    disca.configure(state="disabled", fg_color="gray")
          
            elif option == "NPTEL Course" or option == "Presentation":
                for disca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,noCA1Entry]:
                    disca.configure(state="disabled", fg_color="gray")
                nptelCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"]) 
        
            else:
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                noCA1Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                nptelCA1Text.configure(state="disabled", fg_color="gray") 
                
        def ca2(option):
            if  option == "Select Type":
                for disca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,noCA2Entry,nptelCA2Text]:
                    disca.configure(state="disabled", fg_color="gray")
                
                nptelCA2Text.configure(state="disabled", fg_color="gray")
                noCA2Entry.configure(state="disabled", fg_color="gray") 
            elif option == "NPTEL Course" or option == "Presentation":
                for disca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,noCA2Entry]:
                    disca.configure(state="disabled", fg_color="gray")
                nptelCA2Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"]) 
                
            else:
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                noCA2Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                nptelCA2Text.configure(state="disabled", fg_color="gray") 
                

        def ca3(option):
            if  option == "Select Type":
                for disca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,noCA3Entry,nptelCA3Text]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option == "NPTEL Course" or option == "Presentation":
                for disca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,noCA3Entry]:
                    disca.configure(state="disabled", fg_color="gray")
                nptelCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])    
            else:
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                noCA3Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                nptelCA3Text.configure(state="disabled", fg_color="gray") 
               
        def semester(option):
            if option == "Select Year":
                entry2.configure(values=["Select Sem"])
            elif option == "F.E":
                entry2.configure(values=["Select Sem","I","II"])
            elif option == "S.E":
                entry2.configure(values=["Select Sem","III","IV"])
            elif option == "T.E":
                entry2.configure(values=["Select Sem","V","VI"])
            elif option == "B.E":
                entry2.configure(values=["Select Sem","VII","VIII"])


        def noQuestion1(option):
            if option== "1" :
                for ca in [q1TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "2" :
                for ca in [q1TCA1,q2TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "3" :
                for ca in [q1TCA1,q2TCA1,q3TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "4" :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
              
            elif option== "5" :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
              
            elif option== "6" :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")

            elif option== "7" :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q8TCA1,q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option== "8" :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q9TCA1,q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "9" :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q10TCA1]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option=="10" :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                
             
            else :
                for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                    ca.configure(state="disabled", fg_color="gray")  
                      
        def noQuestion2(option):
            if option== "1" :
                for ca in [q1TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "2" :
                for ca in [q1TCA2,q2TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "3" :
                for ca in [q1TCA2,q2TCA2,q3TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "4" :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "5" :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "6" :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "7" :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q8TCA2,q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option== "8" :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q9TCA2,q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option== "9" :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q10TCA2]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option=="10" :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
            else :
                for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                    ca.configure(state="disabled", fg_color="gray")
                
        def noQuestion3(option):
            if option== "1" :
                for ca in [q1TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "2" :
                for ca in [q1TCA3,q2TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "3" :
                for ca in [q1TCA3,q2TCA3,q3TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "4" :
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "5" :
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option== "6" :
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option== "7" :
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q8TCA3,q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option== "8" :
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q9TCA3,q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
               
            elif option== "9" :
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                for disca in [q10TCA3]:
                    disca.configure(state="disabled", fg_color="gray")
                
            elif option=="10" :
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
               
            else:
                for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    ca.configure(state="disabled", fg_color="gray")
                
                
        def subject(option):
            if option == "Select Sem":
                entry3.configure(values=["Select Subject"])
            elif option == "I":
                entry3.configure(values=["Select Subject","Universal Human Values - 1","Fundamentals of Vedic Mathematics (Indian Knowledge System)", "Basic Electrical Engineering", "Engineering Drawing", "Engineering Mechanics", "Engineering Physics", "Matrices and Differential Calculus", "Python Programming"])
            elif option == "II":
                entry3.configure(values=["Select Subject","Universal Human Values - 2","Basic Workshop Practice", "Computer Programming", "Integral Calculus and Complex Numbers", "Biology for Engineers", "Engineering Chemistry", "Professional Communication and Ethics - 1"])
            elif option == "III":
                entry3.configure(values=["Select Subject","Engineering Mathematics III", "Data Structures and Analysis", "Database Management System", "Principle of Communications", "Paradigm and computer programming fundamentals"])
            elif option == "IV":
                entry3.configure(values=["Select Subject","Engineering Mathematics IV", "Computer Network and Network Design", "Operating System", "Automata Theory", "Computer Organization and Architecture"])
            elif option == "V":
                entry3.configure(values=["Select Subject","Internet Programming", "Computer Network Security", "Entrepreneurship and E- business", "Software Engineering", "Advance Data Management Technologies", "Advanced Data structure and Analysis"])
            elif option == "VI":
                entry3.configure(values=["Select Subject","Data Mining & Business Intelligence", "Web X.0", "Wireless Technology", "AI and DS 1", "Optional Course 2"])
            elif option == "VII":
                entry3.configure(values=["Select Subject","AI and DS II", "Internet of Everything", "Department Optional Course 3", "Department Optional Course 4", "Institute Optional Course 1"])
            elif option == "VIII":
                entry3.configure(values=["Select Subject","Blockchain and DLT", "Department Optional Course 5", "Department Optional Course 6", "Institute Optional Course 2"])

        def disable(option):
            if option == "3":
                for entry in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                    entry.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                entry15.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
               
            else:
                for entry in [entry15,q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,nptelCA3Text,noCA3Entry]:
                    entry.configure(state="disabled", fg_color='gray')
               
                
        ctk.set_appearance_mode("dark")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
        
        
        self.app = ctk.CTk()  # creating custom tkinter window
        self.app.title('CO-PO')
       
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
        
        newLabel= ctk.CTkLabel(master= tabview.tab(" Basic Information "), text="Year :", font=("Arial",15))
        newLabel.place(x = 200, y = 155)

        yearDropDown = ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select Year","F.E","S.E","T.E","B.E"],font=("Arial",15),width=300,command=semester)
        yearDropDown.place(x=400,y=155)

        label8=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Department :",font=("Arial",15))
        label8.place(x=200,y=105)
    
        entry8=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select Department","Humanities and Applied Science(FE)","Information Technology","Computer","AI and Data Science","Electronics and Telecommunication","Electronics","Instrumentation"],font=("Arial",15),width=300)
        entry8.place(x=400,y=105)

        label2=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Semester :",font=("Arial",15))
        label2.place(x=200,y=205)
        
        entry2=ctk.CTkOptionMenu(master= tabview.tab(" Basic Information "), values=["Select Sem"],font=("Arial", 15), width=300, command=subject)
        entry2.place(x=400,y=205)
        
        label3=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Subject :",font=("Arial",15))
        label3.place(x=200,y=255)
        
        entry3=ctk.CTkOptionMenu(master= tabview.tab(" Basic Information "), values=["Select Subject"],font=("Arial", 15), width=300)
        entry3.place(x=400,y=255)
        
        label4=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Academic Year :",font=("Arial",15))
        label4.place(x=200,y=305)
    
        entry4=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="YYYY-YYYY",font=("Arial",15),width=300)
        entry4.place(x=400,y=305)
        
        label5=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Subject Teacher :",font=("Arial",15))
        label5.place(x=200,y=355)
    
        entry5=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="Subject Teacher",font=("Arial",15),width=300)
        entry5.place(x=400,y=355)
        
        label7=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Class :",font=("Arial",15))
        label7.place(x=200,y=405)
    
        entry7=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="Eg.D10 C",font=("Arial",15),width=300)
        entry7.place(x=400,y=405)
        
        label11=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="EndSem CO's :",font=("Arial",15))
        label11.place(x=200,y=455)
    
        entry11=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=300)
        entry11.place(x=400,y=455)
        
        # label8=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Department :",font=("Arial",15))
        # label8.place(x=875,y=55)
    
        # entry8=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select Department","Humanities and Applied Science(FE)","Information Technology","Computer","AI and Data Science","Electronics and Telecommunication","Electronics","Instrumentation"],font=("Arial",15),width=300)
        # entry8.place(x=1075,y=55)
        
        label12=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="Attainment Target :",font=("Arial",15))
        label12.place(x=875,y=55)
    
        entry12=ctk.CTkEntry(master=tabview.tab(" Basic Information "),placeholder_text="52.5",font=("Arial",15),width=300)
        entry12.place(x=1075,y=55)

        label10=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="No. of CA's :",font=("Arial",15))
        label10.place(x=875,y=105)
    
        entry10=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["2", "3"],font=("Arial",15),width=300,command=disable)
        entry10.place(x=1075,y=105)
        
        
        label13=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="CA1 type :",font=("Arial",15))
        label13.place(x=875,y=155)
    
        entry13=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select Type","Quiz", "NPTEL Course", "Presentation"],font=("Arial",15),width=300,command=ca1)
        entry13.place(x=1075,y=155)
        
        noCA1Label=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="No of Question CA1 :",font=("Arial",15))
        noCA1Label.place(x=875,y=205)
        
        noCA1Entry=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select No","1","2","3","4","5","6","7","8","9","10"],font=("Arial",15),width=300,command=noQuestion1)
        noCA1Entry.configure(state="disabled", fg_color="gray") 
        noCA1Entry.place(x=1075,y=205)
        
        
        label14=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="CA2 type :",font=("Arial",15))
        label14.place(x=875,y=255)
    
        entry14=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select Type","Quiz", "NPTEL Course", "Presentation"],font=("Arial",15),width=300,command=ca2)
        entry14.place(x=1075,y=255)
        
        noCA2Label=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="No of Question CA2 :",font=("Arial",15))
        noCA2Label.place(x=875,y=305)
        
        noCA2Entry=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select No","1","2","3","4","5","6","7","8","9","10"],font=("Arial",15),width=300,command=noQuestion2)
        noCA2Entry.configure(state="disabled", fg_color="gray") 
        noCA2Entry.place(x=1075,y=305)
        
        label15=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="CA3 type :",font=("Arial",15))
        label15.place(x=875,y=355)
    
        entry15=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select Type","Quiz", "NPTEL Course", "Presentation"],font=("Arial",15),width=300, state="disabled", fg_color='gray', command=ca3)
        entry15.place(x=1075,y=355)
        
        noCA3Label=ctk.CTkLabel(master=tabview.tab(" Basic Information "),text="No of Question CA3 :",font=("Arial",15))
        noCA3Label.place(x=875,y=405)
        
        noCA3Entry=ctk.CTkOptionMenu(master=tabview.tab(" Basic Information "),values=["Select No","1","2","3","4","5","6","7","8","9","10"],font=("Arial",15),width=300,command=noQuestion3)
        noCA3Entry.configure(state="disabled", fg_color="gray") 
        noCA3Entry.place(x=1075,y=405)
        #print(noCA3Entry.get())

        nptel = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CO's for NPTEL/Presentation (CA)", font=("Arial", 20))
        nptel.place(x=650, y=505)

        nptelCA1Label = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CA1: ", font=("Arial", 15))
        nptelCA1Label.place(x=100, y=555)

        nptelCA1Text = ctk.CTkEntry(master=tabview.tab(" Basic Information "), placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=300)
        nptelCA1Text.configure(state="disabled", fg_color="gray")
        nptelCA1Text.place(x=200, y=555)

        nptelCA2Label = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CA2: ", font=("Arial", 15))
        nptelCA2Label.place(x=550, y=555)

        nptelCA2Text = ctk.CTkEntry(master=tabview.tab(" Basic Information "), placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=300)
        nptelCA2Text.configure(state="disabled", fg_color="gray")
        nptelCA2Text.place(x=650, y=555)

        nptelCA3Label = ctk.CTkLabel(master=tabview.tab(" Basic Information "), text="CA3: ", font=("Arial", 15))
        nptelCA3Label.place(x=1000, y=555)

        nptelCA3Text = ctk.CTkEntry(master=tabview.tab(" Basic Information "), placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=300)
        nptelCA3Text.configure(state="disabled", fg_color="gray")
        nptelCA3Text.place(x=1100, y=555)
        
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
        q1TCA3.configure(state="disabled", fg_color="gray")
        q1TCA3.place(x=900,y=380)
        
        q2LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q2 :",font=("Arial",15))
        q2LCA3.place(x=850,y=430)
    
        q2TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q2TCA3.configure(state="disabled", fg_color="gray")
        q2TCA3.place(x=900,y=430)
        
        q3LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q3 :",font=("Arial",15))
        q3LCA3.place(x=850,y=480)
    
        q3TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q3TCA3.configure(state="disabled", fg_color="gray")
        q3TCA3.place(x=900,y=480)
        
        q4LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q4 :",font=("Arial",15))
        q4LCA3.place(x=850,y=530)
    
        q4TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q4TCA3.configure(state="disabled", fg_color="gray")
        q4TCA3.place(x=900,y=530)
    
        q5LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q5 :",font=("Arial",15))
        q5LCA3.place(x=850,y=580)

        q5TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q5TCA3.configure(state="disabled", fg_color="gray")
        q5TCA3.place(x=900,y=580)
        
        q6LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q6 :",font=("Arial",15))
        q6LCA3.place(x=1150,y=380)

        q6TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q6TCA3.configure(state="disabled", fg_color="gray")
        q6TCA3.place(x=1200,y=380)
        
        q7LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q7 :",font=("Arial",15))
        q7LCA3.place(x=1150,y=430)
    
        q7TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q7TCA3.configure(state="disabled", fg_color="gray")
        q7TCA3.place(x=1200,y=430)
        
        q8LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q8 :",font=("Arial",15))
        q8LCA3.place(x=1150,y=480)
    
        q8TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q8TCA3.configure(state="disabled", fg_color="gray")
        q8TCA3.place(x=1200,y=480)
        
        q9LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q9 :",font=("Arial",15))
        q9LCA3.place(x=1150,y=530)
    
        q9TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q9TCA3.configure(state="disabled", fg_color="gray")
        q9TCA3.place(x=1200,y=530)
        
        q10LCA3=ctk.CTkLabel(master=tabview.tab(" CO Mapping "),text="Q10 :",font=("Arial",15))
        q10LCA3.place(x=1150,y=580)
    
        q10TCA3=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
        q10TCA3.configure(state="disabled", fg_color="gray")
        q10TCA3.place(x=1200,y=580)
        
        button = ctk.CTkButton(master=tabview.tab(" CO Mapping "),text=" Download ",font=("Arial",20),command=download)
        button.place(x=1000,y=640)
        
        def upload_file():
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            path_entry.delete(0, ctk.END)  # Clear any existing text in the entry widget
            path_entry.insert(0, file_path)  # Insert the file path into the entry widget

         
        def process_file():
            file_path = path_entry.get()
            import Cal
            Cal.cal_sheet(file_path)
            

        button_upload=ctk.CTkButton(tabview.tab(" Upload Excel File "),text="Upload",width=100,height=30,command=upload_file)
        button_upload.place(x=100,y=100)
        
        path_entry=ctk.CTkEntry(tabview.tab(" Upload Excel File "))
        
        button_process=ctk.CTkButton(tabview.tab(" Upload Excel File "),text="Process",width=100,height=30,command=process_file)
        button_process.place(x=500,y=100)
        self.app.mainloop()
        
if __name__ == "__main__":
    User_mode()