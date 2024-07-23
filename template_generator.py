import openpyxl
from openpyxl.styles import Alignment,Font
from openpyxl.styles.borders import Border,Side
from openpyxl import Workbook
def template_gen(array,array1):
    workbook=Workbook() 

    sheet1 = workbook.active
    sheet1.title = "Midsem"

    sheet2=workbook.create_sheet(title="Endsem")
    sheet3=workbook.create_sheet(title="CA")
    sheet4=workbook.create_sheet(title="Quiz")
    sheet5=workbook.create_sheet(title="Survey")
    sheet6=workbook.create_sheet(title='Attainment')
    
    sheet1.column_dimensions['A'].width = 42
    sheet1.merge_cells("A1:O1")
    sheet1["A1"].value="Vivekanand Education Society's Institute of Technology"

    
    sheet1.merge_cells("A2:O2")
    sheet1["A2"].value="Department of "+array[16]+""

    sheet1.merge_cells("A3:O3")
    sheet1["A3"].value="Academic Year :"+array[3]+""
    
    sheet1.merge_cells("A4:O4")
    
    sheet1.merge_cells("A5:O5")
    sheet1["A5"].value="  Subject : "+array[1]+"                                                                                                                                                                       Class : "+array[15]+""


    sheet1.merge_cells("A6:O6")
    sheet1["A6"].value="  Subject Teacher :"+array[4]+"                                                                                                                                                                Semester : "+array[2]+""
  
    

    sheet1.merge_cells("A7:O7")
    sheet1["A7"].value="Number of Students ="+array[0]+""

    sheet1["A8"]="Roll Nos."
    sheet1["B8"]="1a"
    sheet1["C8"]="1b"
    sheet1["D8"]="1c"
    sheet1["E8"]="1d"
    sheet1["F8"]="1e"
    sheet1["G8"]="1f"
    sheet1["H8"]="Q1"
    sheet1["I8"]="2a"
    sheet1["J8"]="2b"
    sheet1["K8"]="Q2"
    sheet1["L8"]="3a"
    sheet1["M8"]="3b"
    sheet1["N8"]="Q3"
    sheet1["O8"]="Total"

    sheet1["A9"]="COs"
    sheet1["B9"]="CO"+array[5]+""
    sheet1["C9"]="CO"+array[6]+""
    sheet1["D9"]="CO"+array[7]+""
    sheet1["E9"]="CO"+array[8]+""
    sheet1["F9"]="CO"+array[9]+""
    sheet1["G9"]="CO"+array[10]+""
    # sheet1["H9"]=""
    sheet1["I9"]="CO"+array[11]+""
    sheet1["J9"]="CO"+array[12]+""
    # sheet1["K9"]=""
    sheet1["L9"]="CO"+array[13]+""
    sheet1["M9"]="CO"+array[14]+""
    # sheet1["N9"]=""
    sheet1["O9"]=20

    total_roll=int(array[0])
    for i in range(1,total_roll+1):
        sheet1[f'A{i+9}']=i
        sheet1[f'A{i+9}'].alignment= Alignment(horizontal='center', vertical='center')
        for col in ['B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L','M','N','O']:
            sheet1[f'{col}{i+9}'].alignment= Alignment(horizontal='center', vertical='center')
        
    sheet1['A1'].alignment= Alignment(horizontal='center', vertical='center')
    sheet1['A2'].alignment= Alignment(horizontal='center', vertical='center')
    sheet1['A3'].alignment= Alignment(horizontal='center', vertical='center')
    sheet1['A4'].alignment= Alignment(horizontal='left', vertical='center')
    sheet1['A5'].alignment= Alignment(horizontal='left', vertical='center')
    sheet1['A6'].alignment= Alignment(horizontal='left', vertical='center')
    sheet1['A7'].alignment= Alignment(horizontal='left', vertical='center')

    
    
    column_range = ['A','B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L','M','N','O']
    for i in range(7,10):
        for col in column_range :
            sheet1[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')
    
    for i in range(total_roll+10,total_roll+17):
        for col in ['B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L','M','N','O'] :
            sheet1[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')

    for i in range(1,total_roll+17):
        for col in column_range:
            sheet1[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
    
    for col in column_range:
        sheet1[f'{col}{total_roll+16}'].font=Font(bold=True)
          
    for i in range(1,10):
        for col in column_range:
          sheet1[f'{col}{i}'].font=Font(bold=True)
    
    sheet1[f'A{total_roll+10}'] = 'Count(Attempted)'
    sheet1[f'A{total_roll+11}'] = 'Average Marks'
    sheet1[f'A{total_roll+12}'] = f'Count(>={array[19]}%)'
    sheet1[f'A{total_roll+13}'] = f'% Count(>={array[19]}% w.r.t appeared)'
    sheet1[f'A{total_roll+14}'] = 'Count(>=Average Marks of class)'
    sheet1[f'A{total_roll+15}']= "% Count(>=Average Marks of class w.r.t appeared)"
    sheet1[f'A{total_roll+16}'] = f'AL(Based on >={array[19]}% Count) (All COs)'
    
    sheet1[f'F{total_roll+19}'] = "COs"
    sheet1[f'F{total_roll+19}'].font=Font(bold=True)
    sheet1[f'G{total_roll+19}'] = "AL"
    sheet1[f'G{total_roll+19}'].font=Font(bold=True)
    sheet1[f'F{total_roll+20}'] = 'CO1'
    sheet1[f'F{total_roll+21}'] = 'CO2'
    sheet1[f'F{total_roll+22}'] = 'CO3'
    sheet1[f'F{total_roll+23}'] = 'CO4'
    sheet1[f'F{total_roll+24}'] = 'CO5'
    sheet1[f'F{total_roll+25}'] = 'CO6'

    for i in range(total_roll+19,total_roll+26):
        sheet1[f'F{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet1[f'G{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet1[f'F{i}'].alignment= Alignment(horizontal='center', vertical='center')
        sheet1[f'G{i}'].alignment= Alignment(horizontal='center', vertical='center')
        
    #<-----------End Semester Template------------->
    
    sheet2.column_dimensions['A'].width =42
    sheet2.column_dimensions['B'].width =22
    sheet2['A1']="Roll No."  
    sheet2['A1'].font=Font(bold=True) 
    sheet2['B1']="ESE(TH)"
    sheet2['B1'].font=Font(bold=True)     
    sheet2['B2']="ALL Qs"
    sheet2['B2'].font=Font(bold=True) 
    sheet2['B3']="CO"+array[18]+""
    sheet2['B3'].font=Font(bold=True) 
    
    for i in range(1,total_roll+1):
        sheet2[f'A{i+3}'] =i
        # sheet2[f'A{i+3}'].alignment= Alignment(horizontal='center', vertical='center')
        # sheet2[f'B{i+3}'].alignment= Alignment(horizontal='center', vertical='center')     
        # sheet2[f'A{i+3}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        # sheet2[f'B{i+3}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))

    
    sheet2[f'A{total_roll+4}']="Count(Attempted)"
    sheet2[f'A{total_roll+5}']="Average Marks"
    
   
    sheet2[f'A{total_roll+6}']=f"Count(>={array[19]}%)"
    
    sheet2[f'A{total_roll+7}']=f"% Count(>={array[19]}% w.r.t appeared)"
    
    sheet2[f'A{total_roll+8}']="Count(>=Average Marks of class)"
    sheet2[f'A{total_roll+9}']="% Count(>=Average Marks of class w.r.t appeared)"
    
    sheet2[f'A{total_roll+10}']=f"AL(Based on >={array[19]}% Count) (All COs)"
    sheet2[f'A{total_roll+10}'].font=Font(bold=True) 
    sheet2[f'B{total_roll+10}'].font=Font(bold=True)
    
    for i in range(1,total_roll+11):
        sheet2[f'A{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet2[f'A{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet2[f'B{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet2[f'B{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        if i>=total_roll+4 :
            sheet2[f'A{i}'].alignment= Alignment(horizontal='left', vertical='center') 
     
    sheet2[f'A{total_roll+13}'] = "COs"
    sheet2[f'A{total_roll+13}'].font=Font(bold=True)
    sheet2[f'B{total_roll+13}'] = "AL"
    sheet2[f'B{total_roll+13}'].font=Font(bold=True)
    sheet2[f'A{total_roll+14}'] = 'CO1'
    sheet2[f'A{total_roll+15}'] = 'CO2'
    sheet2[f'A{total_roll+16}'] = 'CO3'
    sheet2[f'A{total_roll+17}'] = 'CO4'
    sheet2[f'A{total_roll+18}'] = 'CO5'
    sheet2[f'A{total_roll+19}'] = 'CO6' 
    
    for i in range(total_roll+13,total_roll+20):
        sheet2[f'A{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet2[f'B{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet2[f'A{i}'].alignment= Alignment(horizontal='center', vertical='center')
        sheet2[f'B{i}'].alignment= Alignment(horizontal='center', vertical='center')
        
    #<-----------CA Template------------->  
    
    sheet3.column_dimensions['A'].width = 42
    sheet3.merge_cells("A1:D1")
    sheet3.merge_cells("A2:D2")
    sheet3.merge_cells("A3:D3")
    sheet3['A2']="CA Marksheet"
    sheet3['A2'].font=Font(bold=True)
    sheet3['A4']="Roll No."
    sheet3['A4'].font=Font(bold=True)
    sheet3['B4']="CA1"
    sheet3['B4'].font=Font(bold=True)
    sheet3['C4']="CA2"
    sheet3['C4'].font=Font(bold=True)
    sheet3['D4']="CA3"
    sheet3['D4'].font=Font(bold=True)
    sheet3['B5']="CO"+array1[0]+""
    sheet3['B5'].font=Font(bold=True)
    sheet3['C5']="CO"+array1[1]+""
    sheet3['C5'].font=Font(bold=True)
    sheet3['D5']="CO"+array1[2]+""
    sheet3['D5'].font=Font(bold=True)
    
    for i in range(1 ,total_roll+1):
        sheet3[f'A{i+5}']=i
     
    sheet3[f'A{total_roll+6}']="Count(Attempted)"
    sheet3[f'A{total_roll+7}']="Average Marks"
    
   
    sheet3[f'A{total_roll+8}']=f"Count(>={array[19]}%)"
    
    
    sheet3[f'A{total_roll+9}']=f"% Count(>={array[19]}% w.r.t appeared)"
    
    sheet3[f'A{total_roll+10}']="Count(>=Average Marks of class)"
    sheet3[f'A{total_roll+11}']="% Count(>=Average Marks of class w.r.t appeared)"
    
    sheet3[f'A{total_roll+12}']=f"AL(Based on >={array[19]}% Count) (All COs)"
    sheet3[f'A{total_roll+12}'].font=Font(bold=True)
     
    for i in range(1,total_roll+13):
        sheet3[f'A{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet3[f'A{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet3[f'B{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet3[f'B{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet3[f'C{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet3[f'C{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet3[f'D{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet3[f'D{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        
        if i>=total_roll+6 :
            sheet3[f'A{i}'].alignment= Alignment(horizontal='left', vertical='center') 
     
    for i in range(total_roll+15,total_roll+22):
        sheet3[f'B{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet3[f'B{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet3[f'C{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet3[f'C{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        
    sheet3[f'B{total_roll+15}'] = "COs"
    sheet3[f'B{total_roll+15}'].font=Font(bold=True)
    sheet3[f'C{total_roll+15}'] = "AL"
    sheet3[f'C{total_roll+15}'].font=Font(bold=True)
    sheet3[f'B{total_roll+16}'] = 'CO1'
    sheet3[f'B{total_roll+17}'] = 'CO2'
    sheet3[f'B{total_roll+18}'] = 'CO3'
    sheet3[f'B{total_roll+19}'] = 'CO4'
    sheet3[f'B{total_roll+20}'] = 'CO5'
    sheet3[f'B{total_roll+21}'] = 'CO6' 
    
    #<-----------Quiz Template------------->  
    
    sheet4.column_dimensions['B'].width =42
    sheet4['A1']="Roll No."
    sheet4['B1']="Name"
    sheet4['C1']="Q1"
    sheet4['D1']="Q2"
    sheet4['E1']="Q3"
    sheet4['F1']="Q4"
    sheet4['G1']="Q5"
    sheet4['H1']="Q6"
    sheet4['I1']="Q7"
    sheet4['J1']="Q8"
    sheet4['K1']="Q9"
    sheet4['L1']="Q10"
    
    for i in range(1,3):
        for col in  ['A','B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L'] :
            sheet4[f'{col}{i}'].font=Font(bold=True)
            
    sheet4.merge_cells('A2:B2')
    sheet4['C2']="CO"+array1[3]+""
    sheet4['D2']="CO"+array1[4]+""
    sheet4['E2']="CO"+array1[5]+""
    sheet4['F2']="CO"+array1[6]+""
    sheet4['G2']="CO"+array1[7]+""
    sheet4['H2']="CO"+array1[8]+""
    sheet4['I2']="CO"+array1[9]+""
    sheet4['J2']="CO"+array1[10]+""
    sheet4['K2']="CO"+array1[11]+""
    sheet4['L2']="CO"+array1[12]+"" 
    
    for i in range(1 ,total_roll+1):
        sheet4[f'A{i+2}']=i
     
    for i in range(total_roll+3,total_roll+11) :
        sheet4.merge_cells(f'A{i}:B{i}')
        
    for i in range(1,total_roll+10):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L'] :
            sheet4[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            sheet4[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
            if i>total_roll+2 :
                sheet4[f'{col}{i}'].alignment= Alignment(horizontal='left', vertical='center')     
            
    for i in range(total_roll+3,total_roll+10):
        for col in ['C', 'D', 'E', 'F', 'G','H','I', 'J','K','L'] :
            sheet4[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            sheet4[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        if i==total_roll+9:
            sheet4[f'{col}{i}'].font=Font(bold=True)        
               
   
    sheet4[f'A{total_roll+3}']="Count(Attempted)"       
    sheet4[f'A{total_roll+4}']="Average Marks"
    
   
    sheet4[f'A{total_roll+5}']=f"Count(>={array[19]}%)"
    
    
    sheet4[f'A{total_roll+6}']=f"% Count(>={array[19]}% w.r.t appeared)"

    sheet4[f'A{total_roll+7}']="Count(>=Average Marks of class)"
    sheet4[f'A{total_roll+8}']="% Count(>=Average Marks of class w.r.t appeared)"
    
    sheet4[f'A{total_roll+9}']=f"AL(Based on >={array[19]}% Count) (All COs)"
    sheet4[f'A{total_roll+9}'].font=Font(bold=True)
    
    for i in range(total_roll+12,total_roll+19):
        sheet4[f'C{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet4[f'C{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        sheet4[f'D{i}'].alignment= Alignment(horizontal='center', vertical='center')     
        sheet4[f'D{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        
    sheet4[f'C{total_roll+12}'] = "COs"
    sheet4[f'C{total_roll+12}'].font=Font(bold=True)
    sheet4[f'D{total_roll+12}'] = "AL"
    sheet4[f'D{total_roll+12}'].font=Font(bold=True)
    sheet4[f'C{total_roll+13}'] = 'CO1'
    sheet4[f'C{total_roll+14}'] = 'CO2'
    sheet4[f'C{total_roll+15}'] = 'CO3'
    sheet4[f'C{total_roll+16}'] = 'CO4'
    sheet4[f'C{total_roll+17}'] = 'CO5'
    sheet4[f'C{total_roll+18}'] = 'CO6' 
    
    #<-----------Survey Template------------->
    sheet5.column_dimensions['B'].width =40
    sheet5.column_dimensions['C'].width =40
    sheet5.column_dimensions['F'].width =40
    
    sheet5['A1']="Sr. No."
    sheet5['B1']="Email Address"
    sheet5['C1']="Full name of Student"
    sheet5['D1']="Roll No."
    sheet5['E1']="Class"
    sheet5['F1']="Branch"
    sheet5['G1']="Q1"
    sheet5['H1']="Q2"
    sheet5['I1']="Q3"
    sheet5['J1']="Q4"
    sheet5['K1']="Q5"
    sheet5['L1']="Q6"
    
    for col in  ['A','B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L'] :
            sheet5[f'{col}1'].font=Font(bold=True)
            
    for i in range(1 ,total_roll+1):
        sheet5[f'A{i+1}']=i
        sheet5[f'E{i+1}']=""+array[15]+""
        sheet5[f'F{i+1}']=""+array[16]+""
        
    for i in range(1,total_roll+2):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L'] :
            sheet5[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            sheet5[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
    
    sheet5[f'F{total_roll+4}']= 'Total' 
    sheet5[f'F{total_roll+4}'].font=Font(bold=True) 
    sheet5[f'F{total_roll+5}']= 'SA + A Count'
    sheet5[f'F{total_roll+5}'].font=Font(bold=True)
    sheet5[f'F{total_roll+6}']= 'SA + A Percentage' 
    sheet5[f'F{total_roll+6}'].font=Font(bold=True)
    sheet5[f'F{total_roll+7}']= 'CO Mapped' 
    sheet5[f'F{total_roll+7}'].font=Font(bold=True)
    sheet5[f'F{total_roll+8}']= 'AL'
    
    for i in range(total_roll+4,total_roll+9):
        for col in ['F','G','H','I','J','K','L'] :
            sheet5[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            sheet5[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
            if col=="F":
                sheet5[f'F{i}'].alignment= Alignment(horizontal='left', vertical='center')     
    
            if i==total_roll+8:
                sheet5[f'{col}{i}'].font=Font(bold=True)
        
    sheet5[f'G{total_roll+7}']= 'CO1' 
    sheet5[f'H{total_roll+7}']= 'CO2' 
    sheet5[f'I{total_roll+7}']= 'CO3' 
    sheet5[f'J{total_roll+7}']= 'CO4' 
    sheet5[f'K{total_roll+7}']= 'CO5' 
    sheet5[f'L{total_roll+7}']= 'CO6'
    
    
    
    #<-----------------------Attainment--------------------->
    
    sheet6.column_dimensions['A'].width =16
    sheet6.column_dimensions['B'].width =25
    sheet6.column_dimensions['C'].width =25
    sheet6.column_dimensions['D'].width =25
    sheet6.column_dimensions['E'].width =25
    sheet6.column_dimensions['F'].width =34
    sheet6.column_dimensions['G'].width =25
    
    for i in range (1,9):
        sheet6.merge_cells(f"A{i}:H{i}")
        sheet6[f'A{i}'].font=Font(bold=True)
        
    
    for i in range (9,15):
        sheet6.merge_cells(f"B{i}:H{i}") 
         
    

    sheet6["A1"].value="Vivekanand Education Society's Institute of Technology"
    sheet6["A1"].alignment= Alignment(horizontal='center', vertical='center')     
    
    sheet6["A2"].value="Department of "+array[16]+""
    sheet6["A2"].alignment= Alignment(horizontal='center', vertical='center')     
    
    sheet6["A3"].value="Academic Year :"+array[3]+""
    sheet6["A3"].alignment= Alignment(horizontal='center', vertical='center')     
    
    sheet6["A5"].value="  Subject : "+array[1]+"                                                                                                                                                                       Class : "+array[15]+""
    sheet6["A5"].alignment= Alignment(horizontal='left', vertical='center')     
    
    sheet6["A6"].value="  Subject Teacher :"+array[4]+"                                                                                                                                                                Semester : "+array[2]+""
    sheet6["A6"].alignment= Alignment(horizontal='left', vertical='center')     
    
    
    sheet6['A8']='Course Outcomes(COs): Upon successful completion of this course, students will be able to:'
    sheet6['A8'].font=Font(bold=True)
    sheet6["A8"].alignment= Alignment(horizontal='left', vertical='center')     
    
    sheet6['A9'] ='CO1'
    sheet6['A10']='CO2'
    sheet6['A11']='CO3'
    sheet6['A12']='CO4'
    sheet6['A13']='CO5'
    sheet6['A14']='CO6'
    
    for i in range (9,15):
        sheet6[f'A{i}'].alignment= Alignment(horizontal='center', vertical='center')
        sheet6[f'B{i}'].alignment= Alignment(horizontal='left', vertical='center')         
    
    for i in range(9,15):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H']:
            sheet6[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))  
                    
    sheet6.merge_cells("A15:H15")
    sheet6.merge_cells("A16:H16")
    
    sheet6['A16']='CO Rubrics Mapping'
    sheet6['A16'].alignment= Alignment(horizontal='center', vertical='center')     
    
    sheet6.merge_cells("A17:H17")
    
    sheet6.merge_cells("A18:A19")  
    sheet6.merge_cells("F18:F19")
    
    sheet6.merge_cells("B18:E18")
    sheet6.merge_cells("B19:D19") 
    
    for i in range(16,21):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H']:
            sheet6[f'{col}{i}'].font=Font(bold=True)
    
    for i in range(18,27):
        for col in ['A','B', 'C', 'D', 'E', 'F']:
            sheet6[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000')) 
            sheet6[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
                
    sheet6['A18']='Assessment'
    sheet6['B18']='Direct Assessment' 
    sheet6['F18']='Indirect Assessment' 
    
    sheet6['B19']='Internal Assessment' 
    sheet6['E19']='External Assessment' 
    
    sheet6['A20']="CO's"
    sheet6['B20']="Mid Term Test"
    sheet6['C20']="Countinous Assessment"
    sheet6['D20']="Quiz"
    sheet6['E20']="ESE(TH)"
    sheet6['F20']="Course Exit Survey"
    
    sheet6['A21']='CO1'
    sheet6['A22']='CO2'
    sheet6['A23']='CO3'
    sheet6['A24']='CO4'
    sheet6['A25']='CO5'
    sheet6['A26']='CO6'
    
    sheet6.merge_cells("A27:H27")
    sheet6.merge_cells("A28:H28")
    sheet6.merge_cells("A29:H29")
    
    sheet6['A28']='CO Attainment (Level)'
    sheet6['A28'].alignment= Alignment(horizontal='center', vertical='center')     
    
    sheet6.merge_cells("A30:A31")
    sheet6.merge_cells("G30:G31")
    
    sheet6.merge_cells("B30:F30")
    sheet6.merge_cells("B31:D31")
    
    for i in range(28,33):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H']:
            sheet6[f'{col}{i}'].font=Font(bold=True)
    
    for i in range(30,39):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G']:
            sheet6[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000')) 
            sheet6[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
       
    sheet6['A30']='Assessment'
    sheet6['B30']='Direct Assessment'
    sheet6['G30']='Indirect Assessment'
    
    sheet6['B31']='Internal Assessment'
    sheet6['E31']='External Assessment'
    sheet6['F31']='Attainment Level'
    
    sheet6['A32']="CO's"
    sheet6['B32']="Mid Term Test"
    sheet6['C32']="Countinous Assessment"
    sheet6['D32']="Quiz"
    sheet6['E32']="ESE(TH)"
    sheet6['F32']="70% (External) + 30% (Internal)"
    sheet6['G32']="Course Exit Survey"
    
    sheet6['A33']='CO1'
    sheet6['A34']='CO2'
    sheet6['A35']='CO3'
    sheet6['A36']='CO4'
    sheet6['A37']='CO5'
    sheet6['A38']='CO6'
    
    sheet6.merge_cells("A39:H39")
    sheet6.merge_cells("A40:H40")
    sheet6.merge_cells("A41:H41")
    
    for i in range(42,49):
        for col in ['C', 'D', 'E']:
            sheet6[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000')) 
            sheet6[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
     
    sheet6['A40']='Final CO Attainment'
    sheet6['A40'].font=Font(bold=True)
    sheet6['A40'].alignment= Alignment(horizontal='center', vertical='center')     
    
    sheet6['C42']='Course Outcomes'
    sheet6['C42'].font=Font(bold=True)
    
    sheet6['D42']='Direct AL'
    sheet6['D42'].font=Font(bold=True)
    
    sheet6['E42']='Indirect AL'
    sheet6['E42'].font=Font(bold=True)
    
    sheet6['C43']='CO1'
    sheet6['C43'].font=Font(bold=True)
    
    sheet6['C44']='CO2'
    sheet6['C44'].font=Font(bold=True)
    
    sheet6['C45']='CO3'
    sheet6['C45'].font=Font(bold=True)
    
    sheet6['C46']='CO4'
    sheet6['C46'].font=Font(bold=True)
    
    sheet6['C47']='CO5'
    sheet6['C47'].font=Font(bold=True)
    
    sheet6['C48']='CO6'
    sheet6['C48'].font=Font(bold=True)
    
    
    workbook.save('C:/Users/saira/Downloads/Template.xlsx')
