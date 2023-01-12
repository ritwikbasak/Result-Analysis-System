from tkinter import filedialog
from functools import *
import tkinter as tk
import tkinter.messagebox
import pandas as pd
import xlrd
import random
import matplotlib.pyplot as plt
import math
from PIL import ImageTk,Image
file_windows=[]
def get_subjects(data_file,year,stream):
    column_heads=list(data_file.head(0))
    year_list=list(data_file[column_heads[2]])
    stream_list=list(data_file[column_heads[3]])
    subject_wise_grade={}
    for i in range(4,len(column_heads)):
        grade_list=list(data_file[column_heads[i]])
        dictionary_key=column_heads[i]
        if "." in dictionary_key:
            dictionary_key=dictionary_key[:dictionary_key.index(".")]
        for j in range(len(grade_list)):
            if str(year_list[j])!=year[0] or stream_list[j]!=stream:
                continue
            current_grade=str(grade_list[j])
            if current_grade.lower()=="nan":
                continue
            if not(dictionary_key in subject_wise_grade.keys()):
                subject_wise_grade[dictionary_key]=[current_grade[0]]
            else:
                subject_wise_grade[dictionary_key].append(current_grade[0])
    return subject_wise_grade
def showSubjects(data_file,year,stream,file_window):
    subjects=get_subjects(data_file,year,stream)
    subject_labels=sorted(subjects.keys())
    max_length=len(max(subject_labels,key=len))+2
    choice_list=[tk.IntVar() for i in subject_labels]
    def enable_analysis(current_checkbutton_pressed):
        if current_checkbutton_pressed.get()==0:
            current_checkbutton_pressed.set(1)
        else:
            current_checkbutton_pressed.set(0)
        for i in choice_list:
            if i.get()==1:
                analysis_btn.config(state=tk.NORMAL)
                return
        analysis_btn.config(state=tk.DISABLED)
    chkbtn_list=[tk.Checkbutton(file_window,text="",command=partial(enable_analysis,var)) for var in choice_list]
    x_place=50
    y_place=90
    def analysis():
        tk.messagebox.showinfo("Graph","Please Close One Graph Before Plotting Another.\nOtherwise They Might Be Superimposed.",parent=file_window)
        subject_labels_selected=[subject_labels[i] for i in range(len(choice_list)) if choice_list[i].get()==1]
        subjects_selected={}
        for i in subject_labels_selected:
            subjects_selected[i]=subjects[i][:]
        x_axis_labels=["O","E","A","B","C","D","F","I"]
        y_values=[[subjects_selected[name].count(i) for i in x_axis_labels] for name in subject_labels_selected]
        subject_colours=[]
        count=1
        while(count<=len(subject_labels_selected)):
            generated_colour=(random.randint(0,100)/100,random.randint(0,100)/100,random.randint(0,100)/100,random.randint(0,100)/100)
            if (generated_colour in subject_colours) or generated_colour[3]<0.5:
                continue
            subject_colours.append(generated_colour)
            count+=1
        barWidth=1/(2+len(subject_labels_selected))
        for i in range(len(subject_labels_selected)):
            plt.bar([j+barWidth*i for j in range(len(x_axis_labels))],y_values[i],color=subject_colours[i],width=barWidth,edgecolor="white",label=subject_labels_selected[i])
        plt.xticks([i+(len(subject_labels_selected)//2)*barWidth for i in range(len(x_axis_labels))],x_axis_labels)
        plt.xlabel("Grade",fontweight="bold")
        plt.ylabel("No. Of Students",fontweight="bold")
        plt.legend()
        plt.suptitle("Result Analysis Of "+year+","+stream,fontsize=17)
        plt.show()
    analysis_btn=tk.Button(file_window,text="Analysis",padx=20,command=analysis,state=tk.DISABLED)
    analysis_btn.place(x=300,y=y_place)
    for i in range(len(subject_labels)):
        label=tk.Label(file_window,text=str(i+1)+"."+subject_labels[i],font=("times new roman",14,"normal"))
        label.place(x=x_place-int(math.log(i+1,10))*10,y=y_place)
        chkbtn_list[i].place(x=x_place+max_length*7+80,y=y_place)
        y_place+=30
def upload_file(year,stream):
    global file_windows
    if file_windows!=[]:
        tk.messagebox.showerror("","Two Upload Windows Cannot Be Simultaneously Opened",parent=first_window)
        return
    file_windows.append((year,stream))
    flag=False
    data_file=None
    file_window=tk.Tk()
    file_window.title("")
    file_window.resizable(0,0)
    file_window.geometry("800x600")
    main_label=tk.Label(file_window,text="Result Analysis For "+year+","+stream,font=("times new roman",20,"bold"))
    main_label.place(x=2,y=2)
    def get_excel():
        global data_file
        import_file_path = filedialog.askopenfilename(parent=file_window)
        try:
            data_file = pd.read_excel(import_file_path)
        except FileNotFoundError:
            tk.messagebox.showerror("Upload Error","No File Selected",parent=file_window)
        except xlrd.biffh.XLRDError:
            tk.messagebox.showerror("Upload Error","Wrong File Type Selected",parent=file_window)
        else:
            file_path_label=tk.Label(file_window,text=import_file_path[:import_file_path.rindex("/" if "/" in import_file_path else "\\")+1],font=("times new roman",10,"normal"))
            file_path_label.place(x=300,y=50)
            file_name_label=tk.Label(file_window,text=import_file_path[import_file_path.rindex("/" if "/" in import_file_path else "\\")+1:],font=("times new roman",10,"bold"))
            file_name_label.place(x=300,y=70)
            browse_button_Excel.config(state=tk.DISABLED)
            try:
                showSubjects(data_file,year,stream,file_window)
            except (ValueError,TypeError,LookupError,EOFError,AttributeError,AssertionError,ArithmeticError):
                tk.messagebox.showerror("File Error","Corrupted File",parent=file_window)
                browse_button_Excel.config(state=tk.NORMAL)
                file_path_label.destroy()
                file_name_label.destroy()
    browse_button_Excel=tk.Button(file_window,text="Upload Excel File",padx=50,command=get_excel)
    browse_button_Excel.place(x=50,y=50)
    def callback():
        file_window.destroy()
        file_windows.remove((year,stream))
    file_window.protocol("WM_DELETE_WINDOW",callback)
    file_window.mainloop()
first_window=tk.Tk()
first_window.resizable(0,0)
first_window.geometry("500x450")
first_frame=tk.Frame(first_window)
first_frame.pack(fill=tk.X,expand=False,side=tk.TOP)
first_window.title("Result Analysis")
first_mb=[]
first_menuContent=[]
text_list=["1st Year","2nd Year","3rd Year","4th Year"]
stream_list=["CSE","IT","ME","EE","EEE","ECE"]
for i in range(4):
    first_mb.append(tk.Menubutton(first_frame,text=text_list[i],relief=tk.RAISED,justify=tk.CENTER))
    first_menuContent.append(tk.Menu(first_mb[i],tearoff=0))
    first_mb[i].config(menu=first_menuContent[i],height=1,width=8)
    first_mb[i].pack(side=tk.LEFT)
    for txt in stream_list:
        first_menuContent[i].add_command(label=txt,command=partial(upload_file,text_list[i],txt))
def about_callback():
    about_window=tk.Tk()
    about_window.title("About")
    about_window.resizable(0,0)
    about_window.geometry("500x500")
    about_label=tk.Label(about_window,text="About",font=("times new roman",16,"bold"))
    about_label.place(x=10,y=10)
    about_file=open("about.txt","r")
    about_contents_label=tk.Label(about_window,text=about_file.read(),font=("times new roman",10,"normal"))
    about_contents_label.place(x=10,y=80)
    about_file.close()
    about_window.mainloop()
about_button=tk.Button(first_frame,text="About",relief=tk.RAISED,justify=tk.CENTER,padx=10,command=about_callback)
about_button.pack(side=tk.RIGHT)
image_file=ImageTk.PhotoImage(Image.open("resultanalysis.jpg"))
picture_label=tk.Label(first_window,image=image_file)
picture_label.pack(side=tk.TOP,expand=True)
first_window.mainloop()
