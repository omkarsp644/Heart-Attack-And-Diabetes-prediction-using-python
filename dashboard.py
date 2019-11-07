from tkinter import *
from playsound import playsound
import os
from datetime import datetime
from tkinter import messagebox
root=Tk()









root.configure(background="white")
def function1():
    
    os.system("py heartreal.py")
    
def function2():
    
    os.system("py mini.py")
   
def function6():

    root.destroy()


#stting title for the window
root.title("Heart And Diabetise Prediction Algorithms")

#creating a text label
Label(root, text="DISEASE PREDICTION",font=("times new roman",20),fg="white",bg="maroon",height=2,width=50).grid(row=0,rowspan=2,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)

#creating first button
Button(root,text="Heart Disease Prediction",font=("times new roman",20),bg="#0D47A1",fg='white',command=function1).grid(row=3,columnspan=2,sticky=W+E+N+S,padx=5,pady=5)

#creating second button
Button(root,text="Diabetise Prediction",font=("times new roman",20),bg="#0D47A1",fg='white',command=function2).grid(row=4,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)


Button(root,text="Exit",font=('times new roman',20),bg="maroon",fg="white",command=function6).grid(row=9,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)

root.mainloop()

