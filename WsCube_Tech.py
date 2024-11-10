

from tkinter import*

from streamlit import button

# bg="black" ,bg="lightgray"


#______________________________________________________
win= Tk()  # start
win.title("WsCube Tech E Shool")
win.config(bg="lightgray")
win.geometry("600x700")
win.resizable(False,False)


#-----------------------------------------------------------------

# tite name
school_name =Label(win,text="WsCube Tech E Shool",font=("Times New Roman",30,"bold"),bg="lightgray")
school_name.place(x=100,y=20,height=60,width=400)

#-----------------------------------------------------------------
# Name Entry
st_name =Label(win,text="Student Name",font=("Times New Roman",20,"bold"),)
st_name.place(x=10,y=100,height=50,width=200)

st_name_Entry = Entry(win,text="Subject Number",font=("Times New Roman",20,"bold"))
st_name_Entry.place(x=230,y=100,height=50,width=300)


#-----------------------------------------------------------------
# tite name
subject_name =Label(win,text="Subject Number",font=("Times New Roman",30,"bold"))
subject_name.place(x=130,y=170,height=50,width=300)

#-----------------------------------------------------------------
# subject NO 0
l = ["Hindi",'English','Sciene','Maths','Sst']


hindi_name =Label(win,text="Hindi",font=("Times New Roman",20,"bold"),)
hindi_name.place(x=10,y=240,height=50,width=200)

hindi_name_Entry = Entry(win,font=("Times New Roman",20,"bold"))
hindi_name_Entry.place(x=230,y=240,height=50,width=300)

#-----------------------------------------------------------------
# subject NO 1
l = ["Hindi",'English','Sciene','Maths','Sst']


English_name =Label(win,text="English",font=("Times New Roman",20,"bold"),)
English_name.place(x=10,y=300,height=50,width=200)

English_name_Entry = Entry(win,font=("Times New Roman",20,"bold"))
English_name_Entry.place(x=230,y=300,height=50,width=300)


#-----------------------------------------------------------------
# subject NO 2

Sciene_name =Label(win,text="Sciene",font=("Times New Roman",20,"bold"),)
Sciene_name.place(x=10,y=360,height=50,width=200)

Sciene_name_Entry = Entry(win,font=("Times New Roman",20,"bold"))
Sciene_name_Entry.place(x=230,y=360,height=50,width=300)

#-----------------------------------------------------------------
# subject NO 3

Maths_name =Label(win,text="Maths",font=("Times New Roman",20,"bold"),)
Maths_name.place(x=10,y=420,height=50,width=200)

Maths_name_Entry = Entry(win,font=("Times New Roman",20,"bold"))
Maths_name_Entry.place(x=230,y=420,height=50,width=300)

#-----------------------------------------------------------------
# subject NO 4


Sst_name =Label(win,text="Sst",font=("Times New Roman",20,"bold"),)
Sst_name.place(x=10,y=480,height=50,width=200)

Sst_name_Entry = Entry(win,font=("Times New Roman",20,"bold"))
Sst_name_Entry.place(x=230,y=480,height=50,width=300)

#-----------------------------------------------------------------
# Button

button = Button(win,text="Done",font=("Times New Roman",20,"bold"))
button.place(x=230,y=480,height=50,width=300)




win.mainloop()