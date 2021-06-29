from tkinter import *
import tooltip
root = Tk()
 
btn = Button(root, text='A Button')
btn.grid(row=1,column=1)
 
tooltip.Create(btn, "Your tooltip text")
 
root.mainloop()