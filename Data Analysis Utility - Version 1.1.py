 
import math
import os
from tkinter import *
import tkinter.messagebox
from tkinter.filedialog import askopenfilename
from tkinter import filedialog
import collections
from pathlib import Path
import pandas as pd
from functools import reduce
import matplotlib.pyplot as plt
 

class MainWindow():

    def __init__(self, master):


        self.master = master
        self.a = StringVar()
        self.save_vs = StringVar()
       
        #Window Title
        self.master.title("Data Analysis Utility - Version 1.1")
        self.master.geometry('590x700+385+0')
        self.master.resizable(False, False)
        
        #Title
        self.heading = Label(self.master, text="Data Analysis Utility - Version 1.1", font="timesnewroman 20 bold", fg='green', bg='white')
        self.heading.grid(row=0, column=0, columnspan=4, sticky=N+E+W+S)
        
        #Read Frame
        self.generate_frame = LabelFrame(self.master, text="Read", font = "timesnewroman 10 bold", labelanchor=NE)
        self.generate_frame.grid(row=2, column=0, columnspan=4, sticky=N+E+W+S)
        
        #Filename 
        Label(self.generate_frame, text="Source File Path:", font="timesnewroman 12 bold") .grid(row=3, column=0, sticky=W)
        self.filename = Entry(self.generate_frame, width=78, bg="white", textvariable=self.a)
        self.filename.grid(row=4, column=0, columnspan = 1, sticky=W)
        self.browse1 = Button(self.generate_frame, text="Browse", font="timesnewroman 12 bold",  fg="black", width=10, height=1, border = 3, relief='raised',
                             command = lambda: self.get_source_filename())
        self.browse1.grid(row= 4, column=1, sticky=W)

        #Sheet 
        Label(self.generate_frame, text="Sheet Name:", font="timesnewroman 12 bold") .grid(row=7, column=0, sticky=W)
        self.sheet = Entry(self.generate_frame, width=45, bg="white", textvariable=StringVar(value='Sheet1'))
        self.sheet.grid(row=8, column=0, sticky=W)

        #Start Row
        Label(self.generate_frame, text="Start Row:", font="timesnewroman 12 bold") .grid(row=9, column=0, sticky=W)
        self.start = Entry(self.generate_frame, width=20, bg="white", textvariable=IntVar())
        self.start.grid(row=10, column=0, sticky=W)

        #Header Checkbutton
        self.include_header_var = tkinter.IntVar()
        self.include_header = Checkbutton(self.generate_frame,text="Include Header", font="timesnewroman 8 bold", variable = self.include_header_var,
                                          onvalue=True, offvalue=False)
        self.include_header.grid(row=12, column=0, sticky=W)

        #End Row
        Label(self.generate_frame, text="End Row:", font="timesnewroman 12 bold") .grid(row=17, column=0, sticky=W)
        self.stop = Entry(self.generate_frame, width=20, bg="white", textvariable=IntVar())
        self.stop.grid(row=18, column=0, sticky=W)

        #Read Button
        self.read_button = Button(self.generate_frame, text="READ", font="timesnewroman 12 bold",  fg="white", bg = "green", width=10, height=1, border = 3,
                                     relief='raised', command = lambda: self.main_fieldchecker())
        self.read_button.grid(row= 19, column=1, sticky=E)

        #Select Frame
        self.select_frame = LabelFrame(self.master, text="Filter", font = "timesnewroman 10 bold", labelanchor=NE)
        self.select_frame.grid(row=20, column=0, columnspan=4, sticky=N+E+W+S)

        #Columns Listbox
        Label(self.select_frame, text="Columns:", font="timesnewroman 12 bold") .grid(row=21, column=0, sticky=W)
        self.columns_listbox = Listbox(self.select_frame, height = 5, 
                  width = 32, 
                  bg = "white",
                  activestyle = 'dotbox', 
                  font = "timesnewroman",
                  fg = "black",
                  exportselection=False)
        self.columns_listbox.bind("<<ListboxSelect>>", self.callback_column)
        self.columns_listbox.grid(row=22, column=0, columnspan=2, sticky=W)

        #Reset Columns Button
        self.reset_columns = Button(self.select_frame, text="Reset", font="timesnewroman 8 bold",  fg="black",  width=6, height=1, border = 3,
                                     relief='raised', command = lambda: self.re_cols())
        self.reset_columns.grid(row= 21, column=1, sticky=W)

        #Items Listbox
        Label(self.select_frame, text="Items:", font="timesnewroman 12 bold") .grid(row=21, column=2, sticky=W)
        self.items_listbox = Listbox(self.select_frame, height = 5, 
                  width = 32, 
                  bg = "white",
                  activestyle = 'dotbox', 
                  font = "timesnewroman",
                  fg = "black",
                  selectmode='multiple',
                  exportselection=False)
        self.items_listbox.bind("<<ListboxSelect>>", self.callback_item)
        self.items_listbox.grid(row=22, column=2, columnspan=2, sticky=W)

        #Reset Items Button
        self.reset_items = Button(self.select_frame, text="Reset", font="timesnewroman 8 bold",  fg="black",  width=6, height=1, border = 3,
                                     relief='raised', command = lambda: self.re_items())
        self.reset_items.grid(row= 21, column=3, sticky=W)


        #Filter Button
        self.filter_button = Button(self.select_frame, text="FILTER", font="timesnewroman 12 bold",  fg="white", bg = "green", width=10, height=1, border = 3,
                                     relief='raised', command = lambda: self.save_filter())
        self.filter_button.grid(row= 33, column=3, sticky=E)

        #Visualize Frame
        self.vs_frame = LabelFrame(self.master, text="Visualize", font = "timesnewroman 10 bold", labelanchor=NE)
        self.vs_frame.grid(row=34, column=0, columnspan=2, sticky=N+E+W+S)

        #Column Name
        Label(self.vs_frame, text="Select Column:", font="timesnewroman 12 bold") .grid(row=35, column=0, sticky=W)
        self.column_name = StringVar()
        self.columns = OptionMenu(self.vs_frame, self.column_name, 'None')
        self.columns.grid(row = 36, column =0,  sticky=W, ipadx=20)
        self.column_name.set('None')
        
        #Visualization
        Label(self.vs_frame, text="Select Visualization:", font="timesnewroman 12 bold") .grid(row=37, column=0, sticky=W)
        self.vs_name = StringVar()
        self.vs = OptionMenu(self.vs_frame, self.vs_name, 'Pie Chart', 'Bar Chart')
        self.vs.grid(row = 38, column =0,  sticky=W, ipadx=20)
        self.vs_name.set('Bar Chart')

        #Visualize Button
        self.vs_button = Button(self.vs_frame, text="VISUALIZE", font="timesnewroman 12 bold",  fg="white", bg = "green", width=10, height=1, border = 3,
                                     relief='raised', command= lambda: self.save_v())
        self.vs_button.grid(row= 39, column=1, sticky=E)
    

        #Copyright
        Label(self.master, text='\u00a9 2021 Aga Khan University, Institute for Educational Development, Pakistan. All Rights Reserved',
              font="timesnewroman 7 bold" ).grid(row=40, column=0 , columnspan=4, sticky=N+E+W+S)

        #Version
        Label(self.master, text='Version 1.1 (Last Updated on 13th September 2021)',
              font="timesnewroman 7 bold").grid(row=41, column=0, columnspan=4, sticky=N+E+W+S)
                                                                                               
        self.filename.focus()

    def re_cols(self):
        self.items_listbox.delete(0,'end')
        self.columns_listbox.delete(0,'end')

    def re_items(self):
        self.items_listbox.delete(0,'end')

    def filter(self, df, filter_values):
        new_dict = {}
        for key, values in filter_values.items():
            if len(values)!=0:
                regex = '|'.join(values)
                col_find = df[df[key].str.match(regex, case=False, na=False)]
                new_dict[key]=col_find
        return new_dict
 

    def save_filter(self):
        clean = self.cleanup_dict(self.filterdict)
        filtered = self.filter(self.dfs2, clean)
        values = filtered.values()
        values_list = list(values)
 
        df_merged = reduce(lambda df1,df2: pd.merge(df1,df2), values_list)
 
        save_as = self.get_filter_filename()+'.xlsx'
        df_merged.to_excel(save_as)
        tkinter.messagebox.showinfo("Success", "File saved. ")

    def save_v(self):
        
        col_name = self.column_name.get()
        vs_type = self.vs_name.get()
        save_as = self.get_visual_filename()+'.png'
        
        try:
            if vs_type=='Bar Chart':
                vs_type = 'bar'
                ax = self.dfs2[col_name].plot(kind=vs_type, title =col_name, figsize=(10, 4), legend=True, fontsize=12)

            if vs_type=='Pie Chart':
                vs_type = 'pie'
                plot = self.dfs2.plot.pie(y=col_name, figsize=(5, 5))
            plt.savefig(save_as)
            tkinter.messagebox.showinfo("Success", "File saved. ")
                    
        except TypeError:
            tkinter.messagebox.showinfo("Error", "No numeric data to plot. ")    
    
    def cleanup_dict(self, dic):
        new_dict = {}
        for key, values in dic.items():
            if len(values)!=0:
                new_dict[key]=values[-1]
            else:
                new_dict[key]=[]
        return new_dict
            
                
    def special_char(self, s):
        chars = set('[*?/\:]')
        if any((c in chars) for c in s):
            return True
        else:
            return False

    def callback_item(self, event):
        self.selection_col = self.columns_listbox.get(self.columns_listbox.curselection())
        self.selection_item = [self.items_listbox.get(idx) for idx in self.items_listbox.curselection()]
        self.filterdict[self.selection_col].append(self.selection_item)

    def callback_column(self, event):
        selection = self.columns_listbox.curselection()
        self.items_listbox.delete(0,'end')
        if selection:
            self.selection_col = self.columns_listbox.get(self.columns_listbox.curselection())
            a = self.dfs2[self.selection_col].unique()
            for i in a:
                self.items_listbox.insert("end", i)
            self.selected = self.filterdict[self.selection_col]
            if not len(self.selected)== 0:
                for i in self.selected[-1]:
                    self.items_listbox.selection_set(self.items_listbox.get(0, "end").index(i))
            

    def main_filechecker(self):
        try:
            file = self.filename.get()
            sheet = self.sheet.get()
            start = int(self.start.get())-1
            stop = int(self.stop.get())-1
            incl = self.include_header_var.get()
            dfs1 = pd.read_excel(file, sheet_name=sheet, engine='openpyxl')
            dfs2 = dfs1.iloc[start-1:stop, :]
            if incl:
                dfs2.columns = dfs2.iloc[0]
                dfs2 = dfs2[1:]
            self.dfs2 = dfs2.fillna('')
            self.columns_listbox.delete(0,'end')
            self.items_listbox.delete(0,'end')
            self.filterdict = {}
            for i in list(dfs2):
                self.columns_listbox.insert("end", i)
                self.filterdict[i] = []
            self.columns = OptionMenu(self.vs_frame, self.column_name, *list(dfs2))
            self.columns.grid(row = 36, column =0,  sticky=W, ipadx=20)
            self.column_name.set('Select Column')
        except Exception:
            tkinter.messagebox.showinfo("Error", "File or Directory not found. ")
        
    def main_fieldchecker(self):
        try:
            if not os.path.exists(os.path.dirname(self.filename.get())):
                tkinter.messagebox.showinfo("Error", "Please enter valid file path. ")
            elif (self.sheet.get() == '') or (len(self.sheet.get())>31) or (self.special_char(self.sheet.get())):
                tkinter.messagebox.showinfo("Error", "Please enter valid sheet name. ")
            elif int(self.start.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid start row. ")
            elif int(self.stop.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid end row. ")
            else:
                self.main_filechecker()
        except ValueError:
            tkinter.messagebox.showinfo("Error", "Numeric values must be positive integers. ")
                          

    def get_source_filename(self):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
        self.a.set(askopenfilename(initialdir = desktop, title = "Select file", filetypes=(("All files", "*.*"),("Excel files", "*.xlsx"))))

    def get_filter_filename(self):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
        self.save_as = (filedialog.asksaveasfilename(initialdir = desktop, title = "Save As",filetypes = ( ("Excel files","*.xlsx"), ("All files","*.*"))))
        return self.save_as

    def get_visual_filename(self):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
        self.vs_save = (filedialog.asksaveasfilename(initialdir = desktop, title = "Save As", filetypes=(("All files", "*.*"),(" Image Files", "*.png"))))
        return self.vs_save
        


def main():
    root = Tk()
    myMainWindow = MainWindow(root)
    root.mainloop()

if __name__ == '__main__':
    main()

    
  
    
