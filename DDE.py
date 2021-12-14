import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import time    
from matplotlib.font_manager import FontProperties
running = True
hours = []
I_K=[]
J_L=[]
I_J=[]
N_M=[]
fig_5 = []
num=0

def DDE(file_path):
        df = pd.read_excel(file_path,header = None)
        rows = df.shape[0]-1
        cols = df.shape[1]
        global num
        global running 
        global hours 
        global I_K
        global J_L
        global I_J
        global N_M
        global fig_5
        
       
        #plt.ion() #互動效果
        
        while (num<rows)and(running):
    
            df = pd.read_excel(file_path,header = None)#更新excel
            rows =df.shape[0]-1#讀取新的資料筆數扣掉標題
            print(rows)
            plt.clf()
            plt.close()
            plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
            plt.rcParams['axes.unicode_minus'] = False

            fig1 = df.iat[num+1,15]
            fig2 = df.iat[num+1,16]
            fig3 = df.iat[num+1,17]
            fig4 = df.iat[num+1,18]
            fig5 = df.iat[num+1,19]

            
            f = plt.figure(figsize=(7,7),dpi=80,linewidth = 2)
            I_K.append(fig1)
            J_L.append(fig2)
            I_J.append(fig3)
            N_M.append(fig4)
            fig_5.append(fig5)

            hours.append(num+1)
            plt.subplot(2,3,1)
            plt.plot(hours,I_K,'s-r', label="I/K")
            plt.title("I/K",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("value",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            
            plt.subplot(2,3,2)
            plt.plot(hours,J_L,'o-',color = 'g', label="J/L")
            plt.title("J/L",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("value",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            
            plt.subplot(2,3,3)
            plt.plot(hours,I_J,'s-',color = 'b', label="I-J")
            plt.title("I-J",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("value",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            
            plt.subplot(2,3,4)
            plt.plot(hours,N_M,'o-',color = 'm', label="N-M")
            plt.title("N-M",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("value",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            plt.subplots_adjust(hspace=0.4)

            plt.subplot(2,3,5)
            plt.plot(hours,fig_5,'o-',color = 'y', label=titleString.get())
            plt.title(titleString.get(),x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("value",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            plt.subplots_adjust(hspace=0.4)
            plt.tight_layout()
            #plt.pause(1)
            num=num+1
           
            canvas = FigureCanvasTkAgg(f, master=cv)
            canvas.draw_idle()
            
            plot_widget = canvas.get_tk_widget()
           # plot_widget.pack(side=tk.TOP,fill=tk.BOTH,expand=1)
            plot_widget.place(relwidth=1, relheight=1)
            
            root.update_idletasks()
            root.update()
            time.sleep(0.5)
        if num>=rows:
           num=0
           hours = []
           I_K=[]
           J_L=[]
           I_J=[]
           N_M=[]
           fig_5 = []
        #plt.ioff()
        #plt.draw()
    
def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    listbox_file.insert(0,file_path)
    Bt_run['state']=tk.NORMAL
   

def run_fig():
    global running
    running = True
    Bt_stop['state']=tk.NORMAL
    Bt_run['state']=tk.DISABLED
    DDE(listbox_file.get(0))
    Bt_run['state']=tk.NORMAL
    Bt_stop['state']=tk.DISABLED

def fig_stop():
    global running
    running = False

    Bt_run['state']=tk.NORMAL
    Bt_stop['state']=tk.DISABLED

if __name__=="__main__":
    root = tk.Tk()
    root.title('DDE')
    root.configure(bg='#ffffff')
    root.geometry('1080x700')
    cv = tk.Canvas(root, background='white')
    cv.place(relx=0.0,rely=0.1,relwidth=1,relheight=0.9)

    listbox_file=tk.Listbox(root,justify=tk.LEFT)
    listbox_file.place(relx=0.0,rely=0.05,relwidth=0.5,height=30)
    Bt_choose = tk.Button(root,text='選取檔案',fg='black',command=choose_file)
    Bt_choose.place(relx=0.5,rely=0.05,width=100,height=30)
    
    Bt_run = tk.Button(root,text='執行',fg='black',command=run_fig,state=tk.DISABLED)
    Bt_run.place(relx=0.60,rely=0.05,width=60,height=30)
 
    Bt_stop = tk.Button(root,text='暫停',fg='black',command=fig_stop,state=tk.DISABLED)
    Bt_stop.place(relx=0.66,rely=0.05,width=60,height=30)

    lb = tk.Label(root,text='自訂值:\nex:i2-j3',fg='black')
    lb.place(relx=0.73,rely=0.05,width=50,height=30)
    titleString = tk.StringVar()
    titleString.set("自訂")
    entry5 = tk.Entry(root, width=20, textvariable=titleString)
    entry5.place(relx=0.78,rely=0.05,width=100,height=30)
    root.mainloop()
