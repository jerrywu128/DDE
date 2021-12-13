import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import time    

root = tk.Tk()
root.title('DDE')
root.configure(bg='#ffffff')
root.geometry('720x700')
cv = tk.Canvas(root, background='white')
cv.place(relx=0.0,rely=0.1,relwidth=1,relheight=0.9)
def DDE(file_path):
        df = pd.read_excel(file_path,header = None)
        rows = df.shape[0]-1
        cols = df.shape[1]

        hours = []
        I_K=[]
        J_L=[]
        I_J=[]
        N_M=[]

        num=0
        #plt.ion() #互動效果
        
        while num<rows:
            df = pd.read_excel(file_path,header = None)#更新excel
            rows =df.shape[0]-1#讀取新的資料筆數扣掉標題
            print(rows)
            plt.clf()
            plt.close()
            I = df.iat[num+1,8]
            J = df.iat[num+1,9]
            K = df.iat[num+1,10]
            L = df.iat[num+1,11]
            M = df.iat[num+1,12]
            N = df.iat[num+1,13]

            IK = I/K
            JL = J/L
            IJ = IK-JL
            NM = N-M
            
            f = plt.figure(figsize=(7,7),dpi=80,linewidth = 2)
            I_K.append(IK)
            J_L.append(JL)
            I_J.append(IJ)
            N_M.append(NM)

            hours.append(num+1)
            plt.subplot(2,2,1)
            plt.plot(hours,I_K,'s-r', label="I/K")
            plt.title("I/K",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("price",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            
            plt.subplot(2,2,2)
            plt.plot(hours,J_L,'o-',color = 'g', label="J/L")
            plt.title("J/L",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("price",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            
            plt.subplot(2,2,3)
            plt.plot(hours,I_J,'s-',color = 'b', label="I-J")
            plt.title("I-J",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("price",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            
            plt.subplot(2,2,4)
            plt.plot(hours,N_M,'o-',color = 'g', label="N-M")
            plt.title("N-M",x=0.5,y=1.03)
            plt.xlabel("count",fontsize=10,labelpad = 10)
            plt.ylabel("price",fontsize=10,labelpad = 0)
            plt.legend(loc ="best",fontsize=20)
            plt.subplots_adjust(hspace=0.4)
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
        #plt.ioff()
        #plt.draw()
    
def test():
    file_path = filedialog.askopenfilename()
    listbox_file.insert(0,file_path)
    DDE(file_path)

listbox_file=tk.Listbox(root,justify=tk.LEFT)
listbox_file.place(relx=0.0,rely=0.05,relwidth=0.5,height=30)
bt = tk.Button(root,text='選取檔案',fg='black',command=test)
bt.place(relx=0.5,rely=0.05,width=100,height=30)

root.mainloop()
