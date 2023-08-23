import os
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.filedialog import askdirectory
from PyPDF2 import PdfReader, PdfWriter

def select_file():
    if entry1.get() is None:
        path=askdirectory()
        entry1.insert(0,path)
    else:
        path=askdirectory()
        entry1.delete(0,'end')
        entry1.insert(0,path)

def select_file2():
    if entry3.get() is None:
        path=askdirectory()
        entry3.insert(0,path)
    else:
        path=askdirectory()
        entry3.delete(0,'end')
        entry3.insert(0,path)

def go():
    l2=tk.Label(text='資料整理進度')
    l2.grid(row=4,column=0,columnspan=2)
    entry2=tk.Entry(width=60)
    entry2.grid(row=5,column=0,columnspan=2)
    bar = ttk.Progressbar(windows, mode='determinate', length=430)
    bar.grid(row=6,column=0,columnspan=2)
    l3=tk.Label(text='詳細內容')
    l3.grid(row=7,column=0,columnspan=2)
    txt1=tk.Text(width=90,height=10)
    txt1.grid(row=8,column=0,columnspan=3)
    l5=tk.Label(text='成功輸出列表')
    l5.grid(row=9,column=0,columnspan=2)
    txt2=tk.Text(width=90,height=10)
    txt2.grid(row=10,column=0,columnspan=3)
    windows.update()
    entry2.insert(0,'讀取'+entry1.get())
    txt1.insert('end','progress0:正在讀取{}\n\n'.format(entry1.get()))
    windows.update()
    
    alllist=[]
    if var1.get() is True:
        for i in os.listdir(entry1.get()):
            alllist.append('{}\{}'.format(entry1.get(),i))
    else:
        alllist.append(entry1.get())
    datao=entry3.get()
    mergedata=entry4.get().split(',')
    windows.update()
    
    listc=0
    for datalist in alllist:
        print('datalist=',datalist)
        datal=datalist
        listc=listc+1
        turn=0
        listdir=os.listdir(datal)
        for list in listdir:
            print('list=',list)
            turn=turn+1
            for m in mergedata:
                data=[]
                for root, dirs, files in os.walk(r'{0}\{1}'.format(datal,list)):
                    entry2.delete(0,'end')
                    entry2.insert(0,'整理專案中{}/{}：{:.2%}'.format(listc,len(alllist),turn/len(listdir)))
                    txt1.insert('end','讀取判讀專案{}\n\n'.format(list))
                    bar['value'] = int(turn/len(listdir)*100)
                    windows.update()
                    for f in files:
                        if m in f:
                            if 'all' not in f:
                                fullpath = os.path.join(root, f)
                                data.append(fullpath)
                data.sort(reverse=False)
                pdf_writer = PdfWriter()
                if len(data) > 0:
                    for pdf in data:
                        pdf_reader = PdfReader(pdf)
                        for page in range(pdf_reader.numPages):
                            pdf_writer.addPage(pdf_reader.getPage(page))
                    if not os.path.isdir(r'{0}\{1}'.format(datao,list)):
                        os.makedirs(r'{0}\{1}'.format(datao,list))
                    with open(r'{0}\{1}\{2}_all.pdf'.format(datao,list,m), 'wb') as out:
                        pdf_writer.write(out)
                    print(r'成功輸出：{0}\{1}\{2}_all.pdf'.format(datao,list,m))
                    txt1.insert('end','成功輸出：{0}/{1}/{2}_all.pdf\n'.format(datao,list,m))
                    txt2.insert('end','成功輸出：{0}/{1}/{2}_all.pdf\n'.format(datao,list,m))
                    txt1.see('end')
                    txt2.see('end')
                    windows.update()
                else:
                    print(r'沒有資料:{0}\{1}\{2}_all'.format(datal,list,m))
                    txt1.insert('end','沒有資料:{0}/{1}/{2}_all\n'.format(datal,list,m))
                    txt1.see('end')
                    windows.update()
        txt1.insert('end','\n========\合併結束\n\n')
        entry2.delete(0,'end')
        entry2.insert(0,'完成建管資料整理')
        bar['value'] = 100
        windows.update()
    messagebox.showinfo('任務訊息','建管資料整理程序完成')

# GUI介面的部分
windows=tk.Tk()
windows.title('建管資料檔案合併')
windows.geometry('680x500')
windows.minsize(650,500)
windows.resizable(False,False)

var1=tk.BooleanVar()

l1=tk.Label(text='請選取建管資料夾母路徑')
l1.grid(row=0,column=0)
entry1=tk.Entry(width=60)
entry1.grid(row=0,column=1)
bott1=tk.Button(text='匯入',command=select_file)
bott1.grid(row=0,column=2)
l7=tk.Label(text='合併類別編號(以,分隔)：')
l7.grid(row=1,column=0)
entry4=tk.Entry(width=60)
entry4.insert(0,'N0N013000001,N0N029000001')
entry4.grid(row=1,column=1)
l6=tk.Label()
check_btn1 = tk.Checkbutton(l6,text="我的母資料夾內\n含有年份資料夾",variable=var1)
l6.grid(row=1,column=2)
check_btn1.pack()
l4=tk.Label(text='請選擇輸出路徑')
l4.grid(row=2,column=0)
entry3=tk.Entry(width=60)
entry3.grid(row=2,column=1)
bott2=tk.Button(text='匯入',command=select_file2)
bott2.grid(row=2,column=2)
bott3=tk.Button(text='整理並輸出',command=go)
bott3.grid(row=3,column=1)

windows.mainloop()