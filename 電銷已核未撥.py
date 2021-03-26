import tkinter.tix
from tkinter.tix import Tk, Control, ComboBox
from tkinter.messagebox import showinfo, showwarning, showerror
import tkinter as tk
import tkinter.ttk as ttk
import PIL
from PIL import ImageTk, Image, ImageSequence
import time
import pyodbc
from base64 import b16encode

def rgb_color(rgb):
    return(b'#' + b16encode(bytes(rgb)))

#連接資料庫
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\已核未撥.accdb')
cursor = conn.cursor()

#設定主視窗
root = tk.Tk()
root.title("永豐銀行電銷業務部©Victor Lin_M06429")
root.geometry("1920x1080")
root.resizable(width=True, height=True)

#子視窗背景建立路徑
path2 = r'\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\Portable Python-3.8.2\子視窗背景.png'
img2 = ImageTk.PhotoImage(Image.open(path2))

#主視窗背景
canvas = tk.Canvas(root, width=1920,height=1080,bd=0, highlightthickness=0)
imgpath = r'\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\Portable Python-3.8.2\主視窗背景.png'
img = Image.open(imgpath)
photo = ImageTk.PhotoImage(img)


canvas.create_image(960, 505, image=photo)
canvas.pack()




#定義子視窗
def second_win():
    top = tk.Toplevel()
    top.title('已核未撥案件回報系統')
    top.geometry("1920x1080")

    #子視窗背景
    canvas2 = tk.Canvas(top, width=1920,height=1080,bd=0, highlightthickness=0)
    canvas2.create_image(960, 505, image=img2)
    canvas2.pack()
#===============================================================================================================================

    #定義清除函數
    def Reset():
        entry.delete(0,END),
        entry2.delete(0,END),
        combo4.delete(0,END),
        entry4.delete(0,END),
        entry5.delete(0,END),
        combo.delete(0,END),
        entry6.delete(0,END),
        combo2.delete(0,END),
        entry7.delete(0,END),


#===============================================================================================================================

    #定義搜尋函數
    def search():
        try:
            cursor=conn.cursor()
            sql = "select * from 已核未撥回報 where 申請書編號 = ?"
            cursor.execute(sql,(entry.get(),))
            row=cursor.fetchone()


            entryvar.set(row[0])
            entryvar2.set(row[1])
            combovar4.set(row[2])
            entryvar4.set(row[3])
            entryvar5.set(row[4])
            combovar.set(row[5])
            entryvar6.set(row[6])
            combovar2.set(row[7])
            entryvar7.set(row[8])
            entryvar8.set(row[9])

            conn.commit()
        except:
            tkinter.messagebox.showinfo("已核未撥案件回報系統","查無此筆案件")
            Reset()




    #建立子視窗 search 按鈕
    searchbutton = tk.Button(top, command=search,text='搜尋',bg='#AE0000',fg='white',activeforeground="black",activebackground='white',font=("微軟正黑體", 24, 'bold'),cursor='hand2', width=3,height=1,bd=2)
    searchbutton.pack()



    canvas2.create_window(740, 95,window=searchbutton)


#===============================================================================================================================


    #定義更新函數
    def update():
        cursor=conn.cursor()

        cursor.execute("UPDATE 已核未撥回報 SET  業務員=?, 組別=?, 送簽核日=?, 金額=?, 可否撥款=?, 預定撥款日=?, 暫時無法撥款說明=?, 備註=?, 時間=? WHERE 申請書編號 = ?",(



        entryvar2.get(),
        combovar4.get(),
        entryvar4.get(),
        entryvar5.get(),
        combovar.get(),
        entryvar6.get(),
        combovar2.get(),
        entryvar7.get(),
        entryvar8.get(),
        entryvar.get()
        ))

        conn.commit()
        tkinter.messagebox.showinfo("已核未撥案件回報系統","更新成功！")
        Reset()


    #建立子視窗 search 按鈕
    updatebutton = tk.Button(top, command=update,text='更新',bg='#AE0000',fg='white',activeforeground="black",activebackground='white',font=("微軟正黑體", 24, 'bold'),cursor='hand2', width=3,height=1,bd=2)
    updatebutton.pack()



    canvas2.create_window(830, 95,window=updatebutton)




#===============================================================================================================================

    #定義新增函數
    def addData():
        if entry.get() =="" or entry2.get() =="" or combo4.get() =="" or entry4.get() =="" or entry5.get() =="" or entry6.get() =="" or combo.get() =="" or combo2.get() =="":
            tkinter.messagebox.showerror("已核未撥案件回報系統","請輸入完整資料")
        else :
            cursor=conn.cursor()
            cursor.execute("insert into 已核未撥回報 values(?,?,?,?,?,?,?,?,?,?)",(


            entryvar.get(),
            entryvar2.get(),
            combovar4.get(),
            entryvar4.get(),
            entryvar5.get(),
            combovar.get(),
            entryvar6.get(),
            combovar2.get(),
            entryvar7.get(),
            entryvar8.get(),
            ))
            conn.commit()
            tkinter.messagebox.showinfo("已核未撥案件回報系統","儲存成功！")
            Reset()




    #建立子視窗 addData 按鈕
    addDatabutton = tk.Button(top, command=addData,text='新增',bg='#AE0000',fg='white',activeforeground="black",activebackground='white',font=("微軟正黑體", 24, 'bold'),cursor='hand2', width=3,height=1,bd=2)
    addDatabutton.pack()



    canvas2.create_window(920, 95,window=addDatabutton)

#===============================================================================================================================

    #定義組別清單函數
    def display():
        cursor=conn.cursor()
        sql2 = "select * from 已核未撥回報 where 組別 = ?"
        cursor.execute(sql2,(combovar3.get(),))
        result = cursor.fetchall()
        if len(result)!=0:
            records.delete(*records.get_children(),)
            for row in result:
                records.insert("",END,values = row[:9])

            conn.commit()

    def Info(ev):
        viewInfo = records.focus()
        learnerData = records.item(viewInfo)
        row = learnerData['values']
        entryvar.set(row[0])
        entryvar2.set(row[1])
        combovar4.set(row[2])
        entryvar4.set(row[3])
        entryvar5.set(row[4])
        combovar.set(row[5])
        entryvar6.set(row[6])
        combovar2.set(row[7])
        entryvar7.set(row[8])



    displaybutton = tk.Button(top, command=display,text='查詢',bg='#AE0000',fg='white',activeforeground="black",activebackground='white',font=("微軟正黑體", 24, 'bold'),cursor='hand2', width=3,height=1,bd=2)
    displaybutton.pack()

    canvas2.create_window(1660, 95,window=displaybutton)
#===============================================================================================================================

    #定義業務員清單函數
    def display2():
        cursor=conn.cursor()
        sql3 = "select * from 已核未撥回報 where 業務員 = ?"
        cursor.execute(sql3,(entryvar2.get(),))
        result = cursor.fetchall()
        if len(result)!=0:
            records.delete(*records.get_children(),)
            for row in result:
                records.insert("",END,values = row[:9])

            conn.commit()

    def Info2(ev):
        viewInfo = records.focus()
        learnerData = records.item(viewInfo)
        row = learnerData['values']
        entryvar.set(row[0])
        entryvar2.set(row[1])
        combovar4.set(row[2])
        entryvar4.set(row[3])
        entryvar5.set(row[4])
        combovar.set(row[5])
        entryvar6.set(row[6])
        combovar2.set(row[7])
        entryvar7.set(row[8])



    displaybutton2 = tk.Button(top, command=display2,text='業務員清單查詢',bg='#AE0000',fg='white',activeforeground="black",activebackground='white',font=("微軟正黑體", 24, 'bold'),cursor='hand2', width=11,height=1,bd=2)
    displaybutton2.pack()

    canvas2.create_window(820, 187,window=displaybutton2)
#===============================================================================================================================

    #定義時間函數
    def gettime():
        entryvar8.set(time.strftime("%Y-%m-%d  %H:%M:%S"))
        top.after(1000, gettime)

#===============================================================================================================================


    style = ttk.Style()
    #Pick a theme
    style.theme_use("default")
    style.configure("Treeview.Heading", font=("微軟正黑體", 16, 'bold'), background="#AE0000",foreground="white", fieldbackground="#AE0000")
    style.configure("Treeview",rowheight= 24, font=("微軟正黑體", 16), background="lightgrey",foreground="white", fieldbackground="lightgrey")
    # Change selected color
    style.map('Treeview',
	background=[('selected', '#0080FF')])


    #Treeview 建立
    scroll_y = Scrollbar(top,orient = VERTICAL)

    records = ttk.Treeview(top,height = 20,columns = ("申請書編號","業務員","組別","送簽核日","金額","可否撥款","預定撥款日","暫時無法撥款說明","備註"),yscrollcommand = scroll_y.set)
    scroll_y.pack()




    records.heading("申請書編號",text="申請書編號")
    records.heading("業務員",text="業務員")
    records.heading("組別",text="組別")
    records.heading("送簽核日",text="送簽核日")
    records.heading("金額",text="金額")
    records.heading("可否撥款",text="可否撥款")
    records.heading("預定撥款日",text="預定撥款日")
    records.heading("暫時無法撥款說明",text="暫時無法撥款說明")
    records.heading("備註",text="備註")

    records['show'] = 'headings'

    records.column("申請書編號", width = 180)
    records.column("業務員", width = 180)
    records.column("組別", width = 180)
    records.column("送簽核日", width = 180)
    records.column("金額", width = 180)
    records.column("可否撥款", width = 180)
    records.column("預定撥款日", width = 180)
    records.column("暫時無法撥款說明", width = 200)
    records.column("備註", width = 440)

    records.pack()
    records.bind("<ButtonRelease-1>",Info)
    records.bind("<ButtonRelease-1>",Info2)

    canvas2.create_window(959, 755,window=records)


#===============================================================================================================================
    #申請書編號 entry 建立
    entryvar = tk.StringVar()
    entry = tk.Entry(top, insertbackground='black',font=("微軟正黑體",28),highlightcolor="#AE0000" ,highlightthickness =2, textvariable=entryvar)
    entry.pack()
    canvas2.create_window(535, 95, width=300,height=68,
                                   window=entry)
    #定義entry輸入字元數限制
    def character_limit(entryvar):
        if len(entryvar.get()) > 0:
            entryvar.set(entryvar.get()[:12])

    entryvar.trace("w", lambda *args: character_limit(entryvar))

    #業務員 entry 建立
    entryvar2=tk.StringVar()
    entry2 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",28),highlightcolor="#AE0000" ,highlightthickness =2, textvariable=entryvar2)
    entry2.pack()
    canvas2.create_window(535, 187, width=300,height=68,
                                       window=entry2)

    #組別 entry 建立
    combovar4=tk.StringVar()
    combo4 = ttk.Combobox(top,value = ['李叢華','邱秀環','郭曉清','鄭琬儒','數位台中林佳鳳','莊季華','黃湘榆','吳萱淑','蘇筱榛','數位台北'],font=("微軟正黑體",28),textvariable=combovar4)
    combo4.pack()
    canvas2.create_window(535, 282, width=300,height=68,
                                       window=combo4)

    #送簽核日 entry 建立
    entryvar4=tk.StringVar()
    entry4 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",28),highlightcolor="#AE0000" ,highlightthickness =2, textvariable=entryvar4)
    entry4.pack()
    canvas2.create_window(812, 368, width=300,height=68,
                                       window=entry4)

    #金額 entry 建立
    entryvar5=tk.StringVar()
    entry5 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",28),highlightcolor="#AE0000" ,highlightthickness =2, textvariable=entryvar5)
    entry5.pack()
    canvas2.create_window(535, 452, width=300,height=68,
                                       window=entry5)



    #可否撥款 combobox 建立
    combovar=tk.StringVar()
    combo = ttk.Combobox(top,value = ['可','不可'],font=("微軟正黑體",28) ,textvariable=combovar)

    combo.pack()

    canvas2.create_window(1455, 187, width=300, height=68,window=combo)


    #預定撥款日 entry 建立
    entryvar6=tk.StringVar()
    entry6 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",28),highlightcolor="#AE0000" ,highlightthickness =2, textvariable=entryvar6)
    entry6.pack()
    canvas2.create_window(1760, 282, width=300,height=68,
                                       window=entry6)


    #暫時無法撥款說明 combobox 建立
    combovar2=tk.StringVar()
    combo2 = ttk.Combobox(top,value = ['額度不滿意','利率不滿意','已撥他行','暫時無資金需求','失聯'],font=("微軟正黑體",26),textvariable=combovar2)

    combo2.pack()

    canvas2.create_window(1760, 368, width=300, height=68,window=combo2)

    #備註 entry 建立
    entryvar7=tk.StringVar()
    entry7 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",28),highlightcolor="#AE0000" ,highlightthickness =2, textvariable=entryvar7)
    entry7.pack()
    canvas2.create_window(1560, 452, width=700,height=68,
                                       window=entry7)

    #清單 combobox 建立
    combovar3=tk.StringVar()
    combo3 = ttk.Combobox(top,value = ['李叢華','邱秀環','郭曉清','鄭琬儒','數位台中林佳鳳','莊季華','黃湘榆','吳萱淑','蘇筱榛','數位台北'],font=("微軟正黑體",28),textvariable=combovar3)

    combo3.pack()

    canvas2.create_window(1455, 95, width=300, height=68,window=combo3)

    #時間 entry 建立
    entryvar8=tk.StringVar()
    entry8 = tk.Label(top,font=("微軟正黑體",14,"bold"),fg = "white",bg = rgb_color((26, 70, 135)),highlightcolor="#AE0000" ,highlightthickness =2, textvariable=entryvar8)
    entry8.pack()
    canvas2.create_window(1805, 17, width=350,height=15,
                                       window=entry8)
    gettime()


    # 下拉框颜色
    #combostyle = ttk.Style()
    #combostyle.theme_create('combostyle', parent='alt',
    #                        settings={'TCombobox':
    #                            {'configure':
    #                                {
    #                                    'foreground': 'blue',  # 前景色
    #                                    'selectbackground': 'black',  # 选择后的背景颜色
    #                                    'fieldbackground': 'white',  # 下拉框颜色
    #                                    'background': 'red',  # 下拉按钮颜色
    #                                }}}
    #                        )
    #combostyle.theme_use('combostyle')


    top.mainloop()





#===============================================================================================================================

#建立按鈕
button=Button(root,text='進入系統',command=second_win,bg='#AE0000',fg='white',activeforeground="black",activebackground='#66B3FF',font=("微軟正黑體", 80, 'bold'),cursor='hand2',bd=0, width=11, height=2)
button.pack()

canvas.create_window(1400, 500,window=button)


root.mainloop()
