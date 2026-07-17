import os
import pandas
import openpyxl
import sys
from tkinter import Tk,Label,Button,StringVar,messagebox,filedialog,ttk

#検索場所決定
def select_src():
    src_var.set(filedialog.askdirectory())

#保存先決定
def select_save():
    save_var.set(filedialog.asksaveasfilename(defaultextension = ".xlsx",filetypes = [("Excel files","*.xlsx")]))


#ファイル総数
def count_file(src_dic):
    return sum(len(files) for _,_,files in os.walk(src_dic))
       

#フォルダ内検索
def create_file_list(src_dic,save_path,sousuu):

    dic = []
    bunsi = 0
            
    for ps,_,files in os.walk(src_dic):


        for cnt in range(len(files)):
            if stop_flag :
                return

            link = os.path.join(ps,files[cnt])
            link = f'=HYPERLINK("{link}","{files[cnt]}※リンク")'
            bunsi += 1 
            num_var.set(f"{bunsi}/{sousuu}")
            pgs["value"] = bunsi
            
            root.update()
            
            dic.append(
                {"保管先フォルダ":ps,
                 "ファイル名":files[cnt] ,
                 "リンク":link})

    df = pandas.DataFrame(dic)
    df.to_excel(save_path,index = False)

def cxl():
    global stop_flag
    stop_flag = True

#実行部分
def execute():
    try:

        global stop_flag
        stop_flag = False

        if not src_var.get() or not save_var.get():
            messagebox.showwarning("確認","フォルダ選択してください")

            return
        status_var.set("処理中")
        root.update()

        sousuu = count_file(src_var.get())
        pgs["maximum"] = sousuu
        pgs["value"] = 0
        
        create_file_list(src_var.get(),save_var.get(),sousuu)


        status_var.set("終了")

    except Exception as e:
        messagebox.showerror("エラー",str(e))
        
#画面
root = Tk()
root.title("階層内ファイルリスト化")
root.geometry("600x600")



src_var = StringVar()
save_var = StringVar()
status_var = StringVar(value = "待機中")
num_var = StringVar()
stop_flag = False


Label(root,
      text = "対象フォルダ",
      font = ("Meiryo",12)
      ).pack(anchor ="w",padx = 20)
Button(
    root,text = "参照",
    command = select_src,
    font=("Meiryo",12),
    width = 12,
    height = 1
    ).pack(anchor = "w",padx = 20)

Label(root,
      textvariable = src_var,
      font = ("Meiryo",12)
      ).pack(anchor = "w",padx = 20) 

Label(root,text = "保管先",
      font = ("Meiryo",12)
      ).pack(anchor = "w",padx = 20)

Button(root,text = "参照",
       command = select_save,
       font = ("Meiryo",12),
       width = 12,
       height =1
       ).pack(anchor = "w",padx = 20)

Label(root,
      textvariable = save_var,
      font = ("Meiryo",12)
      ).pack(anchor = "w" ,padx =20 )

Button(root,text = "実行",
       command = execute,
       font = ("Meiryo",12),
       width = 12,
       height = 1).pack(anchor = "w" ,padx = 20)


Button(root,text = "キャンセル",
       command = cxl,
       font = ("Meiryo",12),
       width = 12,
       height = 1).pack(anchor = "w" ,padx = 20 ,pady=20)


Label(root,textvariable = status_var,font = ("Meiryo",12)).pack(anchor = "w" ,padx = 20)
Label(root,textvariable = num_var,font = ("Meiryo",12)).pack(anchor = "w",padx = 20)

pgs = ttk.Progressbar(root,orient = "horizontal",length = 400,mode ="determinate")
pgs.pack(anchor = "w",padx = 20,pady = 5)

#wb = openpyxl.load_workbook("階層一覧.xlsx") 
#ws = wb.worksheets[0]


#for cnt in range(end):

 #   print(cnt)
  #  for cntc in range(2,5):
        
   #     link = str(ws.cell(cnt+1,cntc).value)
    #    ws.cell(cnt+1,cntc).hyperlink = link
     #   ws.cell(cnt+1,cntc).style ="Hyperlink"


#wb.save("階層一覧.xlsx")
root.mainloop()
print("end")
