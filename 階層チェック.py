import os
import pandas as pd
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
    # dic = [{"保管先フォルダ":ps,"ファイル名":file} for ps,_,files in os.walk(src_dic) for file in files]
    # bunsi = len(dic)
    
    # print(bunsi)
    # pgs["value"] = bunsi
    # root.update()

    for ps,_,files in os.walk(src_dic):
        for file in files:
            if stop_flag :
                return
            link = os.path.join(ps,file)
            link_h = f'=HYPERLINK("{link}","{file}※リンク")'
            
            # with open(link,"rb") as f:
            #     rr = f.read()

            dic.append(
                {"保管先フォルダ":ps,
                 "ファイル名":file ,
                 "リンク":link_h,
                 "サイズ":os.path.getsize(link),
                #  "バイト":rr,
                 "作成日時":os.path.getmtime(link)})
            bunsi += 1 
            num_var.set(f"{bunsi}/{sousuu}")
            pgs["value"] = bunsi
            root.update()

    df = pd.DataFrame(dic)
    
    df["f"] = df["ファイル名"].str.contains(r"\.")
    df["拡張子"] = df["ファイル名"].str.split(".").str[-1]
    df["拡張子"] = df["拡張子"].where(df["f"],"")
    df = df.drop(columns="f")
    df["作成日時"] = pd.to_datetime(df["作成日時"],unit="s").dt.tz_localize("UTC").dt.tz_convert("Asia/Tokyo").dt.strftime("%Y-%m-%d")
    df_p = df.groupby("保管先フォルダ").agg(リンク=("リンク","count"),サイズ=("サイズ","sum"))
    df.insert(loc=1,column="第1階層" ,value=df["保管先フォルダ"].str.split(r"\\").str[0])
    df.insert(loc=2,column="第2階層" ,value=df["保管先フォルダ"].str.split(r"\\").str[1])
    df.insert(loc=3,column="第3階層" ,value=df["保管先フォルダ"].str.split(r"\\").str[2])
    #df = df.pivot_table(index="保管先フォルダ",columns="拡張子",values="サイズ",aggfunc="count")
    # df.to_excel(save_path)#,index = False)    
    with pd.ExcelWriter(save_path) as wr:
        df.to_excel(wr,sheet_name="一覧")
        df_p.to_excel(wr,sheet_name="集計")

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

#ボタン配置
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


root.mainloop()
print("end")
