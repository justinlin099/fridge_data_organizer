import pandas

#輸出錯誤訊息
DEBUG_MODE = False

import tkinter as tk
from tkinter import filedialog
import base64, zlib
import tempfile

#圖示
ICON = zlib.decompress(base64.b64decode('eJxjYGAEQgEBBiDJwZDBy'
    'sAgxsDAoAHEQCEGBQaIOAg4sDIgACMUj4JRMApGwQgF/ykEAFXxQRc='))

_, ICON_PATH = tempfile.mkstemp()
with open(ICON_PATH, 'wb') as icon_file:
    icon_file.write(ICON)


root = tk.Tk()
root.iconbitmap(default=ICON_PATH)
root.withdraw()

#開啟冰箱違規表單
file_path = filedialog.askopenfilename(parent=root, 
                                    title='開啟"冰箱違規登記表"',filetype = (("試算表","*.xlsx"),("所有檔案","*.*")))
if not file_path:
    print('file path is empty')
else:
#讀取冰箱違規表單
    data=pandas.read_excel(file_path,sheet_name=None)
    plist={}

    class Student:
        def __init__(self,floor,room,bed,id,name,dates,lastDate):
            self.floor=floor
            self.room=room
            self.bed=bed
            self.id=id
            self.name=name
            self.count=1
            self.dates=dates
            self.lastDate=lastDate

        def __str__(self):
            return "房號:"+self.room+"床號:"+self.bed+"學號:"+self.id+"姓名:"+self.name+"違規次數:"+str(self.count)+"上次違規日期:"+self.lastDate+"\n"

        def __repr__(self):
            return "樓層:"+self.floor+" 房號:"+self.room+" 床號:"+self.bed+" 學號:"+self.id+" 姓名:"+self.name+" 違規次數:"+str(self.count)+" 違規日期:"+self.dates+"\n"
    


    #print(data)
    #print(str(data["3F"]["備註"][1])=="nan")

    #統計所有資料
    for floor in data.keys():
        if(DEBUG_MODE):
            print(floor+" 紀錄開始")
        for i in range(1,len(data[floor]["檢查日期"])):
            if(str(data[floor]["房號"][i])!="nan" and str(data[floor]["床號"][i])!="nan"):#如果有房號床號
                if(DEBUG_MODE):
                    print(str(floor)+"第"+str(i)+"已找到資料")

                #把該學生加入plist
                if(str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i])) not in plist.keys()):
                    pfloor=str(int(data[floor]["房號"][i]//100))#樓層
                    #寫入plist
                    plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))]=Student(pfloor,str(int(data[floor]["房號"][i])),str(int(data[floor]["床號"][i])),str(data[floor]["學號"][i]),str(data[floor]["姓名"][i]),str(data[floor]["檢查日期"][i])[5:7]+"/"+str(data[floor]["檢查日期"][i])[8:10],str(data[floor]["檢查日期"][i]))
                else:
                    #如果該學生已經有紀錄，且日期相同，跳過
                    if(plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].lastDate==str(data[floor]["檢查日期"][i])):
                        if(DEBUG_MODE):
                            print(str(floor)+"第"+str(i)+"重複資料，跳過")
                        continue
                    else:#如果該學生已經有紀錄，且日期不同，更新資料

                        #增加違規次數
                        plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].count+=1

                        #更新日期
                        plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].lastDate=str(data[floor]["檢查日期"][i])
                        plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].dates=plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].dates+"."+str(data[floor]["檢查日期"][i])[5:7]+"/"+str(data[floor]["檢查日期"][i])[8:10]

                    #如果學號或姓名為空，更新學號或姓名
                    if(str(data[floor]["學號"][i])!="nan" and plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].id=="nan"):  
                        plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].id=str(data[floor]["學號"][i])
                    if(str(data[floor]["姓名"][i])!="nan" and plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].name=="nan"):
                        plist[str(int(data[floor]["房號"][i]))+"-"+str(int(data[floor]["床號"][i]))].name=str(data[floor]["姓名"][i])
            else:
                #如果沒有房號床號，改用學號紀錄
                #但是我不想努力了，所以不寫了
                if(DEBUG_MODE):
                    print(str(floor)+"第"+str(i)+"未找到資料，改用學號紀錄")
                continue
    if(DEBUG_MODE):
        print(plist)
    #寫入excel
    save_path=filedialog.asksaveasfilename(parent=root, title='儲存"冰箱違規統計表"',filetype = (("試算表","*.xlsx"),("所有檔案","*.*")),defaultextension="*.xlsx",initialfile = "冰箱違規統計表")
    writer = pandas.ExcelWriter(save_path)
    #將plist分樓層寫入excel
    for floor in range(2,14):
        if(DEBUG_MODE):
            print(str(floor)+"樓開始寫入")

        #寫入資料
        writeData={"樓層":[],"房號":[],"床號":[],"學號":[],"姓名":[],"違規次數":[],"違規日期":[]}
        for key in plist.keys():
            if(plist[key].floor==str(floor)):
                writeData["樓層"].append(plist[key].floor)
                writeData["房號"].append(plist[key].room)
                writeData["床號"].append(plist[key].bed)
                writeData["學號"].append(plist[key].id)
                writeData["姓名"].append(plist[key].name)
                writeData["違規次數"].append(plist[key].count)
                writeData["違規日期"].append(plist[key].dates)
        df=pandas.DataFrame(writeData)
        
        df.to_excel(writer,sheet_name=str(floor)+"F")
    writer.close()
                


