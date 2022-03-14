from PIL import Image, ImageTk
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import glob


file_path = "D:\\2022\\0310"
z_range = 1500

x_list = list()
max_list = list()
    
wb = openpyxl.Workbook()
sheet = wb.active
file_name = "20220310"
sheet.title = "test_sheet_1"
all = file_path + "\\" + file_name
wb.save(all +".xlsx")
glob.glob("*.xlsx")

book = openpyxl.load_workbook(all + ".xlsx")
sheet = book[sheet.title]

z_list = []
lumi_list = []

for i in range(z_range):
        now = str(i) + "番目を処理中"
        print(now)
        name = "z = {:0=3}".format(int(i))+"[µm]"
        r = file_path + "\\" + name
        image = Image.open(r + ".png")
        size = image.size
        image = image.convert("L")
        array_image = np.asarray(image)
        Glay_P = 0
        for x in range(0,size[0]):
            for y in range(0,size[1]):
                Glay_P = Glay_P + image.getpixel((x,y))
        
        z_list.append(i)
        lumi_list.append(Glay_P)
    #print(lumi_list)
    #print(max(lumi_list))
    #print(max_list)
    #print("__________________________________________________")
plt.plot(z_list, lumi_list)
plt.show()
book.save(all + ".xlsx")
print("---解析終了---")
print("お疲れさまでした。")