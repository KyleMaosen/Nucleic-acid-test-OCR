"""
本功能基于百度PaddlePaddle的PaddleOcr框架进行ocr识别，识别效果优秀，但是识别速度较慢。
需要安装pandas，PaddlePaddle，paddleocr等库
dir是指装有图片的文件夹地址，注意是文件夹，且前面的字符“r”不可删去
识别完后会保存一个"核酸检测结果表.xls"，位于图片的文件夹地址处
"""
import os
import xlwt
from paddleocr import PaddleOCR
import skimage.io as io

caiyang_time_list=[]
jiance_time_list=[]
address_list=[]
results=[]

dir=r"xxxx"##请注意：本地址需要更改为你的地址，前面的r一定要留下，输入装有图片的文件夹地址，是文件夹地址

name_list = os.listdir(dir)
for i in range(len(name_list)):
    name_list[i]=name_list[i][:len(name_list)-4]
st=[0 for i in range(len(name_list))]
str = dir + '/*'
coll = io.ImageCollection(str)
ocr = PaddleOCR(use_angle_cls=False, lang="ch")
for i in range(len(coll)):
    try:
        result = ocr.ocr(coll[i], cls=False)
        caiyang_time_temporary_list = []
        jiance_time_temporary_list = []
        address_temporary_list = []
        jiance_result = []
        for line in result:
            if "查看更多核酸检测结果" in line[1][0] or "服务说明" in line[1][0]:
                break
            if "采样时间" in line[1][0]:
                caiyang_time_temporary_list.append(line[1][0][5:15])
            if "检测时间" in line[1][0]:
                jiance_time_temporary_list.append(line[1][0][5:15])
            if "市" in line[1][0] or "医" in line[1][0] or "室" in line[1][0] or "验" in  line[1][0]:
                address_temporary_list.append(line[1][0])
            if line[1][0] == "阴性" or line[1][0] == "阳性":
                jiance_result.append(line[1][0])
        caiyang_time_list.append(caiyang_time_temporary_list)
        jiance_time_list.append(jiance_time_temporary_list)
        address_list.append(address_temporary_list)
        results.append(jiance_result)
        print("第{}张图片已经处理完毕!".format(i + 1))
    except:
        st[i]=1
        caiyang_time_temporary_list = []
        jiance_time_temporary_list = []
        address_temporary_list = []
        jiance_result = []
        caiyang_time_list.append(caiyang_time_temporary_list)
        jiance_time_list.append(jiance_time_temporary_list)
        address_list.append(address_temporary_list)
        results.append(jiance_result)
        print("第{}张图片有问题，查看txt文件！".format(i + 1))

if sum(st)>0:
    with open(file=dir+'/未正常识别人员名单.txt', mode="w", encoding="utf-8") as f:
        f.write("未正常识别人员名单\n")
        for i in range(len(st)):
            if st[i]:
                f.write(name_list[i]+'\n')
    f.close()

book = xlwt.Workbook(encoding='utf-8', style_compression=0)
col = ("姓名", "检测情况")
sheet = book.add_sheet('data', cell_overwrite_ok=True)
for i in range(0, 2):
    sheet.write(0, i, col[i])

for i in range(len(name_list)):
    if st[i]==0:
        sheet.write(i + 1, 0, name_list[i])
        number=1
        for j in range(max([len(caiyang_time_list[i]),len(jiance_time_list[i]),len(address_list[i]),len(results[i])])):
            if j<len(caiyang_time_list[i]):
                sheet.write(i + 1, number, "采样时间:" + caiyang_time_list[i][j])
                number+=1
            if j < len(jiance_time_list[i]):
                sheet.write(i + 1,  number, "检测时间:" + jiance_time_list[i][j])
                number+=1
            if j<len(address_list[i]):
                sheet.write(i + 1, number, address_list[i][j])
                number+=1
            if j < len(results[i]):
                sheet.write(i + 1,number, results[i][j])
                number+=1
savepath =dir+'/核酸检测结果表.xls'

book.save(savepath)
print("检测完成!")
