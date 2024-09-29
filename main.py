import xlrd2
import xlwt
import os
#提示
os.system('cls')
print("Version:0.1.2     By:Name_Y")
print("使用时遇见问题联系QQ:3126127106")
print("不可修改本程序的文件名 StatisticalTool.exe ")
print("请确认本文件夹内仅有下列文件")
print("\t待处理的表格\n\tStatisticalTool.exe\n")
print("按回车键开始")
input()
###########
file_names = os.listdir("./")
n=1#表格行
#变量初始化
sum_num_all_all=0
sum_num_all_girl=0
sum_num_all_other=0
sum_num_member_all=0
sum_num_member_girl=0
sum_num_member_other=0
sum_money_theory=0
sum_money_actual=0
#新文档创建
wb = xlwt.Workbook(encoding="UTF-8")
ws = wb.add_sheet("sheet1")
#标题写入
ws.write(0,0,"文件")
ws.write(0,1,"学生总人数")
ws.write(0,2,"其中女生人数")
ws.write(0,3,"其中少数民族")
ws.write(0,4,"团员总人数")
ws.write(0,5,"其中女生人数")
ws.write(0,6,"其中少数民族")
ws.write(0,7,"理论团费")
ws.write(0,8,"实际团费")
#开始
for file in file_names:
    #排除本文件
    if file=="main.py" or file==".idea" or file=="StatisticalTool.exe" or file == "0End.xls":
        continue
       
    print(f"当前为第{n}个文件，名称为{file}")

    ###
    path_file="./"+file
    data = xlrd2.open_workbook(filename=path_file)
    table=data.sheets()[0]
    #读取信息
    num_all_all = table.cell_value(3,0)
    num_all_girl = table.cell_value(3,1)
    num_all_other = table.cell_value(3,2)
    num_member_all = table.cell_value(6,0)
    num_member_girl = table.cell_value(6,1)
    num_member_other = table.cell_value(6,2)
    money_theory = table.cell_value(2,3)
    money_actual = table.cell_value(2, 4)
    print(num_all_all,num_all_girl,num_all_other)
    print(num_member_all,num_member_girl,num_member_other)
    print(money_theory,money_actual)
    #计算总和
    sum_num_all_all=sum_num_all_all+num_all_all
    sum_num_all_girl=sum_num_all_girl+num_all_girl
    sum_num_all_other=sum_num_all_other+num_all_other
    sum_num_member_all=sum_num_member_all+num_member_all
    sum_num_member_girl=sum_num_member_girl+num_member_girl
    sum_num_member_other=sum_num_member_other+num_member_other
    sum_money_theory=sum_money_theory+money_theory
    sum_money_actual=sum_money_actual+money_actual
    #写入信息
    ws.write(n,0,file)
    ws.write(n,1,num_all_all)
    ws.write(n,2,num_all_girl)
    ws.write(n,3,num_all_other)
    ws.write(n,4,num_member_all)
    ws.write(n,5,num_member_girl)
    ws.write(n,6,num_member_other)
    ws.write(n,7,money_theory)
    ws.write(n,8,money_actual)
    n = n + 1
#写入总和
ws.write(n,0,"总计"+str(n-1))
ws.write(n,1,sum_num_all_all)
ws.write(n,2,sum_num_all_girl)
ws.write(n,3,sum_num_all_other)
ws.write(n,4,sum_num_member_all)
ws.write(n,5,sum_num_member_girl)
ws.write(n,6,sum_num_member_other)
ws.write(n,7,sum_money_theory)
ws.write(n,8,sum_money_actual)
#保存文件
wb.save("./0End.xls")
print("\n汇总完成，请查看 0End.xls ")
input("按回车键退出")
print("")