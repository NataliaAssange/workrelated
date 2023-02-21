import os
import datetime
date=datetime.date.today()
if date.month==1:
    year=str(date.year-1)
    month=str(12)
else:
    year=str(date.year)
    month='%02d'%(date.month-1)
date1=str(date)
period=month+'.'+year
def ReFileName(path):
    # 对目录下的文件进行遍历
    for oldname in sorted(os.listdir(path)):
        if  period not in oldname:
            midname=str(oldname.split(' - ')[1])
            midname2=str(midname[:-7])
            ext=str(oldname.split('.')[-1])
            newName = date1+' - '+midname2+' - '+period+' - v01.'+ext
            print("FFF ", os.path.exists(os.path.join(path, oldname)))
            print("FFE ", os.path.exists(os.path.join(path, newName)))
            os.rename(os.path.join(path, oldname), os.path.join(path, newName))
    print("文件名已统一修改成功")
  
if __name__ == '__main__':
    path = r'C:\Users\CN0CHHG\Documents\Monthly Reporting Files\\'+year+ '-' +month
    path1=path+r'\Business Region Profitability'
    path2=path+r'\Sales Performance Report'
    path3=path+r'\Smart Report'
    ReFileName(path1)
    ReFileName(path2)
    ReFileName(path3)