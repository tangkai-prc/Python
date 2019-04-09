#coding: utf-8
#import xlwt
#import xlrd
#import xlsxwriter
import time 
import os
import re
from hashlib import md5, sha1
from zlib import crc32
from openpyxl import Workbook
from openpyxl.styles import colors, Font, Fill, NamedStyle
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl import load_workbook
from openpyxl.writer.excel import ExcelWriter
import xml.etree.ElementTree as ET

EXCEL_NAME = 'NSR_VersionTool.xls'

def getfileCreateTime(file):
    return time.ctime(os.path.getmtime(file))

def searchFile(FileType):
    filelist=[]
    for root, dirs, files in os.walk(".", topdown=False):
        for name in files:
            str = os.path.normpath(os.path.join(root, name))
            if str.split('.')[-1] == FileType:
                filelist.append(str)

    if len(filelist) == 0:
        return None
    else:
        return filelist

def getICDVersion(fileName):
    icd = ET.parse(fileName)
    root = icd.getroot()
    version = None
    for child in root:
        # 第二层节点的标签名称和属性.
        #print(child.tag,":", child.attrib) 
        for item in child.attrib:
            if item == 'configVersion' :    #item是attribute返回的字段?
                version =  child.attrib[item]
    return version

def getMd5(filename): #计算md5
    mdfive = md5()
    with open(filename, 'rb') as f:
        mdfive.update(f.read())
    return str.upper(mdfive.hexdigest())

def getCrc32(filename): #计算crc32
    with open(filename, 'rb') as f:
        res = hex(crc32(f.read()))
    return str.upper(res).replace('0X','')

class JoiFileVersion(object):
    def __init__(self, filename):
        self.__filename = filename
        self.__version = "V0.00"
        self.__crc     = "FFFFFFFF"
        self.__subq    = "RP30000000"
        self.__date    = "2000-01-01 00:00:00"
        self.__getDevVersionInfo()

    def getFileName(self):
        print(self.__filename)

    def __getDevVersionInfo(self):
        filename = self.__filename
        info = {}
        if False == os.path.exists(filename) :
            print("config文件不存在!")
            return None
        with open(filename , 'rt') as f:
            line = " "
            data = " "
            ppcData = " "
            while line != '':
                line = f.readline()
                if line.startswith("[FILE VERSION"):
                    data = ''.join(re.findall(r'\[(.+?)\]', line))
                    continue
                
                if line.startswith("[PPC VERSION="):
                    ppcData = ''.join(re.findall(r'\[(.+?)\]', line))
                    continue

        data = data.split(" ")
        ppcData = ppcData.split(" ")

        info['version']  = data[1].split('=')[1]
        info['Subq']  = data[2].split('=')[1]
        info['date']  = data[3].split('=')[1] + " " + data[4].split('=')[1]
        info['crc']  = data[6].split('=')[1]
        info['ppcversion']  = ppcData[1].split('=')[1]
        info['ppcdate']  = ppcData[2].split('=')[1] + " " + ppcData[3].split('=')[1]
        info['ppccrc']  = ppcData[4].split('=')[1]
        self.__version = info['version'][:5]
        self.__subq    = info['Subq']
        self.__date    = info['date']
        self.__crc     = info['crc']
        self.__ppcVersion = info['ppcversion']
        self.__ppcDate    = info['ppcdate']
        self.__ppcCrc     = info['ppccrc']  

        return info

    def getVersion(self):
        return self.__version
    
    def getSubq(self):
        return self.__subq

    def getDate(self):
        return self.__date
    
    def getCrc(self):
        return self.__crc

    def getPPCVersion(self):
        return self.__ppcVersion

    def getPPCDate(self):
        return self.__ppcDate

    def getPPCCrc(self):
        return self.__ppcCrc    

#对整个表进行设置样式设计
def setStyle(sheet, rows, columns):
    #sheetnames = wb.get_sheet_names() #获得表单名字
    #CurrentSheet = wb.get_sheet_by_name(sheetnames[0])
    # 字体
    font = Font(name='宋体', size=12, b=False)

    # 边框
    line_t = Side(style='thin', color='000000')  # 细边框
    line_m = Side(style='thick', color='000000')  # 粗边框

    border = Border(top=line_m, bottom=line_m, left=line_m, right=line_m)

    # 填充,无
    fill = PatternFill('solid', fgColor='CFCFCF')

    # 对齐
    alignment = Alignment(horizontal='center', vertical='center')

    #打包样式
    sty = NamedStyle(name='sty', font=font, border=border, alignment=alignment, fill=fill)

    for r in range(3, rows+1):
        sheet.row_dimensions[rows].height = 45
        for c in range(1, columns):
            if rows < 3 :
                pass
            else:
                try:
                    sheet.cell(r, c).style = sty 
                except ValueError:
                    sheet.cell(r, c).style = 'sty' #Once registered assign the style using just the name:
                                               #ws['D5'].style = 'highlight'
                

'''
def style():
    ##赋值style为XFStyle()，初始化样式
    style = xlwt.XFStyle()
    #设置单元格内字体样式 
    font = xlwt.Font()
    font.name = '宋体'
    font.bold = False
    return style

def write_excel():
    wb = xlwt.Workbook()#创建工作�?
    sheet = wb.add_sheet(u'sheet1', cell_overwrite_ok=True)#创建第一个sheet�? 第二参数用于确认同一个cell单元是否可以重设�?
    #初始化表头列�?
    tb_head = [
    u'装置类别',
    u'装置型号',
    u'适用硬件归档�?',
    u'装置应用型号',
    u'程序包joi',
    u'ICD文件�?',
    u'ICD文件版本',
    u'ICD文件CRC32校验�?',
    u'ICD文件MD5校验�?',
    u'显示软件版本',
    u'显示生成日期',
    u'打包日期',
    u'管理序号'
    ]

    for i, item in enumerate(tb_head):
        sheet.write(0, i, item, style())

    return wb
'''

if __name__ == '__main__':
    bookName = u"NSR-3641A-DA-G-A0010备用电源自投装置配套的软件执行代码、软件版本及功能说明.xlsx"
    workbook = load_workbook(bookName)
    sheetnames = workbook.get_sheet_names() #获得表单名字
    CurrentSheet = workbook.get_sheet_by_name(sheetnames[0])
    config = JoiFileVersion('config.txt')
    CurrentSheet.insert_rows(3)
    setStyle(CurrentSheet, 3, column_index_from_string('AN')+1) #列号转换为数字

    CurrentSheet['A3'] = "备自投"
    CurrentSheet['B3'] = "NSR-3641"
    CurrentSheet['C3'] = "NSR-3641A-DA-G_A0010"
    CurrentSheet['D3'] = "NSR-3641A-DA-G"
    CurrentSheet['E3'] = " \n".join(searchFile("joi"))
    CurrentSheet['F3'] = " \n".join(searchFile("icd"))
    icdVersion = [getICDVersion(files) for files in searchFile("icd")] #多个文件创建时间需要合并打开?
    CurrentSheet['G3'] = " \n".join(icdVersion)
    icdCrc32 = [getCrc32(files) for files in searchFile("icd")] #多个文件创建时间需要合并打开?
    CurrentSheet['H3'] = " /\n".join(icdCrc32)
    icdMd5 = [getMd5(files) for files in searchFile("icd")] #多个文件创建时间需要合并打开?
    CurrentSheet['I3'] = " /\n".join(icdMd5)
    icdCreateTime = [getfileCreateTime(files) for files in searchFile("icd")] #多个文件创建时间需要合并打开?
    CurrentSheet['J3'] = " /\n".join(icdCreateTime)
    CurrentSheet['K3'] = config.getVersion()
    CurrentSheet['L3'] = config.getDate()
    joiCreateTime = [getfileCreateTime(files) for files in searchFile("joi")] #多个文件创建时间需要合并打开?
    CurrentSheet['M3'] = " /\n".join(joiCreateTime)
    CurrentSheet['N3'] = config.getSubq()
    CurrentSheet['O3'] = config.getCrc()

    CurrentSheet['Q3'] = config.getPPCVersion()
    CurrentSheet['R3'] = config.getPPCDate()
    CurrentSheet['S3'] = config.getPPCCrc()



    try:    
        workbook.save(bookName) 
        workbook.close()
    except Exception as err:
        print("文件保存异常 %s"%(err))
    print('Succeed!')
