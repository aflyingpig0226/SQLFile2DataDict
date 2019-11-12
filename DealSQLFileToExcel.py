# -*- coding: utf-8 -*-
# @Time    : 2019/11/12 11:55
# @Author  : RUISHI3
# @FileName: DealSQLFileToExcel.py
# @Software: PyCharm
import os
from copy import deepcopy

import xlwt


# 处理原生SQL文件
def dealSQLFile(filePath):
    with open(filePath, "r", encoding="utf8") as f:
        lines = f.readlines()

        for line in lines:
            if line.startswith("  `"):
                listFiled.append(line.strip())

            if line.startswith(") ENGINE"):
                listFiled.append("$")

            if line.startswith("DROP TABLE IF EXISTS `"):
                listEnglishName.append(line[21: len(line) - 1])
            elif line.startswith(") ENGINE=InnoDB"):
                tableFields = line.split(" ")
                listChineseName.append(
                    tableFields[len(tableFields) - 1].replace("COMMENT='", "").replace("'", "").replace(";\n", ""))

    i = 0
    while i < len(listEnglishName):
        for words in listFiled:
            if words != "$":
                fields = words.split(" ")
                res1 = fields[0].replace("`", "")
                res2 = fields[1].replace("`", "")
                res3 = fields[len(fields) - 1].replace("'", "").replace(",", "")
                resFields = res1 + "\t" + res2 + "\t" + res3
                tempList.append(resFields)
            else:
                resTable = "表名：" + listEnglishName[i].strip().replace("`", "").replace(";", "") + " " + listChineseName[
                    i]
                tempList.append(resTable)
                i += 1
    return tempList


# 处理好的SQL生成临时文件
def toTempFile(filePath):
    sqlRes = dealSQLFile(filePath)
    with open(tempFilePath, "w", encoding="utf8") as f:
        for sql in sqlRes:
            f.write(sql)
            f.write("\n")
    return True


# 将数据写入到excel中
def toExcel():
    res = toTempFile(filePath)
    if res:
        with open(tempFilePath, "r", encoding="utf8") as f:
            lines = f.readlines()
            for line in lines:
                if line.startswith("表名："):
                    sheet = line.split(" ")[0].split("：")[1]
                    temp = deepcopy(fieldList)
                    dic[sheet] = temp
                    fieldList.clear()
                else:
                    fieldList.append(line)

        f = xlwt.Workbook()
        headerStyle = xlwt.easyxf("pattern: pattern solid, fore_colour ice_blue")

        wk = f.add_sheet("表清单", cell_overwrite_ok=True)
        wk.write(0, 0, "表英文名", headerStyle)
        wk.write(0, 1, "表中文名", headerStyle)
        wk.write(0, 2, "链接", headerStyle)

        for index in range(len(listEnglishName)):
            tableName = listEnglishName[index].replace("`", "").replace(";", "")
            wk.write(index + 1, 0, tableName)
            if len(tableName) >= 31:
                tempTable = tableName[0:31]
                wk.write(index + 1, 2, xlwt.Formula('HYPERLINK("#{}!A1";"{}!A1")'.format(tempTable, tempTable)))
            else:
                wk.write(index + 1, 2, xlwt.Formula('HYPERLINK("#{}!A1";"{}!A1")'.format(tableName, tableName)))

        for index in range(len(listEnglishName)):
            wk.write(index + 1, 1, listChineseName[index].replace("`", "").replace(";", ""))

        for keys in dic.keys():
            if len(keys) >= 31:
                tempKey = keys[0:31]
                sheetName = f.add_sheet(tempKey, cell_overwrite_ok=True)
                sheetName.write(0, 0, "数据项", headerStyle)
                sheetName.write(0, 1, "类型", headerStyle)
                sheetName.write(0, 2, "名称", headerStyle)
                sheetName.write(0, 3, xlwt.Formula('HYPERLINK("#{}!A1";"返回")'.format("表清单", "表清单")), headerStyle)
            else:
                sheetName = f.add_sheet(keys, cell_overwrite_ok=True)
                sheetName.write(0, 0, "数据项", headerStyle)
                sheetName.write(0, 1, "类型", headerStyle)
                sheetName.write(0, 2, "名称", headerStyle)
                sheetName.write(0, 3, xlwt.Formula('HYPERLINK("#{}!A1";"返回")'.format("表清单", "表清单")), headerStyle)
            for i in range(len(dic[keys])):
                for j in range(len(dic[keys][i].split("\t"))):
                    sheetName.write(i + 1, j, dic[keys][i].split("\t")[j])
            f.save(resFile)
        return True
    else:
        return False


if __name__ == '__main__':
    listFiled = []
    listEnglishName = []
    listChineseName = []
    tempList = []

    sheetList = []
    fieldList = []
    dic = {}

    filePath = r"C:\Users\RUIBABA\Desktop\关键业务表.sql"  # 原生SQL文件路径
    tempFilePath = r"C:\Users\RUIBABA\Desktop\temp.sql"  # 临时文件路径
    resFile = r"C:\Users\RUIBABA\Desktop\result.xls"  # 生成了的excel文件路径
    try:
        if toExcel():
            print("成功了！！！")
            os.remove(tempFilePath)
    except Exception as e:
        print(e)
