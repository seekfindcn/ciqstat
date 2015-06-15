#---------------------------------------------------------------
#    程序：CIQ2K导出文本文件字段获取
#    程序名：getlist.py
#    版本：0.1.1
#    作者：懒秀才
#    日期：2014-05-28
#    语言：Python 3
#    操作：参数1为输入文件名，参数2为输出文件名
#    功能：获取ciq2000统计查询导出的文本文件数据中字段名称及编号
#---------------------------------------------------------------

#! /usr/bin/env python
#coding=utf-8

import os
import sys

def main():
	try:
		inTxt = sys.argv[1]
	except IndexError:
		inTxt = input("请输入获取字段txt文件名（需带有.txt）：")
	try:
		outTxt = sys.argv[2]
	except IndexError:
		outTxt = input("请输入字段名保存文件（需带有.txt）：")
	data = open(inTxt)
	decl = data.readline()#获取字段行
	decl = decl.strip('\n')#剔除行末回车
	decls = decl.split('\t')#按字段分割
	data.close()
	data = open(outTxt,'w')
	count = 0
	while count < len(decls):
		print (str(count+1)+'\t'+decls[count],file=data)
		count = count + 1
	data.close()
	
if __name__=="__main__":
		main()