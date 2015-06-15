#---------------------------------------------------------------
#    程序：CIQ编码整理
#    程序名：getciqcode.py
#    版本：0.0.1
#    作者：懒秀才
#    日期：2014-07-14
#    语言：Python 3
#    操作：参数为导入
#    功能：整理ciq2000统计查询导出的CIQ分类编码文本文件数据
#---------------------------------------------------------------

#! /usr/bin/env python
#coding=utf-8

import os
import sys

def main():
#导入CIQ分类列表
	try:
		inTxt = sys.argv[1]
	except IndexError:
		inTxt = "CIQ编码分类.txt"
	try:
		outTxt = sys.argv[2]
	except IndexError:
		outTxt = "CIQ编码分类整理.txt"
	data = open(inTxt)
	ciqCodedata = data.readline()
	ciqCodedata = ciqCodedata.strip('\n')
	ciqCodeLists = ciqCodedata.split('\t')
	data.close()
	
	data = open(outTxt,'w')
	print ('商品统计分类代码'+'\t'+'商品统计分类名称'+'\t'+'CIQ统计编码分类',file=data)
	
	for ciqCodelist in ciqCodelists:
		listdata = open(ciqCodelist[0])
		ciqCodedata = listdata.readline()
		ciqCodedata = ciqCodedata.strip('\n')
		ciqlist = ciqCodedata.split('\t')
		data.close()
		count = 0
		while count < len(ciqlist):
			print (ciqlist[count+1][0]+'\t'+ciqlist[count+1][1]+'\t'+ciqCodelist[1],file=data)
			count = count + 1
	data.close()
	
if __name__=="__main__":
	main()