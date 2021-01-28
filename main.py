#!/usr/bin/env python
# -*- coding:utf-8 -*-
import os
import json

import win32com
from win32com.client import Dispatch


# 处理Word文档的类

class RemoteWord:
    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Word.Application')  # 此处使用的是Dispatch，原文中使用的DispatchEx会报错
        self.xlApp.Visible = 0  # 后台运行，不显示
        self.xlApp.DisplayAlerts = 0  # 不警告
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.doc = self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()  # 创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc = self.xlApp.Documents.Add()
            self.filename = ''

    def add_doc_end(self, string):
        '''在文档末尾添加内容'''
        rangee = self.doc.Range()
        rangee.InsertAfter('\n' + string)

    def add_doc_start(self, string):
        '''在文档开头添加内容'''
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string + '\n')

    def insert_doc(self, insertPos, string):
        '''在文档insertPos位置添加内容'''
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n' + string)

    def replace_doc(self, string, new_string):
        '''替换文字'''
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        # (string--搜索文本,
        # True--区分大小写,
        # True--完全匹配的单词，并非单词中的部分（全字匹配）,
        # True--使用通配符,
        # True--同音,
        # True--查找单词的各种形式,
        # True--向文档尾部搜索,
        # 1,
        # True--带格式的文本,
        # new_string--替换文本,
        # 2--替换个数（全部替换）
        self.xlApp.Selection.Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)

    def replace_docs(self, string, new_string):
        '''采用通配符匹配替换'''
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        self.xlApp.Selection.Find.Execute(string, False, False, True, False, False, False, 1, False, new_string, 2)

    def save(self):
        '''保存文档'''
        self.doc.Save()

    def save_as(self, filename):
        '''文档另存为'''
        self.doc.SaveAs(filename)

    def close(self):
        '''保存文件、关闭文件'''
        self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()

class SysInit:
    def InitJsonConfig(self):
        if os.path.exists('.\\Config'):
            return
        os.mkdir(".\\Config")
        filename = '.\\Config\\userconfig.json'
        data = {'ReplaceCount':'1', '[A]':"YouContent"}
        with open(filename, 'w',encoding='utf-8') as file_obj:
            json.dump(data, file_obj, indent=4,ensure_ascii=False)
if __name__ == '__main__':
    init=SysInit()
    init.InitJsonConfig()
    #目前支持九个替换符
    list=['[A]','[B]','[C]','[D]','[E]','[F]','[G]','[H]','[I]']
    while True:
        print("输入文件夹根目录即可 如D:/File \n")
        FileName=input()
        #读取json文件
        f = open('.\\Config\\userconfig.json', 'r',encoding='utf-8')
        content = f.read()
        context = json.loads(content)
        f.close()
        ReplaceCount=int(context['ReplaceCount'])
        #读取文件列表
        filenames = os.listdir(FileName)
        filelistcount= len(filenames)
        countfile=1
        #按列处理
        print("开始处理")
        for now in filenames:
            doc=RemoteWord(FileName+"//"+now)
            if countfile<=filelistcount:
                print("正在处理第"+str(countfile)+"份文件"+"文件名: "+now)
                countfile=countfile+1
            for count in range(0,ReplaceCount):
                ReplaceContext=context[list[count]]
                doc.replace_doc(list[count],ReplaceContext)
            doc.close()
            del doc
        #提示语言
        print("处理完毕，请重新选择目录或者关闭本工具")