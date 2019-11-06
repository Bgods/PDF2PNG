# -*- coding:utf-8 -*-

import fitz
import glob
import os
from tkinter import *
import easygui
from tkinter import messagebox
import threading
import shutil

info = {
    'Version': 'v1.1',
    'Author': 'Bgods',
    'UpdateTime': '2019-10-15',
}

# GUI窗口类
class UI():
    # 定义GUI界面
    def __init__(self, root, Thread1, Thread2, Thread3, is_quit, about):
        root.protocol("WM_DELETE_WINDOW", is_quit)  # 关闭窗口
        root.title(u"PDF和IMG互转({} by {})".format(info['Version'], info['Author']))
        root.geometry("400x150")  # geometry()  定义窗口的大小

        self.label = Label(root, text='滑动滚动条\n设置像素大小:')
        self.label.grid(row=0, column=2, sticky=W + E, columnspan=2, rowspan=2)  #

        # scale滚动条，数值从10到40，水平滑动
        self.scale = Scale(root, from_=2.0, to=5.0, resolution=1, orient=HORIZONTAL) # 步长1，范围(2.0-5.0)
        self.scale.set(3)  # 设置初始值
        self.scale.grid(row=0, column=4, rowspan=2, columnspan=3, sticky=W + E) # 滑块按钮，设置像素大小

        # 创建查询按钮
        self.button1 = Button(root, text=u"PDF转IMG", command=Thread1)
        self.button2 = Button(root, text=u"IMG转PDF", command=Thread2)
        self.button3 = Button(root, text=u"PDF转图片型PDF", command=Thread3)

        Label(height=4, width=3).grid(row=2, column=0, columnspan=2, sticky=W + E)  # 占位标签
        self.button1.grid(row=2, column=2, sticky=W + E)  # 左右拉伸对齐
        Label(height=1, width=3).grid(row=2, column=3)  # 占位标签
        self.button2.grid(row=2, column=4, sticky=W + E)  # 左右拉伸对齐
        Label(height=1, width=3).grid(row=2, column=5)  # 占位标签
        self.button3.grid(row=2, column=6, sticky=W + E)  # 左右拉伸对齐

        # 菜单
        self.menubar = Menu(root)
        self.menubar.add_command(label=u"关于", command=about)
        root.config(menu=self.menubar)  # 显示菜单

# 功能实现类
class Func():
    def __init__(self, root):
        self.is_start = False

        self.gui = UI(
            root,
            self.Thread1,
            self.Thread2,
            self.Thread3,
            self.is_quit,
            self.about
        )  # 将我们定义的GUI类赋给服务类的属性，将执行的功能函数作为参数传入

    def is_quit(self):  # 是否退出
        import os
        if self.is_start:
            if messagebox.askokcancel(u"警告!", u"程序运行中，强制关闭会导致数据丢失，是否继续关闭!"):
                os._exit(0)  # 强制退出
        else:
            if messagebox.askokcancel(u"关闭", u"是否关闭窗口?"):
                os._exit(0)  # 强制退出

    def about(self):
        text = f'''
        名称: PDF和IMG互转
        作者: {info['Author']}
        版本: {info['Version']}
        更新时间: {info['UpdateTime']}
        说明: 
          1.滑块值设置越大，转换的像素越大，处理时间越长。
        '''
        messagebox.showinfo(u'关于', text)

    # 实现pdf转img功能
    def pdf2img(self):
        self.pdf_name = easygui.fileopenbox()  # 打开pdf文件
        if self.pdf_name:
            if self.pdf_name[-4:].lower() == '.pdf':
                self.is_start = True  # 线程启动，运行状态变为True
                self.gui.button1.config(text=u'运行中...')
                dir_name, base_name = get_dir_name(self.pdf_name)
                path = self.mkdir(dir_name, base_name[:-4])
                os.makedirs(path)

                pdf = fitz.open(self.pdf_name)
                for pg in range(0, pdf.pageCount):
                    page = pdf[pg]  # 获得每一页的对象
                    zoom = self.gui.scale.get()  # 值越大，生成图像像素越高
                    trans = fitz.Matrix(zoom, zoom).preRotate(0)
                    pm = page.getPixmap(matrix=trans, alpha=False)  # 获得每一页的流对象
                    pm.writePNG(path + os.sep + base_name[:-4] + '_' + '{:0>3d}.{}'.format(pg+1, 'png'))  # 保存图片
                pdf.close()

                self.is_start = False  # 线程结束，运行状态变为False
                self.file_dir = None  # 初始化图片文件夹地址
                self.gui.button1.config(text=u'PDF转IMG')
                messagebox.showinfo('信息', '转化成功！')
            else:
                messagebox.showwarning('警告!', '文件类型错误,请打开pdf文件类型!')

    # 实现img转pdf功能
    def img2pdf(self):
        self.file_dir = easygui.diropenbox()  # 打开img文件所在路径
        if self.file_dir:
            self.is_start = True  # 线程启动，运行状态变为True
            self.gui.button2.config(text=u'运行中...')
            img = self.file_dir + "/*"  # 获得文件夹下的所有对象
            dir_name, base_name = get_dir_name(self.file_dir)
            doc = fitz.open()
            for img in sorted(glob.glob(img)):  # 排序获得对象
                img_doc = fitz.open(img)  # 获得图片对象
                pdf_bytes = img_doc.convertToPDF()  # 获得图片流对象
                img_pdf = fitz.open("pdf", pdf_bytes)  # 将图片流创建单个的PDF文件
                doc.insertPDF(img_pdf)  # 将单个文件插入到文档
                img_doc.close()
                img_pdf.close()
            filename = self.mkdir(dir_name, base_name, '.pdf')
            doc.save(filename)  # 保存文档
            doc.close()

            self.gui.button2.config(text=u'IMG转PDF')
            self.is_start = False  # 线程结束，运行状态变为False
            self.pdf_name = None  # 初始化地址
            messagebox.showinfo('信息', '转化成功！')

    def pdf2pdf(self):
        self.pdf_name = easygui.fileopenbox()  # 打开pdf文件
        if self.pdf_name:
            if self.pdf_name[-4:].lower() == '.pdf':
                self.is_start = True  # 线程启动，运行状态变为True
                self.gui.button3.config(text=u'运行中...')

                # 1、创建临时路径path
                dir_name, base_name = get_dir_name(self.pdf_name)
                path = self.mkdir(dir_name, base_name[:-4])
                os.makedirs(path)

                try:
                    # 2、将pdf转成图片，保存到path路径中
                    pdf = fitz.open(self.pdf_name)
                    for pg in range(0, pdf.pageCount):
                        page = pdf[pg]  # 获得每一页的对象
                        zoom = self.gui.scale.get()  # 值越大，生成图像像素越高
                        trans = fitz.Matrix(zoom, zoom).preRotate(0)
                        pm = page.getPixmap(matrix=trans, alpha=False)  # 获得每一页的流对象
                        pm.writePNG(path + os.sep + base_name[:-4] + '_' + '{:0>3d}.png'.format(pg + 1))  # 保存图片
                    pdf.close()

                    # 3、将path路径中图片合并成pdf
                    img = path + "/*"  # 获得文件夹下的所有对象
                    dir_name, base_name = get_dir_name(path)
                    doc = fitz.open()
                    for img in sorted(glob.glob(img)):  # 排序获得对象
                        img_doc = fitz.open(img)  # 获得图片对象
                        pdf_bytes = img_doc.convertToPDF()  # 获得图片流对象
                        img_pdf = fitz.open("pdf", pdf_bytes)  # 将图片流创建单个的PDF文件
                        doc.insertPDF(img_pdf)  # 将单个文件插入到文档
                        img_doc.close()
                        img_pdf.close()
                    filename = self.mkdir(dir_name, base_name, '.pdf')
                    doc.save(filename)  # 保存文档
                    doc.close()

                    # 4、删除临时路径及其下的图片
                    shutil.rmtree(path)
                    self.gui.button3.config(text=u'PDF转图片型PDF')
                    self.is_start = False  # 线程结束，运行状态变为False
                    self.file_dir = None  # 初始化图片文件夹地址
                    messagebox.showinfo('信息', '转化成功！')

                except BaseException as e:
                    messagebox.showinfo('错误信息', f'{e}')

            else:
                messagebox.showwarning('警告!', '文件类型错误,请打开pdf文件类型!')

    def mkdir(self, path, filename, type=''):
        path_file = path + os.sep + filename + type
        if os.path.exists(path_file):
            return self.mkdir(path, filename + '_副本', type)
        else:
            return path_file

    # 为pdf2img单独开一个线程
    def Thread1(self):
        if self.is_start:
            messagebox.showwarning(u'警告!', u'程序运行中,请结束后再打开!')
        else:
            self.Thread = threading.Thread(target=self.pdf2img)
            self.Thread.start()

    # 为img2pdf单独开一个线程
    def Thread2(self):
        if self.is_start:
            messagebox.showwarning(u'警告!', u'程序运行中,请结束后再打开!')
        else:
            self.Thread = threading.Thread(target=self.img2pdf)
            self.Thread.start()

    # 为pdf2pdf单独开一个线程
    def Thread3(self):
        if self.is_start:
            messagebox.showwarning(u'警告!', u'程序运行中,请结束后再打开!')
        else:
            self.Thread = threading.Thread(target=self.pdf2pdf)
            self.Thread.start()


def get_dir_name(file_dir):
    base_name = os.path.basename(file_dir)  # 获得地址的文件名
    dir_name = os.path.dirname(file_dir)  # 获得地址的父链接
    return dir_name, base_name



if __name__ == '__main__':
    root = Tk()
    Func(root)
    root.mainloop()


