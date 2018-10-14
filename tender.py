#!/usr/bin/env python
#-*- coding:utf-8 -*-
# Example from: https://www.jianshu.com/p/1f60cdd9655
from docx          import Document
from docx.shared   import *
from docx.oxml.ns  import qn
from docx.enum.dml import MSO_THEME_COLOR

import gtk
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class PyApp(gtk.Window):
    def __init__(self):
        super(PyApp, self).__init__()
        self.set_default_size(300, 200)
        self.set_title("Tender Template Builder")
        self._label = gtk.Label("Hello World")
        self.add(self._label)
        self.show_all()
    def setText(self, text):
        self._label.set_text(text) 

def main():
    #打开文档
    document = Document(u'input.docx')
    paragraphs = document.paragraphs
    #print(len(paragraphs))

    #查看段落个数
    app = PyApp()
    app.setText(paragraphs[1].text)
    #启动窗口
    gtk.main()

    # 保存文件
    #document.save(u'demo.docx')
    #print('Saved OK!')

if __name__ == "__main__":
    main()

