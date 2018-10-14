#!/usr/bin/env python
#-*- coding:utf-8 -*-
#Reference link: https://blog.csdn.net/hankhanti/article/details/6015488

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

try:
    import gtk
    import pygtk
    pygtk.require("2.0")
    import gtk.glade
except:
    sys.exit(1)

class HelloGTK:
    """This is an Hello GTK application"""
    def __init__(self):
        # Set the glade file
        self.gladefile = "hello.glade"
        self.wtree = gtk.glade.XML(self.gladefile)
        self.window = self.wtree.get_widget("window_tender")
        dic = {
            "on_button_test_clicked"   : self.on_button_test_clicked,
            "on_button_open_clicked"   : self.on_button_open_clicked,
            "on_window_tender_destroy" : gtk.main_quit 
        }
        self.wtree.signal_autoconnect(dic)
        if(self.window):
            self.window.show()
    def on_button_test_clicked(self, widget):
        print("\nCalled => on_button_test_clicked")
    def on_button_open_clicked(self, widget):
        print("\nCalled => on_button_open_clicked")


if __name__ == "__main__":
    hwg = HelloGTK()
    gtk.main()
