#!/usr/bin/python
# -*- coding: UTF-8 -*-

import tkinter as tk

from xmind.XmindGUI import XMIND_GUI

if __name__ == '__main__':
    win = tk.Tk()
    xmind_gui = XMIND_GUI(win)
    xmind_gui.set_init_window()
    win.mainloop()
