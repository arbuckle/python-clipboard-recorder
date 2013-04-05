"""
Copyright (C) 2013 David Arbuckle

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"""
import win32clipboard as cb
import sys
import time

from datetime import datetime
from pyHook import HookManager
from pythoncom import PumpMessages
from PIL import ImageGrab
from threading import Thread

options = {
    'clipboard_path': 'c:/documents and settings/davida.wp/my documents/python/clipboard/',
    'file_fmt_img': "%Y-%m-%d-%H%M%S.png",
    'file_fmt_txt': "%Y-%m-%d.txt",
    'log_separator': "\n\n===========================\n====%Y-%m-%d %H:%M:%S====\n===========================\n"
}


class Clipboard(object):
    def __init__(self, options):
        self.clipboard_path = options.get('clipboard_path', None)
        self.file_fmt_img = options.get('file_fmt_img', None)
        self.file_fmt_txt = options.get('file_fmt_txt', None)
        self.log_separator = options.get('log_separator', None)
        
        self.cb_previous = ''
        self.cb_current = ''
        self.cb_type = None
        
    def getClipboardData(self):
        self.cb_previous = self.cb_current

        cb.OpenClipboard()
        fmt = cb.EnumClipboardFormats()
        cb.CloseClipboard()

        if fmt in [1, 13, 16, 7, 49224, 49327, 49322, 49158, 49459, 49471]: #text, unicode-text, locale, oem text, ShellIDList Array, Preferred DropEffect, Shell Object Offsets, Filename, FileNameW, Ole Private Data
            cb.OpenClipboard()
            self.cb_current, self.cb_type = cb.GetClipboardData(cb.CF_UNICODETEXT), None
            cb.CloseClipboard()
        elif fmt in [2, 8, 17, 5, 49364]: #images
            self.cb_current, self.cb_type = ImageGrab.grabclipboard(), 'img'

    def saveClipboardData(self):
        if self.cb_type == 'img':
            self.cb_current.save(self.getSavePath(), 'PNG')
        else:
            cb_text = open(self.getSavePath(), 'a')
            cb_text.write(datetime.now().strftime(self.log_separator))
            cb_text.write('%s\n' % self.cb_current)
            cb_text.close()

    def getSavePath(self):
        if self.cb_type == 'img':
            return self.clipboard_path + datetime.now().strftime(self.file_fmt_img)
        else:
            return self.clipboard_path + datetime.now().strftime(self.file_fmt_txt)


class Handlers(object):
    def __init__(self, clipboard):
        self.clipboard = clipboard
    
    def handleClipboardChanged(self, sleep=True):
        if sleep:
            time.sleep(0.2)
        self.clipboard.getClipboardData()
        self.clipboard.saveClipboardData()

    def handleKeypress(self, event):
        # Ctrl+C, Ctrl+X, PrtScr
        if event.Ascii == 3 or event.Ascii == 24 or event.Key == 'Snapshot':
            thread = Thread(target=self.handleClipboardChanged)
            thread.start()
        # Exit on Alt+Shift+X
        elif event.Ascii == 88 and event.Alt == 32: 
            sys.exit()
        return True

    def clipboardChangedListener(self):
        while 1:
            last = self.clipboard.cb_current
            self.clipboard.getClipboardData()
            
            # detecting when current type is image and determining if current/last are identical
            # PIL Image objects do not have an equals() method, so I compare the image strings instead.
            fresh_image = True
            if self.clipboard.cb_type == 'img' and hasattr(last, 'getextrema') and hasattr(self.clipboard.cb_current, 'getextrema'):
                fresh_image = not last.tostring() == self.clipboard.cb_current.tostring()

            if last <> self.clipboard.cb_current and fresh_image:
                self.handleClipboardChanged(False)
            time.sleep(2)


def main():
    clipboard = Clipboard(options)
    handlers = Handlers(clipboard)
    
    thread = Thread(target=handlers.clipboardChangedListener)
    thread.daemon = True
    thread.start()

    hm = HookManager()
    hm.KeyDown = handlers.handleKeypress
    hm.HookKeyboard()
    PumpMessages()
    
if __name__ == "__main__":
    main()
