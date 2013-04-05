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
    # Change this to reflect the path to the folder where you'd like to store your clipboard
    'clipboard_path': 'c:/documents and settings/[USERNAME]/my documents/python/clipboard/',
    
    # Files are saved with date string filenames, to prevent overwriting.  http://docs.python.org/2/library/datetime.html#strftime-and-strptime-behavior
    'file_fmt_img': "%Y-%m-%d-%H%M%S.png",
    'file_fmt_txt': "%Y-%m-%d.txt",
    'log_separator': "\n\n===========================\n====%Y-%m-%d %H:%M:%S====\n===========================\n"
}


class Clipboard(object):
    """
    Clipboard provides a means of accessing the clipboard via its getClipboardData() function, which returns
    the clipboard data and a datatype of either ("img", None) signifiying whether the clipboard data is an 
    img or text.
    
    The previous clipboard values are accessible via 'self.clipboard' and 'self.datatype'
    """
    def __init__(self, options):
        self.clipboard = ''
        self.datatype = None

    def getClipboardData(self):
        cb.OpenClipboard()
        fmt = cb.EnumClipboardFormats()
        cb.CloseClipboard()

        if fmt in [1, 13, 16, 7, 49224, 49327, 49322, 49158, 49459, 49471]: #text, unicode-text, locale, oem text, ShellIDList Array, Preferred DropEffect, Shell Object Offsets, Filename, FileNameW, Ole Private Data
            cb.OpenClipboard()
            self.clipboard, self.datatype = cb.GetClipboardData(cb.CF_UNICODETEXT), None
            cb.CloseClipboard()
        elif fmt in [2, 8, 17, 5, 49364]: #images
            self.clipboard, self.datatype = ImageGrab.grabclipboard(), 'img'
        return self.clipboard, self.datatype


class FileSystem(object):
    """
    FileSystem provides a means of saving clipboard data to disk.
    """
    def __init__(self,options):
        self.clipboard_path = options.get('clipboard_path', None)
        self.file_fmt_img = options.get('file_fmt_img', None)
        self.file_fmt_txt = options.get('file_fmt_txt', None)
        self.log_separator = options.get('log_separator', None)

    def saveClipboardData(self, clipboard, datatype):
        if datatype == 'img':
            clipboard.save(self._getSavePath(datatype), 'PNG')
        else:
            cb_text = open(self._getSavePath(datatype), 'a')
            cb_text.write(datetime.now().strftime(self.log_separator))
            cb_text.write('%s\n' % clipboard)
            cb_text.close()

    def _getSavePath(self, datatype):
        if datatype == 'img':
            return self.clipboard_path + datetime.now().strftime(self.file_fmt_img)
        else:
            return self.clipboard_path + datetime.now().strftime(self.file_fmt_txt)


class Handlers(object):
    """
    Handlers for keyboard-initiated clipboard change events, as well as incidental (mouse-driven) clipboard
    changes.
    """
    def __init__(self, clipboard, filesystem):
        self.clipboard = clipboard
        self.save = filesystem.saveClipboardData
    
    def handleClipboardChanged(self, sleep=True):
        if sleep:
            time.sleep(0.2)
        self.save(*self.clipboard.getClipboardData())

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
            last = self.clipboard.clipboard
            clipboard, datatype = self.clipboard.getClipboardData()
            
            # detecting when current type is image and determining if current/last are identical
            # PIL Image objects do not have an equals() method, so I compare the image strings instead.
            fresh_image = True
            if datatype == 'img' and hasattr(last, 'getextrema') and hasattr(clipboard, 'getextrema'):
                fresh_image = not last.tostring() == clipboard.tostring()

            if last <> clipboard and fresh_image:
                self.handleClipboardChanged(False)
            time.sleep(2)


def main():
    filesystem = FileSystem(options)
    clipboard = Clipboard(options)
    handlers = Handlers(clipboard, filesystem)
    
    thread = Thread(target=handlers.clipboardChangedListener)
    thread.daemon = True
    thread.start()

    hm = HookManager()
    hm.KeyDown = handlers.handleKeypress
    hm.HookKeyboard()
    PumpMessages()
    
if __name__ == "__main__":
    main()
