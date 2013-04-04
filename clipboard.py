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
from os import mkdir
from os.path import expanduser, exists
from PIL import ImageGrab
from threading import Thread

user_folder = expanduser('~')
base_dir = 'documents'
app_dir = 'MyClipboard'

file_fmt_img = "%Y-%m-%d-%H%M%S.png"
file_fmt_txt = "%Y-%m-%d.txt"
cb_separator = "\n\n===========================\n====%Y-%m-%d %H:%M:%S====\n===========================\n"


# global object used to prevent conflict between keyboard and change handlers.  
previous = ''

#primary change handler for the application.
def handleClipboardChanged(sleep=True):
    data, cb_type = getClipboardData(sleep)
    saveClipboardData(data, cb_type)

#getClipboardData reads the clipboard and returns the data and the cb_type (img or None)
def getClipboardData(sleep=True):
    global previous
    if sleep:
        time.sleep(0.2) # Required to ensure the keylogger captures the existing clipboard data rather than the prior clipboard data. Race condition!

    cb.OpenClipboard()
    fmt = cb.EnumClipboardFormats()
    cb.CloseClipboard()
    #print fmt
    data, cb_type = '', None
    if fmt in [1, 13, 16, 7, 49224, 49327, 49322, 49158, 49459, 49471]: #text, unicode-text, locale, oem text, ShellIDList Array, Preferred DropEffect, Shell Object Offsets, Filename, FileNameW, Ole Private Data
        cb.OpenClipboard()
        data, cb_type = cb.GetClipboardData(cb.CF_UNICODETEXT), None
        cb.CloseClipboard()
    elif fmt in [2, 8, 17, 5, 49364]: #images
        data, cb_type = ImageGrab.grabclipboard(), 'img'
        #print data
    
    previous = data
    return data, cb_type

#saveClipboardData accepts clipboard data and an optional clipboard cb_type, either "img" or None
def saveClipboardData(data, cb_type=None):
    if cb_type == 'img':
        #print 'saving'
        data.save(getSavePath(cb_type), 'PNG')
        #print 'img saved'
    else:
        cb_text = open(getSavePath(), 'a')
        cb_text.write(datetime.now().strftime(cb_separator))
        cb_text.write('%s\n' % data)
        cb_text.close()

#getSavePath returns the filename for saving clipboard data.  optional argument 'cb_type' is passed when clipboard content is an image.
def getSavePath(cb_type=None):
    if cb_type == 'img':
        return getSaveFolder() + datetime.now().strftime(file_fmt_img)
    else:
        return getSaveFolder() + datetime.now().strftime(file_fmt_txt)

#getSaveFolder returns the path of the folder in which to save files.    
def getSaveFolder():
    return "%s/%s/%s/" % (user_folder.replace('\\', '/'), base_dir, app_dir)

#makeAppFolder creates the app_dir if it does not already exist.  Win7 only.
def makeAppFolder():
    if exists(user_folder) and exists('%s\\%s' % (user_folder, base_dir)):
        if not exists('%s\\%s\\%s' % (user_folder, base_dir, app_dir)):
            mkdir('%s\\%s\\%s' % (user_folder, base_dir, app_dir))

#handleKeypress listens to all windows keypresses and when Ctrl+C or Ctrl+x is intercepted, spawns a new thread to capture the Clipboard data.
def handleKeypress(event):
    if event.Ascii == 3 or event.Ascii == 24 or event.Key == 'Snapshot':
        thread = Thread(target=handleClipboardChanged)
        thread.start()
    elif event.Ascii == 88 and event.Alt == 32: # quit on alt+shift+x
        sys.exit()
    return True

#clipboardChangedListener checks every two seconds to see whether the clipboard data has been modified.
def clipboardChangedListener():
    global previous
    while 1:
        last = previous
        current, cb_type = getClipboardData(False)
        
        # detecting when current type is image and determining if current/last are identical
        # PIL Image objects do not have an equals() method, so I compare the image strings instead.
        fresh_image = True
        if cb_type == 'img' and hasattr(last, 'getextrema') and hasattr(current, 'getextrema'):
            fresh_image = not last.tostring() == current.tostring()

        if last <> current and fresh_image:
            previous = current
            handleClipboardChanged(False)
        time.sleep(2)


def main():
    makeAppFolder()

    # You could probably get rid of 90% of this program and just call clipboardChangedListener.
    # This was added as an afterthought after I realized that mouse-initiated clipboard changes were ignored by keypress handlers.
    thread = Thread(target=clipboardChangedListener)
    thread.daemon = True
    thread.start()

    hm = HookManager()
    hm.KeyDown = handleKeypress
    hm.HookKeyboard()
    PumpMessages()
    
if __name__ == "__main__":
    main()
