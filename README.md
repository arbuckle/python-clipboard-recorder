python-clipboard-recorder
=========================

Records a permanent history of the Windows clipboard to disk.


Specifically:
-------------
Listens for Ctrl+C, Ctrl+X, and PrtScr keypresses, in addition to monitoring the clipboard for changes on a 2-second interval.

Exit program with Alt+Shift+X

Clipboard data is saved to c:/users/[username]/documents/MyClipboard

- Text content is grouped into daily files.
- Images are saved individually.
- Folder and content permissions are probably 0777.  This isn't explicitly set.
- Filesystem settings are the first declarations in the file if you're interested in changing this.
- Tested on Windows 7 with Python 2.7.

Dependencies:
-------------
- pywin32
- pyHook
- PIL
- pythoncom


MIT Licensed.
