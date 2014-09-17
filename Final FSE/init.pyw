import wx
import PyWord

if __name__ == "__main__":
    app = wx.App(0)
    try:
        frame = PyWord.PyWindow(None, -1, "PyWord", size=(1024,768))
        app.MainLoop()
    finally:
        del app
