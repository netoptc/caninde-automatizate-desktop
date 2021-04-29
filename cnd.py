import sys
sys.path.append("lib")

from lib import Faturamento
from lib import GUI
import wx

def main():
    fatura = Faturamento.fatura()
    app = wx.App()
    frame = GUI.my_window(None, title='CND', fatura=fatura)
    app.MainLoop()
    
if __name__ == "__main__":
    main()
