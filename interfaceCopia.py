# -*- coding: utf-8 -*-
import openpyxl
import wx
from openpyxl import Workbook
import datetime
import os
import leitura

# -*- coding: utf-8 -*-
class Main (wx.Frame):
        def __init__(self, title="Apontamentos"):
            wx.Frame.__init__(self, None, title=title)

            panel = wx.Panel(self)
            
            meutexto = wx.StaticText(self, label=u"PECVAL", pos=(45, 30))
            meutexto.SetForegroundColour("white")
    
            # # btnOP = wx.Button(self, label= 'OP', pos=(0,50))
            # btnPerda = wx.Button(self, label = 'PERDA', pos=(240,50))

            btnAP = wx.Button(self, label = 'Copiar', pos=(55,55))

            self.Bind( wx.EVT_BUTTON, self.eventBtnAP, btnAP)
            # self.Bind( wx.EVT_BUTTON, self.eventBtnOP, btnOP)
            
            self.SetBackgroundColour("#130f40")
            self.SetTitle('PECVAL')
            self.SetSize((200,170))
            self.Centre()
            self.Show(True)
            # self.btnOP.Bind(wx.EVT_BUTTON, self.OnClicked) 

            # def OnClicked(self, event): 
            #     btn = event.GetEventObject().GetLabel() 
            #     print "Label of pressed button = ",btnOP  


        def eventBtnAP(self, event): 
                leitura.leitura()
                wx.MessageBox('Tudo Pronto!', 'Info', wx.OK | wx.ICON_INFORMATION)


                
            
        
        



if __name__ == "__main__":
   ex = wx.App()
   Main()
   ex.MainLoop()


