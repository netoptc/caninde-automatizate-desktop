import wx
import os
import subprocess

class my_window(wx.Frame):
    def __init__(self, parent, title, fatura):
        super(my_window, self).__init__(parent, title=title,size=(480 , 320))
        self.fatura = fatura
        self.SetBackgroundColour('white')
        self.Centre()
        self.InitUI()
        self.Show()

    def InitUI(self):
        self.basePath = os.getcwd()
        self.panel = wx.Panel(self, -1)   
        my_sizer = wx.BoxSizer(wx.VERTICAL) 
        fgs = wx.FlexGridSizer(2, 4, 20)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)

        self.banner = wx.StaticBitmap(self.panel, wx.ID_ANY, wx.Bitmap(self.basePath+"/img/cndImage.jpg", wx.BITMAP_TYPE_ANY))
        
        dom = wx.StaticText(self.panel, label="Dominio")
        cpf = wx.StaticText(self.panel, label="CPF")
        user = wx.StaticText(self.panel, label="User")
        password = wx.StaticText(self.panel, label="Password")
        epty = wx.StaticText(self.panel, label="")

        self.text_ctrl_dominio = wx.TextCtrl(self.panel)
        self.text_ctrl_cpf = wx.TextCtrl(self.panel)
        self.text_ctrl_user = wx.TextCtrl(self.panel)
        self.text_ctrl_password = wx.TextCtrl(self.panel)
       
        send_btn = wx.Button(self.panel, label=' ► ', size=(20,20))
        send_btn.Bind(wx.EVT_BUTTON, self.on_press)
      
        hbox1.Add(self.banner, flag=wx.CENTER)
        fgs.AddMany([(dom),(self.text_ctrl_dominio),(cpf),(self.text_ctrl_cpf),(user),(self.text_ctrl_user),(password),(self.text_ctrl_password),(epty),(send_btn)])
       
        my_sizer.Add(hbox1, proportion=1, flag=wx.ALL|wx.CENTER, border=1)
        my_sizer.Add(fgs, proportion=2, flag=wx.ALL|wx.CENTER, border=1)
     
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(self.basePath+"/img/icon.jpg", wx.BITMAP_TYPE_ANY))
        self.SetIcon(icon)
        
        self.panel.SetSizer(my_sizer)

       

    def secondPanel(self):

        panel = wx.Panel(self)
    
        sizer = wx.GridBagSizer(4, 5)

        text1 = wx.StaticText(panel, label="Conferir Fatura")
        line = wx.StaticLine(panel)
        text2 = wx.StaticText(panel, label="Fatura")
        self.tc1 = wx.FilePickerCtrl(panel)
        text3 = wx.StaticText(panel, label='Código de Barras:   ')
        self.tc2 = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)

        self.tc1.Bind(wx.EVT_FILEPICKER_CHANGED, self.on_textFatura)
        self.tc2.Bind(wx.EVT_TEXT_ENTER, self.on_text)
        
        buttonSave =wx.Button(panel, wx.ID_SAVE)
        buttonSave.SetBackgroundColour(wx.Colour(80, 200, 80))
        buttonSave.Bind(wx.EVT_BUTTON, self.on_save)

        sizer.Add(text1, pos=(0, 1), flag=wx.CENTER|wx.TOP|wx.BOTTOM,border=15)
        sizer.Add(line, pos=(1, 0), span=(1, 5),flag=wx.EXPAND|wx.BOTTOM, border=10)
        sizer.Add(text2, pos=(2, 0), flag=wx.LEFT, border=10)
        sizer.Add(self.tc1, pos=(2, 1), span=(1, 3), flag=wx.TOP|wx.EXPAND)
        sizer.Add(text3, pos=(3, 0), flag=wx.LEFT|wx.TOP, border=10)
        sizer.Add(self.tc2, pos=(3, 1), span=(1, 3), flag=wx.TOP|wx.EXPAND,border=5)
        sizer.Add(buttonSave, pos=(4, 3), flag=wx.RIGHT|wx.TOP)

        sizer.AddGrowableCol(2)

        panel.SetSizer(sizer)
    
        self.Layout()

    def on_text(self, event):

        pathFatura = self.tc1.GetPath()
        if not (pathFatura ==""):
            numFatura = self.tc2.GetValue()
            log_Result = self.fatura.removingSubcontracts(numFatura)
            if not (log_Result ==""):
                wx.MessageDialog(None, log_Result , 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal()
        else:
            wx.MessageDialog(None, 'Nenhuma Fatura foi selecionada' , 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal()   
        self.tc2.SetValue("")
    
    def on_textFatura(self, event):
        pathFatura = self.tc1.GetPath()
        if(os.path.isfile(pathFatura)):
            if(self.fatura.checkExtensionFile(pathFatura)):
                self.fatura.createFilesXLXS(pathFatura)
                self.fatura.openOPC()
            else:
                wx.MessageDialog(None, 'Arquivo tem que ser do tipo .xslx' , 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal()
                self.tc1.SetPath("")
        else:
            wx.MessageDialog(None, 'Arquivo não encontrado' , 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal() 
            self.tc1.SetPath("")  
        
    def on_press(self, event):
        value_dominio = self.text_ctrl_dominio.GetValue()
        value_cpf = self.text_ctrl_cpf.GetValue()
        value_user = self.text_ctrl_user.GetValue()
        value_password = self.text_ctrl_password.GetValue()
        if value_dominio and value_cpf and value_user and value_password:
            log_login = self.fatura.loginSSW(value_dominio,value_cpf,value_user,value_password)  
            if(log_login == True):
                self.panel.Destroy()
                self.secondPanel()
            else:
                wx.MessageDialog(None, 'Login invalido', 'Error', wx.OK | wx.ICON_ERROR).ShowModal()
        else:
            wx.MessageDialog(None, 'Insira todos os valores', 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal()

    def on_save(self,event):
        if(self.tc1.GetPath() !=""):
            if not (self.fatura.checkProcessExcel()):
                self.fatura.saveprocess(self.tc1.GetPath())
                wx.MessageDialog(None, 'Processo Finalizado', 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal()
                wx.Exit()
            else:
                wx.MessageDialog(None, 'Feche qualquer arquivo Excel aberto' , 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal()       
        else:
            wx.MessageDialog(None, 'Nenhuma Fatura foi selecionada' , 'Warning', wx.OK | wx.ICON_INFORMATION).ShowModal()  
        