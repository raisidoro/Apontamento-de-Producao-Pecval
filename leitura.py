# importing openpyxl module
import openpyxl as xl
from openpyxl import Workbook
from platform import java_ver
import wx
import openpyxl
import datetime
import os
import shutil
import getpass as gt
escolha=['IE-INJECAO','IE-PINTURA','IE-MONTAGEM']
rdtype=None
rtypes=[]
final=['SIM','NAO']

def leitura():
    usuario = gt.getuser()
    resposta = 'SIM'
    filename1 = f"C:\\Users\\{usuario}\\Desktop\\REPORT_Master__FUNCIONAL.xlsx"
    original = r'\\files-gdbr01\\GDBR\\ADMINISTRATION\\IT\\Desenvolvimento\\Apontamentos de producao - Programas\\Um-a-Um_Homologado\\modelo\\REPORT_Master__FUNCIONAL.xlsx'
    target = f"C:\\Users\\{usuario}\\Desktop\\REPORT_Master__FUNCIONAL.xlsx"

    shutil.copyfile(original, target)
    wb_modelo = xl.load_workbook(filename1, data_only=True)
    ws_modeloPec = wb_modelo['REPORT']
    cont = 2

    while resposta == 'SIM':
        with wx.FileDialog(None, "Open XYZ file", wildcard="Excel Files (*.xlsm;*.xlsx)|*.xlsm;*.xlsx", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            pathname = fileDialog.GetPath()
            print(f"\nOrigem: {pathname}")

        dlg = wx.SingleChoiceDialog(None, ("Escolha seu setor"), "SETOR", escolha)
        if dlg.ShowModal() == wx.ID_OK:
            setor = dlg.GetStringSelection()
            print(f"\nSetor: {setor}")
            escolha.remove(setor)
        dlg.Destroy()

        wb_reporte = xl.load_workbook(pathname, data_only=True)
        ws_reporte = wb_reporte['REPORT']

    #    wb_modelo = xl.load_workbook(filename1, data_only=True)
     #   ws_modeloPec = wb_modelo['REPORT']

        maxcol  = ws_reporte.max_column
        maxrow  = ws_reporte.max_row
        i =2
        col_reporte = [ws_reporte.cell(row=1, column=j).value for j in range(1, maxcol+1)]
        col_modelo  = [ws_modeloPec.cell(row=11, column=j).value for j in range(1, maxcol+1)]

        #cont = 2
        
        skip_columns = []
        for idx, col_name in enumerate(col_reporte, start=1):
            if col_name in ["QTD REPINTURA", "QTD RETRABALHO", "QTD SCRAP", "QTD REVISÃO"]:
                skip_columns.append(idx)

        for i in range(i, maxrow+1):
            # Só copia se tem dados relevantes
            val1 = ws_reporte.cell(row=i, column=1).value
            val2 = ws_reporte.cell(row=i, column=2).value
            if val1 is not None and val2 not in (None, ""):
                for j in range(1, maxcol+1):
                    if j in skip_columns:
                        continue
                    v = ws_reporte.cell(row=i, column=j).value
                    if v is None or v == "":
                        if 7 <= j <= 33 or j == 5:
                            v = ""
                    ws_modeloPec.cell(row=cont, column=j).value = v
            cont = cont + 1

        wb_modelo.save(filename1)

        dlg = wx.SingleChoiceDialog(None, ("Deseja continuar?"), "CONTINUAR", final)
        if dlg.ShowModal() == wx.ID_OK:
            resposta = dlg.GetStringSelection()
        dlg.Destroy()