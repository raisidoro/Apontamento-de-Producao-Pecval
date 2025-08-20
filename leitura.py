import openpyxl as xl
import wx
import datetime
import os
import shutil
import getpass as gt
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

setores = ['IE-INJECAO', 'IE-PINTURA', 'IE-MONTAGEM']
opcoes  = ['SIM', 'NAO']

def leitura():
    usuario   = gt.getuser()
    continuar = 'SIM'
    gerada    = f"C:\\Users\\{usuario}\\Desktop\\REPORT_Master__FUNCIONAL.xlsx"
    original  = r'\\files-gdbr01\\GDBR\\ADMINISTRATION\\IT\\Desenvolvimento\\Apontamentos de producao - Programas\\Um-a-Um_Homologado\\modelo\\REPORT_Master__FUNCIONAL.xlsx'
    target    = gerada

    shutil.copyfile(original, target)
    wb_modelo    = xl.load_workbook(gerada, data_only=True)
    ws_modeloPec = wb_modelo['REPORT']

    cont = 2
    while continuar == 'SIM':

        with wx.FileDialog(None, "Selecione arquivos Excel", wildcard="Excel Files (*.xlsm;*.xlsx)|*.xlsm;*.xlsx", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE) as fileDialog:
            
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            
            pathnames = fileDialog.GetPaths()
            print(f"\nArquivos selecionados: {pathnames}")

        dlg = wx.SingleChoiceDialog(None, ("Escolha seu setor"), "SETOR", setores)

        if dlg.ShowModal() == wx.ID_OK:
            setor = dlg.GetStringSelection()
            print(f"\nSetor: {setor}")
            setores.remove(setor)
        dlg.Destroy()

        for pathname in pathnames:

            print(f"\nProcessando: {pathname}")

            wb_reporte = xl.load_workbook(pathname, data_only=True)
            ws_reporte = wb_reporte['REPORT']

            maxcol = ws_reporte.max_column
            maxrow = ws_reporte.max_row

            col_reporte = [ws_reporte.cell(row=1, column=j).value for j in range(1, maxcol+1)]

            colunas_ignoradas = []
            for idx, col_name in enumerate(col_reporte, start=1):
                if col_name in ["QTD REPINTURA", "QTD RETRABALHO", "QTD SCRAP", "QTD REVISÃO"]:
                    colunas_ignoradas.append(idx)

            # i --> linha
            # j --> coluna
            for i in range(2, maxrow+1):
                val1 = ws_reporte.cell(row=i, column=1).value #Coluna1
                val2 = ws_reporte.cell(row=i, column=2).value #Coluna2

                # Verifica se as colunas 1 e 2 possuem valores
                if val1 is not None and val2 not in (None, ""):
                    # Processa as colunas restantes se tiver valor nas colunas 1 e 2
                    for j in range(1, maxcol+1):
                        # Ignora colunas que não devem ser processadas
                        if j in colunas_ignoradas:
                            continue
                        # Obtém o valor da célula
                        valor_celula = ws_reporte.cell(row=i, column=j).value
        
                        if valor_celula is None or valor_celula == "":
                            if 7 <= j <= 33 or j == 5:
                                valor_celula = ""

                        # Formata a célula como data, se necessário
                        if ("DATA" in col_reporte or "DATA DO DIA" in col_reporte) and isinstance(valor_celula, datetime.datetime):
                            valor_celula = valor_celula.date()
                            ws_modeloPec.cell(row=cont, column=j).number_format = "DD/MM/YYYY"

                        ws_modeloPec.cell(row=cont, column=j).value = valor_celula
                        ws_modeloPec.cell(row=cont, column=j).alignment = Alignment(horizontal='center')

                    cont += 1

        for col in ws_modeloPec.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws_modeloPec.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

        wb_modelo.save(gerada)

        dlg = wx.SingleChoiceDialog(None, ("Deseja continuar?"), "CONTINUAR", opcoes)

        if dlg.ShowModal() == wx.ID_OK:
            continuar = dlg.GetStringSelection()
        dlg.Destroy()
