import pandas as pd 
import time 
import win32com.client
import subprocess
import time
from datetime import datetime, timedelta
import psutil
def subtrair_quantidade(row):
    produto = row['Produto']
    quantidade_planilha1 = row['Quantidade']
    quantidade_planilha2 = produtos_quantidades_planilha2.get(produto, 0)
    return quantidade_planilha2 - quantidade_planilha1

time.sleep(3)
sap_gui_auto = win32com.client.GetObject("SAPGUI")
application = sap_gui_auto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)



time.sleep (5)


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/N/SCWM/PRDO"
session.findById("wnd[0]").sendVKey (0)

session.findById("wnd[0]/tbar[1]/btn[23]").press()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/btnGV_BUTTON_TEXT").press()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_FD:5000/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5012/ctxtSO_DGI_I-LOW").text = "1"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_FD:5000/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5012/ctxtSO_DGI_I-HIGH").text = "2"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_FD:5000/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5012/ctxtSO_DGI_I-HIGH").setFocus
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_FD:5000/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_FD:5012/ctxtSO_DGI_I-HIGH").caretPosition = 1
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_FD:5000/subSUB_ADV_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_START_ADVANCED").press()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/btnGV_BUTTON_TEXT").press()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").setCurrentCell (-1,"")
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").selectAll()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton ("OIP_POST_GM")

#SE DER ERRO -- 
#session.findById("wnd[1]/tbar[0]/btn[0]").press()
#session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressContextButton ("OIP_LOAD")
#session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").selectContextMenuItem ("OIP_CANCEL_LOAD")
#session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton ("OIP_DELETE")
#session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton ("SAVE")
#session.findById("wnd[0]/tbar[0]/btn[12]").press()






session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/scwm/mon"
session.findById("wnd[0]").sendVKey(0)
if session.ActiveWindow.Name == "wnd[1]":
    session.findById("wnd[1]/usr/ctxtP_LGNUM").text = "ibi"
    session.findById("wnd[1]/usr/ctxtP_MONIT").text = "z001"
    session.findById("wnd[1]/usr/ctxtP_MONIT").setFocus
    session.findById("wnd[1]/usr/ctxtP_MONIT").caretPosition = 4
    session.findById("wnd[1]").sendVKey (8)

session.findById("wnd[0]/tbar[1]/btn[18]").press()
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").expandNode("C000000001")
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").expandNode("C000000004")
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").expandNode("N000000010")
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedNode = "N000000011"
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").topNode = "C000000001"
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").doubleClickNode("N000000011")
session.findById("wnd[1]/usr/ctxtS_DOCTY-LOW").text = "zon1"
session.findById("wnd[1]/usr/ctxtS_DOCTY-HIGH").text = "zon2"
session.findById("wnd[1]/usr/ctxtS_DGII-LOW").text = "1"
session.findById("wnd[1]/usr/ctxtS_DGII-HIGH").text = "2"
session.findById("wnd[1]/usr/ctxtS_DGII-HIGH").setFocus
session.findById("wnd[1]/usr/ctxtS_DGII-HIGH").caretPosition = 1
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell(3,"PRODUCTNO")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").contextMenu()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\HE65465\OneDrive - AGCO Corp\Desktop\Python PRDO\Resultado"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "PRDO.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[11]").press()

time.sleep(10)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedNode = "0000000171"
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").doubleClickNode("0000000171")
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\HE65465\OneDrive - AGCO Corp\Desktop\Python PRDO\Resultado"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "VIRK.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[11]").press()

time.sleep(30)

nome_processo = 'EXCEL.EXE'

processos = [p for p in psutil.process_iter(
    attrs=['pid', 'name'])if nome_processo in p.info['name']]

for processo in processos:
    pid = processo.info['pid']
    psutil.Process(pid).terminate()
time.sleep(5)
tabela_PRDO = pd.read_excel(r'C:\Users\HE65465\OneDrive - AGCO Corp\Desktop\Python PRDO\Resultado\PRDO.XLSX')
tabela_VIRK = pd.read_excel(r'C:\Users\HE65465\OneDrive - AGCO Corp\Desktop\Python PRDO\Resultado\VIRK.XLSX')

writer = pd.ExcelWriter(r'C:\Users\HE65465\OneDrive - AGCO Corp\Desktop\Python PRDO\Resultado\PN_Problemas.xlsx',engine='xlsxwriter')
agrupa_prdo = tabela_PRDO.groupby(['Produto',],as_index=False)[['Quantidade']].sum()
agrupa_virk = tabela_VIRK.groupby(['Produto'],as_index=False)[['Quantidade']].sum()



produtos_quantidades_planilha2 = dict(zip(agrupa_virk['Produto'],agrupa_virk['Quantidade']))



agrupa_prdo['Quantidade_VIRK']=''
agrupa_prdo['Quantidade_PRDO']=''
for index, row in agrupa_prdo.iterrows():
  
    valor_x = row['Produto']
    
 
    if valor_x in agrupa_virk['Produto'].values:
       
        valor_ods = agrupa_virk.loc[agrupa_virk['Produto'] == valor_x, 'Quantidade'].values[0]
        
       
        agrupa_prdo.loc[index, 'Quantidade_VIRK'] = valor_ods
    else:
       
        agrupa_prdo.loc[index, 'Quantidade_VIRK'] = 0
for index, row in agrupa_prdo.iterrows():
  
    valor_x = row['Produto']
    
 
    if valor_x in agrupa_prdo['Produto'].values:
       
        valor_ods = agrupa_prdo.loc[agrupa_prdo['Produto'] == valor_x, 'Quantidade'].values[0]
        
       
        agrupa_prdo.loc[index, 'Quantidade_PRDO'] = valor_ods
    else:
       
        agrupa_prdo.loc[index, 'Quantidade_PRDO'] = 0


       
agrupa_prdo['Quantidade'] = agrupa_prdo.apply(subtrair_quantidade, axis=1)
valores_negativos = agrupa_prdo[agrupa_prdo['Quantidade'] <0]



produtos_negativos = valores_negativos['Produto']
df_produtos_negativos = pd.DataFrame({'Produtos com Quantidade Negativa': produtos_negativos})

tabela2 = pd.DataFrame(tabela_PRDO['Documento'].unique())

agrupa_prdo.to_excel(writer,sheet_name='ANALISE',index=False)
df_produtos_negativos.to_excel(writer,sheet_name='PN',index=False)
tabela2.to_excel(writer,sheet_name='DOC',index=False)
writer.close()