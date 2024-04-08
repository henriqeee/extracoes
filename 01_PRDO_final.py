import pandas as pd 
import time 
import win32com.client
import time
from datetime import datetime, timedelta
from tkinter import Tk, Button,StringVar,messagebox
from tkcalendar import Calendar
import pandas as pd 
import time 
import win32com.client
import subprocess
import time
from datetime import datetime, timedelta
import psutil

def select_date():
    def print_sel():
        selected_date.set(cal.selection_get())
        root.quit()

    root = Tk()
    selected_date = StringVar()
    cal = Calendar(root, selectmode='day', year=2024, month=4, day=5)
    cal.pack(padx=10,pady=10)

    button = Button(root, text="OK", command=print_sel)
    button.pack()

    root.mainloop()
    root.quit()
    return selected_date.get()


selected_date = select_date()


formatted_date = datetime.strptime(selected_date, '%Y-%m-%d').strftime('%d.%m.%Y')
print(formatted_date)
sap_gui_auto = win32com.client.GetObject("SAPGUI")
application = sap_gui_auto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)
tomorrow = (datetime.now() - timedelta(days=30)).strftime('%d.%m.%Y')




caminho_arquivo_excel = r'C:\\Users\\HE65465\\OneDrive - AGCO Corp\\Desktop\\Python PRDO\\Resultado\\PN_Problemas.XLSX'
nome_aba_excel = 'PN' 
df = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_aba_excel)
nome_da_coluna = 'Produtos com Quantidade Negativa'
coluna_selecionada = df[nome_da_coluna].to_clipboard(index=False)


row = 0

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/scwm/prdo"
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/btnGV_BUTTON_TEXT").press()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3").select()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5014/ctxtSO_DOCTY-LOW").text = "zon1"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5014/ctxtSO_DOCTY-HIGH").text = "zon2"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5014/ctxtSO_DGI_I-LOW").text = "1"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB1/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5010/ssubSUB_SEL_TAB1:/SCWM/SAPLUI_DLV_PRD:5014/ctxtSO_DGI_I-HIGH").text = "2"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDFR").text = formatted_date
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDTO").text = formatted_date
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDTO").setFocus()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDTO").caretPosition = 10
session.findById("wnd[0]").sendVKey(8)
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/btnGV_BUTTON_TEXT").press()
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_PRD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").currentCellColumn = ""
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_PRD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton("OIP_DISPLAY")
session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell").setCurrentCell (-1,"PRODUCTNO")
session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell").selectColumn("DOCNO_OD")
session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell").selectColumn("PRODUCTNO")
session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell").pressToolbarButton("&MB_FILTER")
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").setFocus()
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").press()
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(1).width = 35
session.findById("wnd[2]/tbar[0]/btn[8]").press()
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press()
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select()
session.findById("wnd[2]/tbar[0]/btn[24]").press()
session.findById("wnd[3]/usr/btnBUTTON_1").press()
session.findById("wnd[2]/tbar[0]/btn[8]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
lup=0
while True:
    data = datetime.now()
    dt = datetime.strptime(formatted_date, '%d.%m.%Y')
    dt = dt - timedelta(days=lup)
    formatted_date = dt.strftime('%d.%m.%Y')
    print(formatted_date)
    data_a = data.strftime('%d.%m.20%y')
    ex=0
    
    try: 
       
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/btnGV_BUTTON_TEXT").press()
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDFR").text = formatted_date
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDTO").text = formatted_date
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDTO").setFocus()
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_DLV_PRD:5001/tabsTAB_SEL/tabpOK_SEL_TAB3/ssubSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5018/subSUB_SEL_TAB3:/SCWM/SAPLUI_DLV_PRD:5017/ctxtPO_CRDTO").caretPosition = 10
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/btnGV_BUTTON_TEXT").press()
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_PRD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").setCurrentCell (0,"")
        session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_PRD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").selectedRows = "0"
        time.sleep(2)
        ex =1
        row = 0
    except:
        pass
    
    while ex == 1 :
        try: 
        
            session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_PRD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").setCurrentCell (0,"")
            session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_PRD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").selectedRows = row
            session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton ("OIP_DISPLAY")
            session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell").setCurrentCell( -1,"")
            session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell").selectAll()
            session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/cntlCONTAINER_TB_ODP1_1/shellcont/shell").pressButton ("ODP1_CREATE_FD")
            texto = session.findById('wnd[0]/sbar/pane[0]').text
            if texto == 'Nenhuma entrada marcada':
                pass
            else: 
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton("SAVE")
                session.findById("wnd[0]/tbar[1]/btn[23]").press()
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").currentCellColumn = "STATUS_GI"
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").selectColumn ("CHGDAT")
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").selectColumn ("STATUS_GI")
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").deselectColumn ("STATUS_LOADING")
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").pressToolbarButton ("&MB_FILTER")
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
                session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = data_a
                session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 10
                session.findById("wnd[2]").sendVKey (8)
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press()
                session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "n*"
                session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 2
                session.findById("wnd[2]").sendVKey (8)
                session.findById("wnd[1]").sendVKey (0)
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").setCurrentCell (-1,"")
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").selectAll()
                session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton ("OIP_POST_GM")
                texto = session.findById('wnd[0]/sbar/pane[0]').text
                if texto == 'Nenhuma entrada marcada':
                    
                    session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_FD:2210/subSUB_OIP_1_CONTENT:/SCWM/SAPLUI_DLV_FD:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell").setCurrentCell (-1,"")
                
                if session.ActiveWindow.Name == "wnd[1]":
                   
                    session.findById("wnd[1]").close()

                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                time.sleep(15)
           

            row += 1
        except:
            ex = 0
            
    lup = 1       
    if formatted_date == tomorrow:
        root = Tk()
        root.withdraw()  
        messagebox.showinfo("Informaçao", "Operaçao finalizada.")
        root.destroy()  
        break