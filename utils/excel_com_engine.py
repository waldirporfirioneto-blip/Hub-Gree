import os
import time
import shutil
import tempfile
import subprocess
import traceback
import pywintypes
import pythoncom
import win32com.client as win32
import win32com.client.dynamic as dynamic

def _codigo_com_error(e):
    if isinstance(e, pywintypes.com_error):
        args = e.args or ()
        return args[0] if args else None
    return None

def _deve_tentar_reabrir_excel(e):
    codigo = _codigo_com_error(e)
    return codigo in {-2147418111, -2146827284, -2147023174}

def _normalizar_titulo_planilha(titulo):
    texto = str(titulo or "").strip()
    return texto[:31] if texto else "Planilha"

def _fechar_excel_forcado(excel):
    if excel is None: return None
    try:
        for workbook in list(excel.Workbooks):
            try: workbook.Close(SaveChanges=False)
            except: pass
    except: pass
    try: excel.Quit()
    except: pass
    return None

def _reinicializar_excel(excel):
    _fechar_excel_forcado(excel)
    try: pythoncom.CoUninitialize()
    except: pass
    try: pythoncom.CoInitialize()
    except: pass
    try: subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], check=False, capture_output=True, text=True, timeout=10)
    except: pass
    return abrir_excel_silencioso()

def abrir_excel_silencioso():
    excel = None
    try: pythoncom.CoInitialize()
    except: pass
    try:
        caminho_cache = os.path.join(tempfile.gettempdir(), 'gen_py')
        if os.path.exists(caminho_cache): shutil.rmtree(caminho_cache, ignore_errors=True)
    except: pass

    for _ in range(4):
        try:
            excel = dynamic.Dispatch("Excel.Application")
            _ = excel.Workbooks
            break
        except:
            try:
                excel = win32.DispatchEx("Excel.Application")
                _ = excel.Workbooks
                break
            except:
                try:
                    excel = win32.Dispatch("Excel.Application")
                    _ = excel.Workbooks
                    break
                except: excel = None
        if excel is None: time.sleep(1)

    if excel is None: return None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
    except: pass
    return excel

def abrir_workbook_com_retry(excel, caminho_absoluto, tentativas=8, atraso=1.0):
    instancia_excel = excel
    caminho_absoluto = os.path.abspath(caminho_absoluto)
    for tentativa in range(1, tentativas + 1):
        try:
            wb = instancia_excel.Workbooks.Open(
                Filename=caminho_absoluto, ReadOnly=False, UpdateLinks=False, 
                Notify=False, IgnoreReadOnlyRecommended=True, Editable=True
            )
            return instancia_excel, wb
        except pywintypes.com_error as e:
            if _deve_tentar_reabrir_excel(e) and tentativa < tentativas:
                instancia_excel = _reinicializar_excel(instancia_excel)
                if instancia_excel is None: raise
                time.sleep(atraso)
                continue
            raise
        except Exception: raise
    return instancia_excel, None

def obter_campo_dinamico(pt, nome_procurado):
    try:
        for i in range(1, pt.PivotFields().Count + 1):
            campo = pt.PivotFields(i)
            nome_campo = str(campo.Name).strip().lower()
            if nome_campo == nome_procurado.strip().lower(): return campo
        for i in range(1, pt.PivotFields().Count + 1):
            campo = pt.PivotFields(i)
            nome_campo = str(campo.Name).strip().lower()
            if nome_procurado.strip().lower() in nome_campo: return campo
    except: pass
    try: return pt.PivotFields(nome_procurado)
    except: return None