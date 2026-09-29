import json,sys
from pathlib import Path

def main():
    if len(sys.argv)<2: print(json.dumps({'ok':False,'error':'missing path'})); return 3
    # Word COM relatīvu ceļu risina pret Word darba mapi (C:\WINDOWS\system32),
    # tāpēc probe vienmēr padod ABSOLŪTU ceļu.
    raw=str(sys.argv[1]); target=str(Path(raw).expanduser().resolve())
    session=None; doc=None
    try:
        import pythoncom, win32com.client
        from docmerge.engines.word_com import open_word_session
        pythoncom.CoInitialize(); session=open_word_session(win32com.client)
        doc=session.word.Documents.Open(target,ReadOnly=True,AddToRecentFiles=False,ConfirmConversions=False,NoEncodingDialog=True)
        print(json.dumps({'ok':True,'status':'OK','word_owned':session.owned,'dispatch_method':session.method})); return 0
    except Exception as e:
        print(json.dumps({'ok':False,'status':'ERROR','error':str(e)},ensure_ascii=False)); return 2
    finally:
        try:
            if doc is not None: doc.Close(False)
        except Exception: pass
        # Tikai pašu izveidoto instanci drīkst aizvērt (lietotāja sesija paliek).
        try:
            if session is not None: session.close()
        except Exception: pass
        try:
            import pythoncom; pythoncom.CoUninitialize()
        except Exception: pass
if __name__=='__main__': raise SystemExit(main())
