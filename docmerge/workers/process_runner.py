import subprocess,sys,json
from pathlib import Path

def run_word_probe(path,timeout_seconds=30):
    # Absolūts ceļš: Word COM relatīvos risina pret savu darba mapi.
    target=str(Path(path).expanduser().resolve())
    try: cp=subprocess.run([sys.executable,'-m','docmerge.workers.word_probe',target],capture_output=True,text=True,timeout=timeout_seconds)
    except subprocess.TimeoutExpired: return {'ok':False,'status':'TIMEOUT','error':f'Word probe > {timeout_seconds}s'}
    text=(cp.stdout or '').strip().splitlines()
    try: data=json.loads(text[-1]) if text else {'ok':False,'error':cp.stderr.strip()}
    except Exception: data={'ok':False,'error':cp.stdout or cp.stderr}
    data['returncode']=cp.returncode; return data
