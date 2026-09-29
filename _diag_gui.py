"""GUI smoke tests: logi + jauno opciju saglabāšana projekta JSON."""
from __future__ import annotations

import json
import tempfile
import tkinter as tk
from pathlib import Path

from docmerge.domain.models import Project
from docmerge.gui.main_window import MainWindow

project = Project()
project.options.insert_titles = False
project.options.unlock_protected_view = False
data = project.to_dict()
restored = Project.from_dict(json.loads(json.dumps(data)))
assert restored.options.insert_titles is False, restored.options
assert restored.options.unlock_protected_view is False, restored.options
print('projekta round-trip:', restored.to_dict()['options'])

root = tk.Tk()
root.withdraw()
window = MainWindow(root)
root.update()
print('checkbox tituli:', window.titles.get(), '| unlock:', window.unlock.get())
assert window.titles.get() is True and window.unlock.get() is True

window.titles.set(False)
window.unlock.set(False)
window.sync()
print('sync ->', window.project.options.insert_titles, window.project.options.unlock_protected_view)
assert window.project.options.insert_titles is False
assert window.project.options.unlock_protected_view is False

window.new()
print('new() atgriež noklusējumus:', window.titles.get(), window.unlock.get())
window.refresh()
assert 'Log:' in window.summary.get()
assert callable(window.merge_all_word)

with tempfile.TemporaryDirectory() as tmp:
    window.output.set(str(Path(tmp) / 'x.docx'))
    window.sync()
    assert window.project.output_path.endswith('x.docx')

root.update()
root.destroy()
print('GUI smoke OK')
