from __future__ import annotations
import shutil, zipfile
from pathlib import Path

root=Path(__file__).resolve().parents[1]; stage=root/'dist'/'fastgpt-file-preprocessor'; archive=root/'dist'/'fastgpt-file-preprocessor.zip'
if stage.exists(): shutil.rmtree(stage)
stage.mkdir(parents=True)
for name in ('src','requirements.txt','requirements-dev.txt','run.ps1','README.md','ARCHITECTURE.md','config.json','config.example.json'):
    source=root/name; target=stage/name
    if source.is_dir(): shutil.copytree(source,target,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
    else: shutil.copy2(source,target)
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
    for file in stage.rglob('*'):
        if file.is_file(): z.write(file,file.relative_to(stage.parent))
print(archive)
