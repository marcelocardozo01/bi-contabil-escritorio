"""Aponta o PBIP para a pasta dados/exports DESTA copia do projeto.

O PBIP vem configurado para C:\\lastro-contabil\\dados\\exports\\.
Duas formas de abrir em outra pasta:

  python ferramentas/configurar_caminho_powerbi.py --atalho
      cria C:\\lastro-contabil como atalho (junction) para esta pasta;
      os arquivos do PBIP nao mudam (recomendado)

  python ferramentas/configurar_caminho_powerbi.py
      grava o caminho desta pasta no parametro PastaExports do PBIP
"""
import re
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
PADRAO = Path(r"C:\lastro-contabil")

if "--atalho" in sys.argv:
    if PADRAO.exists():
        sys.exit(f"{PADRAO} ja existe.")
    subprocess.run(["cmd", "/c", "mklink", "/J", str(PADRAO), str(RAIZ)], check=True)
else:
    pasta = str(RAIZ / "dados" / "exports") + "\\"
    for f in (RAIZ / "powerbi").glob("*.SemanticModel/definition/expressions.tmdl"):
        t = f.read_text(encoding="utf-8")
        f.write_text(re.sub(r'PastaExports = "[^"]*"', lambda m: f'PastaExports = "{pasta}"', t),
                     encoding="utf-8", newline="")
        print(f"{f.parts[-3]}: PastaExports = {pasta}")
