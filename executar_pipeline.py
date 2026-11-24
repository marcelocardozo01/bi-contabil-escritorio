"""Roda o projeto de ponta a ponta.

  python executar_pipeline.py            -> ingestao, regras + validacao e painel HTML
  python executar_pipeline.py --gerar    -> antes, gera de novo a carteira ficticia
"""
import os
import subprocess
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
ETAPAS = [("Ingestao dos balancetes", "pipeline/01_ingestao.py"), ("Regras, exportacao e validacao", "pipeline/02_regras.py"),
          ("Painel HTML", "painel_html/gerar_painel.py")]
if "--gerar" in sys.argv:
    ETAPAS.insert(0, ("Geracao da carteira ficticia", "gerador/gerar_dados.py"))
env = dict(os.environ, PYTHONIOENCODING="utf-8")
for nome, script in ETAPAS:
    print(f"\n=== {nome} ===", flush=True)
    t0 = time.time()
    if subprocess.run([sys.executable, str(RAIZ / script)], cwd=RAIZ, env=env).returncode != 0:
        sys.exit(f"Falhou em: {nome}")
    print(f"({time.time() - t0:.1f}s)", flush=True)
print("\nPipeline concluido.")
