from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path


ROOT = Path(r"C:\Users\Roberto Moura\Downloads\Propostas do sistema")
BANCO = Path(r"C:\Users\Roberto Moura\Downloads\Banco de Dados Operacional")
ENTRADA = BANCO / "00 - Entrada diaria"
FONTE = ROOT / "04 - Fonte"
LOG = ENTRADA / "99 - Logs" / "atualizacoes_ao_salvar.jsonl"


def wait_stable(path: Path, attempts: int = 10, delay: float = 1.0) -> bool:
    last = -1
    for _ in range(attempts):
        if not path.exists():
            time.sleep(delay)
            continue
        size = path.stat().st_size
        if size == last and size > 0:
            return True
        last = size
        time.sleep(delay)
    return path.exists() and path.stat().st_size > 0


def classify(path: Path) -> tuple[Path, Path]:
    rel = path.relative_to(ENTRADA)
    parts = rel.parts
    if "Em aberto" in parts:
        return BANCO / "01 - ERP Olist" / "Propostas" / "Em aberto" / path.name, FONTE / "01 - ERP Propostas" / "CSVs por status - Em aberto" / path.name
    if "Concluidas" in parts:
        return BANCO / "01 - ERP Olist" / "Propostas" / "Concluidas" / path.name, FONTE / "01 - ERP Propostas" / "CSVs por status - Concluidas" / path.name
    if "Nao aprovadas" in parts:
        return BANCO / "01 - ERP Olist" / "Propostas" / "Nao aprovadas" / path.name, FONTE / "01 - ERP Propostas" / "CSVs por status - Nao aprovadas" / path.name
    if "Marcadores" in parts:
        return BANCO / "01 - ERP Olist" / "Marcadores" / path.name, FONTE / "04 - Marcadores e setores" / "Propostas por mes com marcadores" / path.name
    if "Propostas comerciais xlsx" in parts:
        return BANCO / "01 - ERP Olist" / "Propostas" / "Propostas comerciais.xlsx", FONTE / "01 - ERP Propostas" / "Propostas comerciais.xlsx"
    if "02 - SellFlux Atendimentos" in parts or "02 - Atendimentos" in parts:
        return BANCO / "02 - SellFlux" / "Atendimentos" / path.name, FONTE / "02 - Atendimentos" / path.name
    if "Compras ERP" in parts or "03 - Compras" in parts:
        return BANCO / "01 - ERP Olist" / "Compras" / path.name, FONTE / "03 - Compras" / path.name
    if "Vendedores setores" in parts:
        return BANCO / "03 - Regras internas" / "Vendedores setores" / path.name, FONTE / "04 - Marcadores e setores" / path.name
    if "Produtos kits" in parts:
        return BANCO / "01 - ERP Olist" / "Produtos" / path.name, FONTE / "04 - Marcadores e setores" / path.name
    if "Clientes" in parts:
        return BANCO / "01 - ERP Olist" / "Clientes e auxiliares" / path.name, FONTE / "01 - ERP Propostas" / path.name
    return BANCO / "99 - A classificar" / path.name, FONTE / "99 - A classificar" / path.name


def copy_to(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def log_event(payload: dict) -> None:
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")


def run_project_scripts() -> list[dict]:
    scripts = [
        ROOT / "08 - Códigos" / "organizar_fontes_para_atualizacao_diaria.py",
        ROOT / "08 - Códigos" / "criar_apresentacao_e_banco_operacional.py",
    ]
    results = []
    for script in scripts:
        if not script.exists():
            results.append({"script": str(script), "status": "não encontrado"})
            continue
        proc = subprocess.run(["python", str(script)], cwd=str(ROOT), capture_output=True, text=True, timeout=300)
        results.append({"script": str(script), "returncode": proc.returncode, "stdout": proc.stdout[-1000:], "stderr": proc.stderr[-1000:]})
    return results


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--arquivo", required=True)
    args = parser.parse_args()
    src = Path(args.arquivo)
    if src.name.startswith("~$"):
        return
    ok = wait_stable(src)
    banco_dest, fonte_dest = classify(src)
    if ok:
        copy_to(src, banco_dest)
        copy_to(src, fonte_dest)
        scripts = run_project_scripts()
    else:
        scripts = []
    log_event({
        "quando": datetime.now().isoformat(timespec="seconds"),
        "arquivo": str(src),
        "estavel": ok,
        "banco_destino": str(banco_dest),
        "fonte_destino": str(fonte_dest),
        "scripts": scripts,
    })
    print(json.dumps({"ok": ok, "banco": str(banco_dest), "fonte": str(fonte_dest)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
