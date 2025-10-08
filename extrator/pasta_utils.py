# extrator/pasta_utils.py
# Funções auxiliares para salvar e carregar o caminho da última pasta usada
# Caminho: relatorio_horas/extrator/pasta_utils.py
#
# Autor: Valdinei Lankewicz
# Criado em: 22/07/2025
# Histórico:
# - 22/07/2025: Módulo separado para controle de persistência de pasta.

import os


# =================================================================================
# FUNÇÃO: carregar_ultima_pasta(caminho_txt)
# Retorna o caminho salvo da última pasta utilizada, se existir.
# =================================================================================
def carregar_ultima_pasta(caminho_txt: str) -> str:
    if os.path.exists(caminho_txt):
        try:
            with open(caminho_txt, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            return ""
    return ""


# =================================================================================
# FUNÇÃO: salvar_ultima_pasta(caminho_txt, pasta)
# Salva o caminho da última pasta em um arquivo local.
# =================================================================================
def salvar_ultima_pasta(caminho_txt: str, pasta: str):
    try:
        with open(caminho_txt, "w", encoding="utf-8") as f:
            f.write(pasta.strip())
    except Exception as e:
        print(f"[ERRO] Falha ao salvar a última pasta: {e}")
