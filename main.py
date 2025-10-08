# main_gui.py
# Ponto de entrada principal para a aplicação GUI
# Caminho: relatorio_horas/main_gui.py
#
# Autor: Valdinei Lankewicz
# Criado em: 22/07/2025
# Histórico:
# - 22/07/2025: Criação inicial do ponto de entrada da aplicação
# main.py

# main.py
import importlib.util
import os
import subprocess
import sys

sys.path.insert(0, "./utils")  # garante que utils esteja no path

# 1) Desliga o auto-run em thread e força verificação completa bloqueante
# os.environ["BOLETIM_AUTORUN"] = "0"
os.environ["BOLETIM_FORCE_DEP_CHECK"] = "1"

if importlib.util.find_spec("packaging") is None:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "packaging"])

# 3) Importa o verificador e roda de forma BLOQUEANTE
import utils.verificador_dependencias as deps

deps.verificar_dependencias_automaticamente(".")  # só retorna quando terminar

# 4) Agora é seguro carregar o resto
from interface.app import RelatorioHorasApp  # noqa: E402

if __name__ == "__main__":
    app = RelatorioHorasApp()
