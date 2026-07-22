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
import os
import sys
from pathlib import Path

# Configura path
sys.path.insert(0, str(Path(__file__).parent / "utils"))

# Configura verificação forçada
os.environ["BOLETIM_FORCE_DEP_CHECK"] = "1"

try:
    # Verifica dependências
    import utils.verificador_dependencias as deps
    deps.verificar_dependencias_automaticamente(".")
    
    # Inicia aplicação
    from interface.app import RelatorioHorasApp
    app = RelatorioHorasApp()
    
except Exception as e:
    print(f"Erro ao iniciar aplicação: {e}")
    input("Pressione Enter para sair...")