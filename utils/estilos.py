# utils/estilos.py
# Configurações opcionais de estilos personalizados (fonte, cor, padding, etc.)
# Caminho: relatorio_horas/utils/estilos.py
#
# Autor: Valdinei Lankewicz
# Criado em: 22/07/2025
# Histórico:
# - 22/07/2025: Arquivo reservado para customização visual futura

# Este módulo pode ser usado no futuro para:
# - Criar estilos personalizados de Treeview, Label, Botões
# - Registrar temas do ttkbootstrap
# - Aplicar padronizações visuais (ex: fonte única, alinhamento, cores)

# Exemplo (a ser usado em app):
# from utils.estilos import aplicar_estilos
# aplicar_estilos(ttkb.Style())


def aplicar_estilos(style):
    style.configure("Treeview", font=("Segoe UI", 10))
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("TButton", padding=6)
    style.configure("TLabel", font=("Segoe UI", 10))
