# utils/verificador_dependencias.py
# Função: Verifica e instala automaticamente as dependências do projeto.
# Autor: Valdinei Lankewicz
# Criado em: 01/08/2025 - Última atualização: 09/08/2025

import ast
import concurrent.futures
import hashlib
import importlib.util
import json
import os
import subprocess
import sys
import threading
import time
from importlib import metadata as importlib_metadata  # Py 3.8+
from pathlib import Path
from typing import Callable, Dict, Set

# =========================
# REPORTER (GUI status bar)
# =========================
_reporter: Callable[[str], None] | None = None
_REPORT_LOCK = threading.Lock()


def set_reporter(fn: Callable[[str], None] | None) -> None:
    global _reporter
    _reporter = fn


def _report(msg: str, level: str = "INFO") -> None:
    """Thread-safe reporting com níveis de log."""
    with _REPORT_LOCK:
        timestamp = time.strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {msg}"
        try:
            print(formatted_msg)
        except Exception:
            pass
        if _reporter:
            try:
                _reporter(msg)
            except Exception as e:
                print(f"Erro no reporter: {e}")


# =========================
# CONFIGURAÇÃO
# =========================
CACHE_FILE = ".deps_state.json"
ENV_SKIP = os.getenv("BOLETIM_SKIP_DEP_CHECK") == "1"
ENV_FORCE = os.getenv("BOLETIM_FORCE_DEP_CHECK") == "1"
PARALLEL_INSTALLS = int(os.getenv("BOLETIM_PARALLEL_INSTALLS", "3"))
AUTORUN = os.getenv("BOLETIM_AUTORUN", "0") == "1"


def atualizar_pip() -> bool:
    """Atualiza o pip se necessário. Retorna True se bem-sucedido."""
    try:
        import pip
        from packaging import version

        current_version = version.parse(pip.__version__)
        min_version = version.parse("23.0")

        if current_version < min_version:
            _report(f"⬆️ Atualizando pip de {pip.__version__}...")
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--upgrade", "pip"],
                capture_output=True,
                text=True,
                timeout=120,
            )
            if result.returncode == 0:
                _report("✅ pip atualizado com sucesso.")
                return True
            _report(f"❌ Falha ao atualizar pip: {result.stderr}")
            return False
        _report(f"✔️ pip ok (v{pip.__version__})")
        return True
    except Exception as e:
        _report(f"⚠️ Falha ao verificar/atualizar pip: {e}")
        return False


MODULOS_NATIVOS = frozenset(sys.builtin_module_names) | {
    "os",
    "sys",
    "time",
    "datetime",
    "math",
    "re",
    "subprocess",
    "pathlib",
    "logging",
    "threading",
    "typing",
    "json",
    "csv",
    "ast",
    "shutil",
    "itertools",
    "functools",
    "collections",
    "tkinter",
    "email",
    "copy",
    "platform",
    "http",
    "unittest",
    "argparse",
    "zipfile",
    "inspect",
    "urllib",
    "hashlib",
    "base64",
    "uuid",
    "tempfile",
    "sqlite3",
}

MAPPING_IMPORT_PIP = {
    "fitz": "PyMuPDF",
    "cv2": "opencv-python",
    "PIL": "Pillow",
    "sklearn": "scikit-learn",
    "yaml": "pyyaml",
    "bs4": "beautifulsoup4",
    "requests_html": "requests-html",
    "win32api": "pywin32",
    "psutil": "psutil",
    "ttkbootstrap": "ttkbootstrap",
    "tksheet": "tksheet",
    "openpyxl": "openpyxl",
    "reportlab": "reportlab",
    "matplotlib": "matplotlib",
    "pyarrow": "pyarrow",
    "fastparquet": "fastparquet",
}

DEPENDENCIAS_OBRIGATORIAS = frozenset(
    {
        "pyarrow",
        "packaging",
        "ttkbootstrap",
        "tksheet",
        "openpyxl",
        "reportlab",
        "matplotlib",
        "PyMuPDF",
        "pandas",
        "numpy",
    }
)


def encontrar_imports_em_arquivo(caminho_arquivo: Path) -> Set[str]:
    """Extrai imports de um arquivo Python usando AST."""
    pacotes: set[str] = set()
    try:
        with open(caminho_arquivo, "r", encoding="utf-8", errors="ignore") as file:
            conteudo = file.read()
        if len(conteudo.strip()) < 10:
            return pacotes
        arvore = ast.parse(conteudo, filename=str(caminho_arquivo))
        for node in ast.walk(arvore):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    nome_base = alias.name.split(".")[0]
                    if nome_base and nome_base.isidentifier():
                        pacotes.add(nome_base)
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    nome_base = node.module.split(".")[0]
                    if nome_base and nome_base.isidentifier():
                        pacotes.add(nome_base)
    except SyntaxError:
        _report(f"⚠️ Erro de sintaxe em {caminho_arquivo}")
    except Exception as e:
        _report(f"⚠️ Erro ao analisar {caminho_arquivo}: {e}")
    return pacotes


def coletar_imports_do_projeto(diretorio_raiz: str = ".") -> Set[str]:
    pacotes_encontrados: set[str] = set()
    root = Path(diretorio_raiz).resolve()

    ignore_patterns = {
        "venv",
        "env",
        ".venv",
        ".env",
        ".git",
        "__pycache__",
        ".pytest_cache",
        "site-packages",
        "dist",
        "build",
        ".tox",
        "node_modules",
    }

    arquivos_py: list[Path] = []
    for pasta, _, lista_nomes in os.walk(root):  # <— era `arquivos`
        if any(pattern in pasta for pattern in ignore_patterns):
            continue
        for nome_arquivo in lista_nomes:  # <— era `arquivo`
            if nome_arquivo.endswith(".py") and not nome_arquivo.startswith("test_"):
                caminho = Path(pasta) / nome_arquivo
                try:
                    if caminho.stat().st_size > 1024 * 1024:
                        continue
                    arquivos_py.append(caminho)
                except OSError:
                    continue

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        futures: dict[concurrent.futures.Future[Set[str]], Path] = {
            executor.submit(encontrar_imports_em_arquivo, arq): arq for arq in arquivos_py
        }

        for future in concurrent.futures.as_completed(futures):
            try:
                imports = future.result(timeout=5)
                pacotes_encontrados.update(imports)
            except concurrent.futures.TimeoutError:
                arq_path = futures[future]  # <— era `arquivo = ...`
                _report(f"⏰ Timeout ao processar {arq_path}")
            except Exception as e:
                arq_path = futures[future]  # <— era `arquivo = ...`
                _report(f"⚠️ Erro ao processar {arq_path}: {e}")

    return pacotes_encontrados


def normalizar_pacote(nome: str) -> str:
    """Converte nome de import para nome do pacote pip."""
    return MAPPING_IMPORT_PIP.get(nome, nome)


def instalar_pacote(pacote: str) -> bool:
    """Instala um pacote usando pip. Retorna True se bem-sucedido."""
    try:
        resultado = subprocess.run(
            [sys.executable, "-m", "pip", "install", pacote, "--quiet"],
            capture_output=True,
            text=True,
            timeout=300,
        )
        if resultado.returncode == 0:
            _report(f"✅ Instalado: {pacote}")
            return True
        _report(f"❌ Falha ao instalar {pacote}: {resultado.stderr.strip()}")
        return False
    except subprocess.TimeoutExpired:
        _report(f"⏰ Timeout ao instalar: {pacote}")
        return False
    except Exception as e:
        _report(f"❌ Erro ao instalar {pacote}: {e}")
        return False


def modulo_instalado(nome_import: str) -> bool:
    """Verifica se um módulo está instalado e importável."""
    try:
        spec = importlib.util.find_spec(nome_import)
        if spec is not None:
            return True
    except (ImportError, AttributeError, ValueError, ModuleNotFoundError):
        pass
    pacote_pip = normalizar_pacote(nome_import)
    try:
        importlib_metadata.version(pacote_pip)
        return True
    except Exception:
        pass
    return False


def get_version(nome_import: str) -> str | None:
    """Obtém a versão de um pacote instalado."""
    pacote_pip = normalizar_pacote(nome_import)
    try:
        return importlib_metadata.version(pacote_pip)
    except Exception:
        return None


def verificar_e_instalar(pacotes: Set[str]) -> Dict[str, bool]:
    """Verifica e instala pacotes. Retorna dict com status de cada pacote."""
    resultados: Dict[str, bool] = {}
    pacotes_verificar: list[str] = []
    for pacote in sorted(pacotes):
        pacote_pip = normalizar_pacote(pacote)
        if pacote_pip in MODULOS_NATIVOS or not pacote_pip or pacote in ("tkinter", "tk"):
            continue
        pacotes_verificar.append(pacote)

    pacotes_instalar: list[str] = []
    for pacote in pacotes_verificar:
        if modulo_instalado(pacote):
            pacote_pip = normalizar_pacote(pacote)
            versao = get_version(pacote)
            if versao:
                _report(f"✔️ {pacote_pip} já instalado (v{versao})")
            else:
                _report(f"✔️ {pacote_pip} já instalado")
            resultados[pacote] = True
        else:
            pacotes_instalar.append(pacote)

    if pacotes_instalar:
        _report(f"📦 Instalando {len(pacotes_instalar)} pacotes...")
        with concurrent.futures.ThreadPoolExecutor(max_workers=PARALLEL_INSTALLS) as executor:
            futures: Dict[concurrent.futures.Future[bool], str] = {}
            for pacote in pacotes_instalar:
                pacote_pip = normalizar_pacote(pacote)
                _report(f"📦 Instalando: {pacote_pip}")
                future = executor.submit(instalar_pacote, pacote_pip)
                futures[future] = pacote
            for future in concurrent.futures.as_completed(futures):
                pacote = futures[future]
                try:
                    sucesso = future.result()
                    resultados[pacote] = sucesso
                except Exception as e:
                    _report(f"❌ Erro inesperado ao instalar {pacote}: {e}")
                    resultados[pacote] = False
    return resultados


def _cache_path() -> Path:
    """Retorna o caminho absoluto do arquivo de cache como Path."""
    return Path(CACHE_FILE).resolve()


def _load_cache() -> Dict[str, object]:
    """Carrega o cache de dependências."""
    path: Path = _cache_path()
    if path.exists():
        try:
            with path.open("r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            _report(f"⚠️ Cache corrompido, será recriado: {e}")
            return {}
    return {}


def _save_cache(data: Dict[str, object]) -> None:
    """Salva o cache de dependências."""
    path: Path = _cache_path()
    try:
        with path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError as e:
        _report(f"⚠️ Falha ao salvar cache: {e}")


def limpar_cache() -> None:
    """Remove o arquivo de cache."""
    cache_file: Path = _cache_path()
    if cache_file.exists():
        cache_file.unlink()
        _report("🗑️ Cache de dependências removido.")


def _project_fingerprint(root: str = ".") -> str:
    """Gera hash do projeto para detectar mudanças."""
    h = hashlib.sha256()
    root_path = Path(root).resolve()
    arquivos_relevantes: list[tuple[Path, float]] = []
    for pasta, _, arquivos in os.walk(root_path):
        if any(ign in pasta for ign in ("venv", ".git", "__pycache__", "site-packages")):
            continue
        for arquivo in arquivos:
            if arquivo.endswith((".py", "requirements.txt", "setup.py", "pyproject.toml")):
                caminho = Path(pasta) / arquivo
                try:
                    stat = caminho.stat()
                    if stat.st_size < 100 * 1024:
                        arquivos_relevantes.append((caminho, stat.st_mtime))
                except OSError:
                    continue
    arquivos_relevantes.sort(key=lambda x: x[0])
    for caminho, mtime in arquivos_relevantes:
        try:
            caminho_relativo = caminho.relative_to(root_path)
            h.update(str(caminho_relativo).encode("utf-8"))
            h.update(str(int(mtime)).encode("utf-8"))
        except Exception:
            continue
    return h.hexdigest()


_already_ran = False
_EXECUTION_LOCK = threading.Lock()


def verificar_dependencias_automaticamente(diretorio_raiz: str = ".") -> Dict[str, bool]:
    """Escaneia e instala dependências automaticamente."""
    global _already_ran
    with _EXECUTION_LOCK:
        if _already_ran:
            _report("⏭️ Verificação já executada nesta sessão.")
            return {}
        _already_ran = True

    if ENV_SKIP:
        _report("⏭️ Verificação pulada (BOLETIM_SKIP_DEP_CHECK=1).")
        return {}

    _report("🔍 Iniciando verificação de dependências...")
    inicio = time.time()
    try:
        fp_atual = _project_fingerprint(diretorio_raiz)
        cache = _load_cache()
        fp_cache = cache.get("fingerprint")
        if fp_cache == fp_atual and not ENV_FORCE:
            _report("✅ Projeto inalterado. Dependências OK (cache).")
            return {}
        _report("📂 Escaneando imports do projeto...")
        pacotes = coletar_imports_do_projeto(diretorio_raiz)
        pacotes.update(DEPENDENCIAS_OBRIGATORIAS)
        _report(f"📋 Encontrados {len(pacotes)} imports únicos.")
        resultados = verificar_e_instalar(pacotes)
        cache.update(
            {
                "fingerprint": fp_atual,
                "checked_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "packages": sorted(list({normalizar_pacote(p) for p in pacotes})),
                "results": resultados,
            }
        )
        _save_cache(cache)
        tempo_total = time.time() - inicio
        sucessos = sum(1 for v in resultados.values() if v)
        total = len(resultados)
        _report(f"✅ Verificação concluída em {tempo_total:.1f}s. Sucessos: {sucessos}/{total}")
        return resultados
    except KeyboardInterrupt:
        _report("⏹️ Verificação interrompida pelo usuário.")
        return {}
    except Exception as e:
        _report(f"❌ Erro durante verificação: {e}")
        return {}


def listar_dependencias_instaladas() -> Dict[str, str]:
    """Lista todas as dependências instaladas com suas versões."""
    instaladas: Dict[str, str] = {}
    for pacote in DEPENDENCIAS_OBRIGATORIAS:
        versao = get_version(pacote)
        if versao:
            instaladas[pacote] = versao
    return instaladas


def _executar_no_import() -> None:
    """Executa verificações básicas na importação do módulo."""
    try:
        if not ENV_SKIP:
            threading.Thread(target=atualizar_pip, daemon=True).start()
            threading.Thread(
                target=verificar_dependencias_automaticamente,
                args=(".",),
                daemon=True,
            ).start()
    except Exception as e:
        _report(f"⚠️ Erro na inicialização: {e}")


if not os.getenv("PYTEST_CURRENT_TEST"):
    if AUTORUN:
        _executar_no_import()
