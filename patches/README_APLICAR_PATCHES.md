# Como aplicar patches automaticamente (Windows/macOS/Linux)

## Opção A — usando Git (recomendado)
1. Abra um terminal na raiz do projeto (pasta que contém a pasta `utils/`).
2. Garanta que o projeto está versionado:
   ```bash
   git init
   git add .
   git commit -m "estado atual"
   ```
3. Aplique os patches em ordem:
   ```bash
   git apply --reject --whitespace=fix patches/0001-dataframe_utils.patch
   git apply --reject --whitespace=fix patches/0002-comparacao.patch
   git apply --reject --whitespace=fix patches/0003-formatadores.patch
   git apply --reject --whitespace=fix patches/0004-relatorios-anticircular.patch
   ```
4. Se algum hunk falhar, o Git criará arquivos `.rej`. Abra-os e aplique manualmente só aquele trecho.

## Opção B — sem Git (usa `patch`)
- Windows: instale via Chocolatey `choco install patch` ou baixe o GnuWin32.
- macOS: já vem o `patch`. Linux: já vem também.
Aplique assim:
```bash
patch -p1 < patches/0001-dataframe_utils.patch
patch -p1 < patches/0002-comparacao.patch
patch -p1 < patches/0003-formatadores.patch
patch -p1 < patches/0004-relatorios-anticircular.patch
```

## Dica
- Faça um backup antes: `git tag before-patches` ou copie a pasta.
- Esses patches **não** removem funções existentes; só acrescentam/ajustam linhas.