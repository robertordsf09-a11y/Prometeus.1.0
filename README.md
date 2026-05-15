# Instalação e Configuração — Prometeus Automation

Este documento descreve os procedimentos para configurar o ambiente de desenvolvimento e gerar o executável de produção para o sistema `Prometeus1.0`.

## 1. Pré-requisitos

- **Python:** >= 3.11
- **Gerenciador de Pacotes:** `uv` (Recomendado)
- **Sistema Operacional:** Windows 10/11

## 2. Configuração do Ambiente (via `uv`)

Recomendamos o uso do `uv` para garantir a reprodutibilidade das dependências e isolamento do ambiente.

```bash
# Criar ambiente virtual
uv venv

# Ativar ambiente (Windows)
.venv\Scripts\activate

# Instalar dependências
uv pip install customtkinter Pillow
```

## 3. Execução em Desenvolvimento

Para iniciar a aplicação diretamente:

```bash
uv run python main.py
```

## 4. Compilação para Produção (Nuitka)

O projeto está otimizado para compilação via Nuitka, gerando um único executável autônomo.

```bash
uv run python -m nuitka \
  --standalone \
  --onefile \
  --disable-console \
  --plugin-enable=tk-inter \
  --include-package-data=customtkinter \
  --windows-icon-from-ico=gothic.ico \
  --windows-uac-admin \
  --output-dir=build \
  main.py
```

### Notas de Compilação:
- O uso de `BASE_DIR` garante que subpastas como `logs/` e diretórios de scripts sejam localizados corretamente em relação ao executável.
- O flag `--windows-uac-admin` é necessário se as automações precisarem de privilégios elevados.
- Recursos externos (ícones, imagens) devem ser mantidos na pasta `assets/` ao lado do código ou embutidos via `--include-data-files`.

## 5. Estrutura de Pastas Esperada

Para o pleno funcionamento do launcher, organize seus scripts Python nas seguintes pastas raiz (que serão criadas automaticamente se não existirem):

- `App.Gemco/`
- `Ar.Excel/`
- `NF_CTE/`
- `Utilitários/`

---
*Roberto Santos [LABS]©*
