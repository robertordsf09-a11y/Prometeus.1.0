# Conversor de NFe para Excel (Modelo 55)

## Pré-requisitos
- Python `>= 3.11`
- Gerenciador de dependências `uv`

## Instalação e Configuração

Crie o ambiente virtual e instale as dependências usando `uv`:
```bash
uv venv
uv pip install customtkinter pandas openpyxl
```

## Execução

Para iniciar a aplicação a partir dos scripts localmente:
```bash
uv run python NFe.Excel_3.0.py
```

## Compilação com Nuitka

O projeto foi estruturado para suportar a compilação garantida usando o Nuitka para geração de executáveis nativos do Windows.

Comando de compilação:
```bash
uv run python -m nuitka \
  --standalone \
  --onefile \
  --disable-console \
  --plugin-enable=tk-inter \
  --windows-uac-admin \
  --output-dir=build \
  NFe.Excel_3.0.py
```

*Nota: Todos os caminhos de persistência e geração de logs já contemplam o escopo `sys.frozen` para a garantia de uso standalone.*
