# 🔐 Excel Password Protector v2

Aplicativo com interface gráfica moderna para proteger planilhas Excel com **criptografia real ECMA-376**.

## 📦 Instalação

```bash
pip install customtkinter msoffcrypto-tool Pillow
```

## ▶️ Execução

```bash
python excel_protector.py
```

## 🖥️ Como usar

1. **Selecione a pasta de origem** → clique em "Selecionar Origem"
2. **Selecione a pasta de destino** → onde a nova pasta será criada
3. **Digite a senha** (mínimo 4 caracteres)
4. **Clique em "Escanear Arquivos"** → o app exibe todos os `.xlsx / .xlsm / .xls / .xlsb` organizados por pasta
5. **Marque/desmarque** arquivos individualmente ou pela pasta inteira (checkbox roxo)
   - Clique em ▼ / ▶ para colapsar/expandir uma pasta
6. **Clique em "Processar e Proteger"** → o app:
   - Cria `Protegido_<nome>_<timestamp>/` no destino
   - Replica toda a estrutura de subpastas
   - ✅ Arquivos marcados → criptografados com senha ECMA-376
   - ☐ Arquivos desmarcados → apenas copiados
   - 📋 Gera log `.log` na pasta de destino

## 📋 Arquivo de Log

Gerado automaticamente em `Protegido_.../relatorio_<timestamp>.log`:

```
2025-04-25 14:30:01  [INFO    ]  Origem  : C:/Documentos/Planilhas
2025-04-25 14:30:01  [INFO    ]  Total de arquivos : 12
2025-04-25 14:30:01  [INFO    ]  Marcados para protecao : 8
2025-04-25 14:30:02  [INFO    ]  [1/12] PROTEGER  financeiro/jan.xlsx
2025-04-25 14:30:02  [INFO    ]    OK  Criptografado e verificado (metodo: encrypt_direto)
2025-04-25 14:30:03  [ERROR   ]  [3/12] PROTEGER  rh/folha.xls
2025-04-25 14:30:03  [ERROR   ]    ERRO  Falha ao proteger: ...
...
2025-04-25 14:30:10  [INFO    ]  RESUMO
2025-04-25 14:30:10  [INFO    ]    Protegidos com sucesso : 7
2025-04-25 14:30:10  [INFO    ]    Falhas de protecao     : 1
2025-04-25 14:30:10  [INFO    ]    Copiados sem senha     : 4
2025-04-25 14:30:10  [INFO    ]    TOTAL DE ERROS         : 1
```

## 🔒 Lógica de Proteção

A proteção usa **duas tentativas** para garantir robustez:

1. `msoffcrypto.encrypt()` direto no arquivo de saída
2. `msoffcrypto.encrypt()` via buffer em memória (fallback)

Após cada proteção, o arquivo é **verificado** abrindo-o com a senha informada. Se a verificação falhar, o log registra o aviso. Em último caso, o arquivo é copiado sem senha e o erro é registrado.

## ⚠️ Observações

- Arquivos **já protegidos** com senha serão copiados sem modificação (registrado no log)
- A barra de progresso fica **vermelha** se houver erros, **verde** se tudo OK
- O caminho do log é exibido na interface após o processamento
