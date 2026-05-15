# 🔓 Excel Unlocker Pro v3.0

Aplicação desktop para gerenciamento e recuperação de proteções em arquivos Microsoft Excel (`.xlsx`, `.xlsm`, `.xls`).

---

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Como Usar](#como-usar)
- [Funções Disponíveis](#funções-disponíveis)
- [Arquivos Gerados](#arquivos-gerados)
- [Logs](#logs)
- [Limitações Técnicas](#limitações-técnicas)
- [Dependências](#dependências)

---

## Visão Geral

O **Excel Unlocker Pro** é uma ferramenta gráfica (CustomTkinter) capaz de identificar, analisar e remover diferentes tipos de proteção presentes em arquivos Excel. Funciona tanto para proteções de planilha/pasta de trabalho quanto para criptografia de arquivo completo.

> **Aviso**: Esta ferramenta destina-se exclusivamente à recuperação de arquivos próprios cujas senhas foram esquecidas. O uso em arquivos de terceiros sem autorização pode infringir leis de acesso não autorizado a dados.

---

## Requisitos

- Python **3.10** ou superior
- Sistema operacional: Windows, macOS ou Linux

---

## Instalação

**1. Clone ou baixe o arquivo `excel_unlocker.py`**

**2. Instale as dependências:**

```bash
pip install customtkinter msoffcrypto-tool openpyxl olefile
```

**3. Execute:**

```bash
python excel_unlocker.py
```

---

## Como Usar

1. **Selecione o arquivo Excel** clicando em "Selecionar" ao lado de "Arquivo Excel"
2. **Selecione a pasta de destino** onde a cópia desbloqueada será salva
3. **Escolha uma das 4 abas** conforme o tipo de operação desejada
4. Clique em **▶ Executar**
5. Acompanhe o progresso na barra e no log em tempo real
6. Para a Opção 4, use **⏹ Parar** a qualquer momento para interromper a busca

---

## Funções Disponíveis

### Opção 1 — Remover Senha

Remove todas as proteções de edição do arquivo.

- **Proteção de planilha** (`sheetProtection`): removida diretamente do XML interno
- **Proteção de pasta de trabalho** (`workbookProtection`): removida do `workbook.xml`
- **Arquivo criptografado** (com senha de abertura): informe a senha atual no campo disponível; o arquivo será decifrado e salvo sem proteção

**Arquivo gerado:** `NomeOriginal_SEM_SENHA.xlsx`

---

### Opção 2 — Trocar Senha

Localiza a senha atual e substitui pela nova.

- Para **proteção de planilha/pasta**: calcula o novo hash XOR 16-bit e substitui no XML
- Para **arquivo criptografado**: decifra com a senha atual e recifra com a nova senha (requer `msoffcrypto-tool`)
- Confirmação da nova senha obrigatória para evitar erros de digitação

**Arquivo gerado:** `NomeOriginal_NOVA_SENHA.xlsx`

---

### Opção 3 — LookPic (Encontrar Senha de Planilha)

Analisa o algoritmo de hash usado pelo Excel para proteger planilhas e tenta recuperar ou simular a senha por análise criptográfica.

**Como funciona:**

| Algoritmo detectado | Estratégia | Tempo estimado |
|---|---|---|
| Hash legado XOR 16-bit | Busca colisão exata no espaço de 65.536 valores | < 1 segundo |
| SHA-256 / SHA-512 moderno | Testa wordlist de ~1.200 senhas comuns + padrões numéricos | segundos a minutos |
| Sem hash reconhecido | Remove o elemento de proteção diretamente | < 1 segundo |

O log detalha o algoritmo encontrado, o valor do hash e o resultado de cada tentativa.

**Arquivo gerado:** `NomeOriginal_DESBLOQUEADO.xlsx`

---

### Opção 4 — Quebrar Senha de Abertura

Identifica o tipo de criptografia do arquivo, analisa todos os parâmetros criptográficos e testa combinações inteligentes até encontrar a senha que abre o arquivo.

**Etapa 1 — Análise de criptografia:**

Lê o stream `EncryptionInfo` do container OLE e identifica:

| Tipo | Algoritmo | Força |
|---|---|---|
| XOR/BIFF | XOR 16-bit | Muito fraca |
| Standard RC4 40-bit | RC4 + SHA-1 | Fraca |
| Standard RC4 128-bit | RC4 + SHA-1 | Moderada |
| Standard AES-128 | AES-128 ECB | Boa |
| Standard AES-192 | AES-192 ECB | Boa |
| **Agile AES-256** | **AES-256 CBC + PBKDF2-SHA512** | **Forte** |

Todos os parâmetros são registrados no log: algoritmo, hash, tamanho da chave em bits, tamanho do salt, spinCount (iterações PBKDF2) e provedor CSP.

**Etapa 2 — Ataque inteligente em ordem de probabilidade:**

```
1. Wordlist embutida (~1.200 entradas)
   ├── Senhas numéricas universais
   ├── Senhas em português (corporativas)
   ├── Senhas em inglês (comuns)
   ├── Padrões com anos (1990–2030)
   └── Padrões mês+ano (jan2023, fev2024...)

2. Wordlist externa (arquivo .txt, opcional)
   └── Uma senha por linha

3. Força bruta crescente (opcional, configurável)
   ├── Charset: dígitos / lower+dígitos / alfanumérico / estendido
   └── Comprimento: 1 até N caracteres (configurável na interface)
```

**Configurações da interface (Opção 4):**

| Campo | Opções | Descrição |
|---|---|---|
| Charset | digits / lower / alphanum / extended | Conjunto de caracteres para força bruta |
| Comprimento máximo | 0 (off) a 6 | Tamanho máximo na fase de força bruta |
| Wordlist externa | Arquivo .txt | Lista personalizada de senhas a testar |

**Etapa 3 — Quando a senha é encontrada:**
- O arquivo é decifrado completamente
- Proteções internas de planilha também são removidas
- O resultado é salvo na pasta de destino
- Um arquivo `.txt` paralelo registra a senha encontrada, número de tentativas e tempo decorrido

**Arquivo gerado:** `NomeOriginal_SENHA_ENCONTRADA.xlsx`  
**Registro adicional:** `NomeOriginal_SENHA_ENCONTRADA_SENHA.txt`

---

## Arquivos Gerados

O nome do arquivo original é sempre preservado. Apenas um sufixo é adicionado:

| Operação | Sufixo adicionado |
|---|---|
| Opção 1 — Remover senha | `_SEM_SENHA` |
| Opção 2 — Trocar senha | `_NOVA_SENHA` |
| Opção 3 — LookPic | `_DESBLOQUEADO` |
| Opção 4 — Senha encontrada | `_SENHA_ENCONTRADA` |

**Exemplo:** `Relatorio_Financeiro.xlsx` → `Relatorio_Financeiro_DESBLOQUEADO.xlsx`

---

## Logs

O log é salvo automaticamente na **mesma pasta do script**, com nome no formato:

```
excel_unlocker_YYYYMMDD.log
```

O log registra:
- Data e hora de cada operação
- Tipo de proteção/criptografia detectado
- Valores de hash encontrados nos XMLs
- Número de tentativas realizadas
- Velocidade de tentativas por segundo
- Senha encontrada (quando aplicável)
- Erros e seus motivos detalhados
- Caminho completo do arquivo salvo

---

## Limitações Técnicas

### Proteção de planilha (Opções 1, 2 e 3)
- Hash legado XOR 16-bit: espaço de apenas 65.536 valores — colisão garantida em menos de 1 segundo
- Hash SHA-512 com spinCount alto (~100.000 iterações): cada tentativa leva ~200–500 ms em Python puro

### Criptografia de arquivo (Opção 4)
- **AES-256 + PBKDF2 com 100.000 iterações**: Python puro consegue aproximadamente 1–5 tentativas por segundo
- Senhas curtas e comuns (até 6 dígitos, padrões corporativos) são cobertas pela wordlist e pela força bruta configurável
- Senhas longas e aleatórias com caracteres especiais estão fora do alcance prático desta ferramenta em Python

**Para ataques de alta performance em AES-256**, ferramentas especializadas com suporte a GPU são mais adequadas:

```bash
# Hashcat (requer extração de hash com office2john)
hashcat -m 9500 hash.txt wordlist.txt   # Office 2013+
hashcat -m 9400 hash.txt wordlist.txt   # Office 2010

# John the Ripper
john --format=office hash.txt
```

---

## Dependências

| Pacote | Versão mínima | Uso |
|---|---|---|
| `customtkinter` | 5.x | Interface gráfica moderna |
| `msoffcrypto-tool` | 5.x | Decifrar/recifrar arquivos criptografados |
| `openpyxl` | 3.x | Leitura opcional de estrutura Excel |
| `olefile` | 0.46+ | Leitura do stream EncryptionInfo (OLE/CFB) |

Todas são instaláveis via pip:

```bash
pip install customtkinter msoffcrypto-tool openpyxl olefile
```

A aplicação detecta automaticamente quais dependências estão ausentes e exibe um aviso na interface.

---

## Estrutura do Projeto

```
excel_unlocker.py        ← Arquivo principal (executar este)
README.md                ← Este arquivo
excel_unlocker_YYYYMMDD.log  ← Gerado automaticamente ao executar
```

---

*Excel Unlocker Pro v3.0 — desenvolvido para recuperação de arquivos próprios com senhas esquecidas.*

O formato é bem simples:

**Arquivo de texto puro (`.txt`), uma senha por linha.**

```
senha123
1234
financeiro2024
MinhaEmpresa@2023
relatorio
```

Regras:
- Codificação: **UTF-8** (aceita acentos e caracteres especiais)
- Uma senha por linha
- Linhas em branco são ignoradas automaticamente
- Espaços no início/fim de cada linha são removidos (`strip`)
- Sem limite de tamanho de arquivo — ele é lido linha a linha, sem carregar tudo na memória

**Fontes populares de wordlists prontas para usar:**
- `rockyou.txt` — lista clássica com senhas reais vazadas, amplamente usada para recuperação
- `SecLists` (GitHub: danielmiessler/SecLists) — coleções organizadas por tema
- Você mesmo pode criar uma com senhas que lembra ter usado no arquivo
