# 🛠️ Guia Definitivo de Compilação com Nuitka

Este documento descreve as etapas necessárias para compilar todo o ecosistema (Pool de Aplicações) do Prometeus usando o **Nuitka**. O projeto foi completamente revisado e atualizado para oferecer 100% de compatibilidade tanto na versão código-fonte (`.py`) quanto na versão compilada executável (`.exe`).

## 1. O que foi ajustado na arquitetura para o Nuitka?
- **Descoberta Dinâmica de Módulos:** O arquivo `main.py` foi atualizado para, inteligentemente, executar arquivos `.exe` caso esteja operando em formato binário (`frozen`), ou arquivos `.py` caso esteja operando via interpretador clássico.
- **Paths Relativos Seguros:** Todos os scripts usam a arquitetura `BASE_DIR`, que resolve de maneira segura a raiz do contexto, impedindo erros ao carregar `Grupo3.png` ou qualquer outro `asset` durante a execução congelada.
- **Isolamento de Processos via Subprocess:** Quando executados via `main.exe`, os aplicativos herdam de forma correta as credenciais via varíavel de ambiente (`PROMETEUS_USER`), sem travar o interpretador, por intermédio da lib `subprocess` apontando diretamente para as sub-aplicações.

---

## 2. Pré-requisitos de Compilação
Antes de iniciar, certifique-se de que o gerenciador e as ferramentas de compilação C estão instaladas e atualizadas em sua máquina (MinGW64 ou MSVC) e instale o Nuitka no ecossistema atual:

```bash
uv pip install -r requirements.txt
uv pip install nuitka zstandard
```

---

## 3. Passo a Passo de Compilação do Sistema

Como o Prometeus opera como um "Pool de Aplicações", cada aplicação do ecossistema e o próprio `main.py` devem ser compilados para gerar binários independentes e coesos (reduzindo conflitos de memória e de concorrência).

### Fase 1: Compilação das Aplicações (Worker Scripts)
Para cada script auxiliar (ex: `AltCust.py`, `Mephisto.py`, `conversor_xmls.py`), execute o comando na raiz apontando para o seu diretório:

**Exemplo para o `AltCust.py`:**
```bash
uv run python -m nuitka \
  --standalone \
  --onefile \
  --disable-console \
  --plugin-enable=tk-inter \
  --include-package-data=customtkinter \
  --output-dir=build \
  App.Gemco\AlCusto\AltCust.py
```
> **Nota:** Repita este comando para todos os arquivos `.py` listados nas subpastas (`App.Gemco`, `Ar.Excel`, `NF_CTE`).

### Fase 2: Compilação do `main.py` (Controlador)
O controlador principal requer o ícone personalizado. Execute:

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

---

## 4. Estrutura Pós-Compilação (O que distribuir para o Cliente/Usuário final)

O Nuitka vai gerar os arquivos compilados na pasta `build\`. Para que o sistema opere corretamente e encontre as ferramentas dentro da interface, a pasta que será entregue para uso deverá possuir a **exata mesma árvore de pastas**. 

Monte o pacote final da seguinte forma:

```text
📦 Prometeus_Release_v1.0
 ┣ 📜 main.exe                    # O executável do seu hub criado na Fase 2
 ┣ 📂 logs/                       # Onde os logs serão salvos (unificados)
 ┣ 📂 App.Gemco/
 ┃  ┣ 📂 AlCusto/
 ┃  ┃  ┣ 📜 AltCust.exe           # Executável gerado na Fase 1
 ┃  ┃  ┣ 🖼️ Grupo3.png            # Arquivos de automação da GUI necessitam estar ao lado do seu script!
 ┃  ┃  ┣ 🖼️ Alteracao.png
 ┃  ┃  ┗ 🖼️ ...
 ┃  ┣ 📂 DIST/
 ┃  ┃  ┗ 📜 DIST.3.0.exe
 ┃  ┗ 📂 Mephisto/
 ┃     ┗ 📜 Mephisto.exe
 ┣ 📂 Ar.Excel/
 ┃  ┣ 📜 Eexcel_Unlocker Pro.exe
 ┃  ┗ 📜 Excel_Protector_v3.exe
 ┗ 📂 NF_CTE/
    ┣ 📜 conversor_xmls.exe
    ┣ 📜 CTe_RelXml.exe
    ┗ 📜 NFe.Excel_3.0.exe
```

### Regras Ouro Para a Entrega
1. A extensão nos subdiretórios **deve** ser `.exe`. O `main.exe` foi reprogramado para ler automaticamente e varrer procurando arquivos `.exe` caso você distribua em modo binário. 
2. Toda e qualquer Imagem que um script do "pool" precise caçar no Excel (`.png` e `.ico`) **NÃO VAI JUNTO DA COMPILAÇÃO DO NUITKA**. Você é obrigado a copiá-las manualmente para dentro da mesma sub-pasta onde está o arquivo `.exe` resultante dele.
