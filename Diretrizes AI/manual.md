---
noteId: "f33b104050ab11f19150278c42fa9df2"
tags: []

---

O Qwen Code lê automaticamente arquivos específicos na raiz do projeto. Veja as opções:

**Na raiz do projeto (recomendado):**
- `QWEN.md` - diretrizes gerais para o assistente
- `AGENTS.md` - alternativa para diretrizes

**Dentro de pastas específicas:**
- `.qwen/QWEN.md` - também é carregado automaticamente
- `.qwen/memory/` - para memórias persistentes (como descrito no sistema)

Infelizmente, não há suporte para colocar `QWEN.md` em subpastas arbitrárias - precisa estar na raiz ou dentro de `.qwen/`.

Para inserir diretrizes de contexto no Qwen Code, siga estas etapas:

1. **Crie ou edite o arquivo `QWEN.md` na raiz do projeto**  
   Esse arquivo é lido automaticamente pelo assistente e contém as regras que ele deve obedecer durante a sessão.

2. **Estruture o conteúdo em seções claras**  
   Use cabeçalhos (`#`, `##`, etc.) para separar as áreas de contexto, por exemplo:  
   ```markdown
   # Contexto do Projeto
   - Descrição geral da aplicação  
   - Tecnologias usadas (Python 3.11, FastAPI, etc.)  

   # Diretrizes de Código
   - Padrões de nomenclatura (snake_case)  
   - Linters e formatação (ruff, black)  

   # Fluxo de Trabalho
   - Como devem ser criados os commits  
   - Quando executar testes (`pytest`)  

   # Política de Segurança
   - Não expor segredos  
   - Validar todas as entradas de usuário
   ```

3. **Adicione configurações específicas de linguagem**  
   Se o projeto usa ferramentas como `pyproject.toml`, inclua-as no `QWEN.md` para que o assistente saiba onde encontrar as configurações de lint, dependências, etc.

4. **Salve o arquivo**  
   Depois de gravar `QWEN.md`, o assistente recarregará automaticamente as novas diretrizes na próxima interação.

5. **Verifique o carregamento**  
   Você pode pedir ao assistente para listar as diretrizes atuais:  
   ```
   /qc-helper mostre as diretrizes de contexto
   ```




Com essas etapas, você define o que o Qwen Code deve considerar ao ajudar no desenvolvimento.


Os arquivos úteis no diretó rio `.qwen` são:

- `.qwen/QWEN.md` – diretrizes gerais e regras críticas  
- `.qwen/manual.md` – instruções auxiliares  
- `.qwen/settings.json` – configurações de permissões  
- `.qwen/memory/` – memórias persistentes (ex.: arquivos *.md* salvos)  
- `.qwen/direitivas/` – módulos detalhados para UI, logs, dependências, arquitetura, etc.  

Todos os arquivos listados são lidos automaticamente pelo assistente para orientOs arquivos úteis no diretório `.qwen` são:

- `.qwen/QWEN.md` – diretrizes gerais e regras críticas  
- `.qwen/manual.md` – instruções auxiliares  
- `.qwen/settings.json` – configurações de permissões  
- `.qwen/memory/` – memórias persistentes (ex.: arquivos *.md*)  
- `.qwen/direitivas/` – módulos detalhados para UI, logs, dependências, arquitetura, etc.  

Todos são lidos automaticamente pelo assistente para orientar o desenvolvimento.  
Ajuste-os conforme precisar; mudanças em `.md` requerem recarregar a sessão.

# Skils

# 📘 Manual: O que são Skills para IA e Como Criar para o Qwen (Terminal)

> ⚠️ **Nota importante:** "Skills" não são um recurso nativo ou plugin oficial do Qwen. Elas são um **padrão de engenharia de prompt + orquestração** que transforma um modelo de linguagem em um assistente especializado. Este manual ensina a criá-las de forma prática, segura e reproduzível para a versão terminal do Qwen.

---

## 🎯 Objetivo
- Entender o conceito de "skills" em IA generativa.
- Aprender a estruturar, salvar e executar skills para o Qwen via terminal.
- Fornecer templates prontos e boas práticas para uso diário.

---

## 🔍 O que são "Skills" para IA?

Uma **skill** (ou "habilidade") é um conjunto de instruções, exemplos e regras que direcionam o modelo a realizar uma tarefa específica com consistência. Pense nela como um **"modo especializado"** que você ativa antes de fazer uma pergunta.

### ✅ Exemplos comuns de skills:
| Skill | Objetivo |
|-------|----------|
| `revisor_python` | Revisar código, sugerir melhorias e corrigir bugs |
| `tradutor_tecnico` | Traduzir documentação mantendo terminologia específica |
| `analista_logs` | Extrair padrões, erros e métricas de logs brutos |
| `gerador_sql` | Criar consultas otimizadas a partir de descrições em linguagem natural |

### 🧠 Como funciona na prática?
1. O modelo **não aprende permanentemente**. A skill só vale durante a sessão/contexto atual.
2. Funciona via **system prompt** + **exemplos (few-shot)** + **formatação de saída**.
3. Pode ser combinada com **ferramentas externas** (scripts, APIs, bancos de dados) se o ambiente suportar *function/tool calling*.

---

## ⚙️ Pré-requisitos para o Terminal

- Qwen rodando localmente ou via API (ex: `ollama`, `qwen-cli`, `vLLM`, ou script Python com `transformers`)
- Terminal com acesso a arquivos `.txt`/`.md`
- Editor de texto (VS Code, nano, vim, etc.)

> 💡 Este manual usa `ollama` como exemplo, mas a lógica é idêntica para qualquer wrapper terminal do Qwen.

---

## 🛠️ Passo a Passo: Criando uma Skill

### 1️⃣ Estrutura de Pastas
```bash
~/qwen-skills/
├── skills/
│   ├── revisor_python.txt
│   ├── tradutor_tecnico.txt
│   └── analista_logs.txt
└── run_skill.sh
```

### 2️⃣ Template da Skill (Arquivo `.txt`)
Cada skill é um arquivo de texto com **apenas instruções**. Não coloque perguntas nele.

**Exemplo: `skills/revisor_python.txt`**
```text
Você é um revisor sênior de código Python. Siga estas regras:
1. Analise apenas o código enviado pelo usuário.
2. Liste os problemas encontrados, classifique por severidade (CRÍTICO, MÉDIO, BAIXO).
3. Forneça o código corrigido em um bloco ```python.
4. Mantenha respostas concisas. Não explique conceitos básicos.
5. Se o código estiver correto, responda apenas: ✅ Código adequado. Nenhuma alteração necessária.

Formato de saída obrigatório:
### 🐛 Problemas
- ...

### 🛠️ Código Corrigido
```python
...
```
```

### 3️⃣ Script de Execução (`run_skill.sh`)
```bash
#!/bin/bash
# Uso: ./run_skill.sh <nome_da_skill> "sua pergunta ou entrada"

if [ $# -lt 2 ]; then
  echo "❌ Uso: $0 <skill> <entrada>"
  echo "Ex: $0 revisor_python \"def soma(a,b): return a+b\""
  exit 1
fi

SKILL_NAME="$1"
USER_INPUT="$2"
SKILL_FILE="skills/${SKILL_NAME}.txt"

if [ ! -f "$SKILL_FILE" ]; then
  echo "❌ Skill não encontrada: $SKILL_FILE"
  exit 1
fi

# Monta o prompt final: System + Entrada do Usuário
FULL_PROMPT="$(cat "$SKILL_FILE")

---
ENTRADA DO USUÁRIO:
$USER_INPUT"

# Executa com Ollama (ajuste o modelo conforme sua instalação)
echo "🔄 Processando com skill: $SKILL_NAME ..."
ollama run qwen2.5:7b "$FULL_PROMPT"
```

Torne executável:
```bash
chmod +x run_skill.sh
```

### 4️⃣ Testando
```bash
./run_skill.sh revisor_python "def divide(a,b): return a/b"
```

---

## 📦 Skill com Chamada de Ferramenta (Avançado)

Se você usa uma versão do Qwen com suporte a **tool/function calling** (ex: Qwen2.5/3 via API OpenAI-compatible), pode integrar scripts externos:

1. Defina a função em JSON (ex: `buscar_dados_banco`)
2. Envie junto ao prompt no parâmetro `tools`
3. O Qwen decidirá quando chamar a ferramenta e retornará um JSON estruturado
4. Seu script terminal executa a ação e devolve o resultado ao modelo

> 🔗 Referência: [Qwen Function Calling Docs](https://qwen.readthedocs.io/)

---

## ✅ Boas Práticas

| ✅ Faça | ❌ Evite |
|--------|---------|
| Usar delimitadores claros (`###`, ```, `---`) | Misturar instruções com a pergunta do usuário |
| Limitar a 150-300 tokens por skill | Prompts genéricos como "seja útil" |
| Incluir formato de saída obrigatório | Assumir que o modelo "lembrará" entre sessões |
| Testar com 3-5 exemplos antes de usar | Skills muito longas ou com regras conflitantes |

---

## 🚨 Solução de Problemas

| Sintoma | Causa Provável | Solução |
|--------|----------------|---------|
| Modelo ignora a skill | Prompt muito longo ou mal posicionado | Coloque a skill no início, use `System:` explícito |
| Saída inconsistente | Falta de formato obrigatório | Adicione um template de resposta no final do arquivo |
| Erro no script | Permissões ou caminhos incorretos | `chmod +x`, verifique `pwd` e nomes de arquivos |
| Respostas lentas | Contexto muito grande ou modelo pesado | Use versão `1.5b`/`3b` para tarefas simples, reduza exemplos |

---

## 📝 Mantendo suas Skills

- Versione no Git: `git init && git add skills/`
- Use nomes descritivos: `analise_contrato_v2.txt`, não `skill1.txt`
- Comente alterações no início do arquivo: `# v1.2 - 2026-05-10 | Adicionado suporte a JSON`
- Faça backup semanal da pasta `skills/`

---

## 📚 Recursos Oficiais

- 🌐 Documentação Qwen: https://qwen.readthedocs.io/
- 📦 Modelos Qwen no Ollama: https://ollama.com/library/qwen2.5
- 🛠️ Guia de Prompt Engineering: https://platform.openai.com/docs/guides/prompt-engineering
- 💬 Comunidade: https://github.com/QwenLM/Qwen

---

> ℹ️ **Lembre-se:** Skills são *padrões de uso*, não atualizações do modelo. Elas funcionam porque o Qwen foi treinado para seguir instruções contextuais com alta fidelidade. Quanto mais clara e estruturada for a skill, mais previsível será o resultado.

🖨️ *Copie este conteúdo para `manual.md` e comece a criar suas skills hoje.*