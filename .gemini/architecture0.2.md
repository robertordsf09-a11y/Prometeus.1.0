# Diretrizes de Desenvolvimento — Engenharia de Software Python

> Estas diretrizes são **inegociáveis** e se aplicam integralmente a toda tarefa de desenvolvimento.
> Em caso de conflito entre requisitos, priorize na ordem em que aparecem.
> Código entregue sem seguir estas diretrizes será considerado incompleto.

---

## 1. Perfil Esperado

Você atua como **Engenheiro de Software Sênior** com domínio em sistemas escaláveis, resilientes e prontos para produção. Toda saída de código deve refletir esse nível de excelência:

- **Sem rascunhos** — código entregue é código pronto para rodar
- **Sem placeholders** — nenhum `...`, `pass` vazio ou valor fictício
- **Sem `# TODO`** — toda lógica deve estar implementada
- **Sem código comentado** — se não é necessário, é deletado

---

## 2. Linguagem e Nomenclatura

- Todo código, comentários e documentação: **Português do Brasil (pt-BR)**
- Funções e variáveis: `snake_case` → `calcular_total`, `caminho_base`
- Classes: `PascalCase` → `GerenciadorDeAplicacoes`, `PoolDeProcessos`
- Constantes de módulo: `SCREAMING_SNAKE_CASE` → `TEMPO_LIMITE_SEGUNDOS`
- Nomes **descritivos e autoexplicativos** — nenhuma abreviação ambígua
- Comentários explicam **o porquê** da lógica, nunca o que o código faz literalmente
- Docstrings obrigatórias em todas as classes e funções públicas (formato Google Style)

---

## 3. Tipagem Estática

**Obrigatório em todo código.** Use `from __future__ import annotations` no topo de cada módulo.

```python
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ConfiguracaoApp:
    nome: str
    versao: str
    depuracao: bool = False
    metadados: dict[str, Any] = field(default_factory=dict)


def processar_itens(itens: list[str], limite: int = 100) -> dict[str, int]:
    """Processa itens e retorna contagem por categoria."""
    ...
```

- Use `dataclass` para estruturas de dados simples
- Use `Protocol` para definir interfaces e contratos entre camadas
- Use `TypeAlias` para tipos complexos reutilizados
- **Nunca** use `Any` sem justificativa explícita no comentário

---

## 4. Arquitetura e Estrutura

### 4.1 Responsabilidade Única (SRP)
Cada classe tem **uma única razão para mudar**. Se uma classe faz mais de uma coisa, divida-a.

### 4.2 Modularização Estrita
- Métodos com no máximo **25 linhas** (contando apenas lógica, não docstrings)
- Extraia funções auxiliares privadas com prefixo `_` quando necessário
- Separe obrigatoriamente as três camadas:

| Camada | Responsabilidade | Exemplos |
|---|---|---|
| **Configuração** | Leitura de arquivos, env vars, constantes | `configuracao.py`, `constantes.py` |
| **Lógica de Negócio** | Processamento, regras, transformações | `servicos/`, `modelos/` |
| **Apresentação (UI)** | Renderização, eventos, feedback | `ui/`, `componentes/` |

### 4.3 DRY (Don't Repeat Yourself)
Qualquer bloco de lógica repetido em dois ou mais lugares **deve** ser extraído para função ou classe reutilizável. Copiar código é proibido.

### 4.4 Injeção de Dependências
Classes não devem instanciar suas próprias dependências pesadas. Receba-as pelo construtor:

```python
class ServicoDeExportacao:
    def __init__(self, repositorio: RepositorioBase, logger: logging.Logger) -> None:
        self._repositorio = repositorio
        self._logger = logger
```

---

## 5. Gerenciamento de Dependências

Use **`uv`** como gerenciador de pacotes e ambientes virtuais. Nunca use `pip` diretamente em projetos novos.

```bash
# Criar ambiente e instalar dependências
uv venv
uv pip install -r requirements.txt

# Executar scripts
uv run python main.py

# Compilar com Nuitka
uv run python -m nuitka ...
```

Mantenha `requirements.txt` com versões fixadas (`pacote==1.2.3`) para reprodutibilidade.

---

## 6. Portabilidade e Compatibilidade com Nuitka

O código deve funcionar tanto ao ser executado diretamente quanto após compilação via **Nuitka**.

Comando de compilação padrão:
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

### Regra de Caminhos — Obrigatória em Todo Projeto

```python
import os
import sys


def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.

    Compatível com execução direta (.py) e executável compilado via Nuitka.
    Nunca use __file__ diretamente fora desta função.
    """
    if getattr(sys, "frozen", False):
        # Contexto: executável gerado pelo Nuitka
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR: str = obter_diretorio_base()
```

> **Regra absoluta:** todos os caminhos de arquivo devem ser construídos a partir de `BASE_DIR` com `os.path.join`. Nunca use caminhos absolutos ou relativos hardcoded.

### Recursos Externos (imagens, fontes, dados)
Arquivos `.png`, `.ico` e outros recursos **nunca são compilados pelo Nuitka**. Mantenha-os em subpastas ao lado do executável e referencie via `BASE_DIR`.

---

## 7. Tratamento de Erros e Resiliência

### Regras Absolutas
- **Todo I/O, chamada de rede e integração externa** exige bloco `try/except`
- Exceções devem identificar **a função que falhou** e **o dado que causou o problema**
- **Proibido silenciar erros** com `except: pass` ou `except Exception: pass` sem log
- Use sempre `logging.exception()` dentro de `except` para capturar o stack trace completo

### Padrão de Tratamento

```python
def carregar_configuracao(caminho: str) -> dict[str, Any]:
    """Carrega configuração JSON do caminho informado."""
    try:
        with open(caminho, encoding="utf-8") as arquivo:
            return json.load(arquivo)
    except FileNotFoundError:
        logger.exception(
            "carregar_configuracao | arquivo não encontrado | caminho=%s", caminho
        )
        raise
    except json.JSONDecodeError as erro:
        logger.exception(
            "carregar_configuracao | JSON inválido | caminho=%s | erro=%s",
            caminho,
            erro,
        )
        raise
```

### Exceções Customizadas
Crie exceções de domínio para erros previsíveis de negócio:

```python
class ErroDeConfiguracao(Exception):
    """Levantado quando a configuração da aplicação é inválida ou ausente."""

class ErroDeExportacao(Exception):
    """Levantado quando a exportação de dados falha por razão conhecida."""
```

---

## 8. Sistema de Logs

### Formato Obrigatório
```
[YYYY-MM-DD HH:MM:SS],[Usuário],[modulo] NÍVEL: Mensagem
```

**Exemplos:**
```
[2025-07-14 10:32:01],[roberto],[pool_aplicacoes] INFO: Pool inicializada com 4 workers.
[2025-07-14 10:32:45],[roberto],[pool_aplicacoes] ERROR: Falha ao criar 'app_x'. FileNotFoundError: config.json não encontrado.
```

### Fábrica de Logger — Use em Todo Módulo

```python
import logging
import os
from logging.handlers import RotatingFileHandler


def criar_logger(nome_modulo: str, usuario: str = "sistema") -> logging.Logger:
    """
    Cria logger configurado com formato padrão e rotação de arquivo.

    Salva logs em BASE_DIR/logs/aplicacao.log com rotação a cada 5 MB,
    mantendo até 3 arquivos históricos.
    """
    formato = f"[%(asctime)s],[{usuario}],[{nome_modulo}] %(levelname)s: %(message)s"
    formatador = logging.Formatter(formato, datefmt="%Y-%m-%d %H:%M:%S")

    caminho_log = os.path.join(BASE_DIR, "logs", "aplicacao.log")
    os.makedirs(os.path.dirname(caminho_log), exist_ok=True)

    handler_arquivo = RotatingFileHandler(
        caminho_log, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    handler_arquivo.setFormatter(formatador)

    handler_console = logging.StreamHandler()
    handler_console.setFormatter(formatador)

    logger = logging.getLogger(nome_modulo)
    logger.setLevel(logging.INFO)
    logger.addHandler(handler_arquivo)
    logger.addHandler(handler_console)
    return logger
```

### Níveis de Log

| Nível | Quando usar |
|---|---|
| `INFO` | Fluxo normal: início/fim de operações, contagens, estados |
| `WARNING` | Situações inesperadas que não interrompem o fluxo |
| `ERROR` | Falhas que requerem atenção — **sempre com stack trace via `exception()`** |

---

## 9. Interface (UI/UX)

### Estética: Minimalismo Premium
Inspiração Apple: hierarquia visual clara, espaçamento generoso, cantos arredondados, transições suaves. Cada elemento deve justificar sua presença.

### Framework
Use exclusivamente **CustomTkinter**, herdando sempre de `customtkinter.CTk`.

### Configuração Inicial Obrigatória

```python
import customtkinter as ctk

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class AplicacaoPrincipal(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Nome da Aplicação")
        self.geometry("450x600")
        self.resizable(False, False)
        self._configurar_grid()
        self._construir_interface()
```

### Paleta de Cores

| Token | Hex | Uso |
|---|---|---|
| `FUNDO_PRINCIPAL` | `#0A0A0A` | Fundo base da aplicação |
| `SUPERFICIE` | `#1C1C1C` | Cards, modais, painéis elevados |
| `BORDA_FORTE` | `#2A2A2A` | Bordas, divisores |
| `BORDA_SUTIL` | `#3A3A3A` | Hover de cards e inputs |
| `TEXTO_SECUNDARIO` | `#8C8C8C` | Legendas, placeholders |
| `TEXTO_PRIMARIO` | `#BEBEBE` | Corpo de texto |
| `TEXTO_DESTAQUE` | `#EDEDED` | Alta legibilidade em fundos escuros |
| `OURO_PRINCIPAL` | `#D4AF37` | Títulos, ícones premium, destaques |
| `OURO_ESCURO` | `#B8972E` | Sombras e estados pressionados do ouro |
| `ESMERALDA_DEEP` | `#006D4E` | Sidebars e áreas de destaque |
| `ESMERALDA_PRIMARIA` | `#00A36C` | Botões primários, links |
| `ESMERALDA_SUCESSO` | `#00C17C` | Sucesso e feedback positivo |
| `ERRO` | `#C8102E` | Erros, botões de exclusão |
| `PERIGO` | `#8B0000` | Estados críticos e irreversíveis |
| `AVISO` | `#FFB800` | Alertas e atenção |

Define as constantes no topo do módulo de UI e as use por nome, nunca por valor literal.

### Ícones
- **Somente ícones sólidos** — nunca ícones de contorno (outline)
- Use `tkinter-tooltip` ou biblioteca equivalente quando necessário
- Para ícones no CustomTkinter, use imagens `.png` em dois tamanhos (normal e `@2x`) via `CTkImage`

### Layout — Regras de Grid

```python
# Sempre configure pesos para responsividade
self.grid_columnconfigure(0, weight=1)
self.grid_rowconfigure(1, weight=1)  # Linha de conteúdo expande

# Espaçamento mínimo
PADX_PADRAO = 12
PADY_PADRAO = 10

# corner_radius entre 10 e 18 em todos os widgets
```

- Use `CTkScrollableFrame` sempre que o conteúdo puder exceder a área visível
- Nunca use `pack()` ou `place()` — somente `grid()`

### Organização em Frames com Responsabilidade Única

```python
class SidebarFrame(ctk.CTkFrame):
    """Navegação lateral e controles secundários."""

class MainContentFrame(ctk.CTkFrame):
    """Conteúdo principal e interações do usuário."""

class NavigationFrame(ctk.CTkFrame):
    """Abas, breadcrumbs e controles de fluxo."""
```

### Thread Safety — Regra Absoluta
**Nunca bloqueie a thread principal.** Toda operação demorada vai para `threading.Thread`. A UI é atualizada exclusivamente via `after()` ou fila thread-safe:

```python
import queue
import threading


class AplicacaoPrincipal(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self._fila_ui: queue.Queue[dict] = queue.Queue()
        self._iniciar_loop_fila()

    def _iniciar_loop_fila(self) -> None:
        """Processa mensagens da fila de UI a cada 50ms."""
        self._processar_fila()

    def _processar_fila(self) -> None:
        try:
            while True:
                mensagem = self._fila_ui.get_nowait()
                self._tratar_mensagem_ui(mensagem)
        except queue.Empty:
            pass
        self.after(50, self._processar_fila)

    def _executar_em_thread(self, funcao, *args) -> None:
        threading.Thread(target=funcao, args=args, daemon=True).start()
```

### Feedback Visual Obrigatório

| Estado | Cor | Comportamento |
|---|---|---|
| **Carregando** | Indicador `CTkProgressBar` indeterminado | Visível enquanto durar a operação |
| **Sucesso** | `ESMERALDA_SUCESSO` `#00C17C` | Exibir por ~2 segundos, depois limpar |
| **Erro** | `ERRO` `#C8102E` + mensagem descritiva | Permanecer até interação do usuário |
| **Aviso** | `AVISO` `#FFB800` + mensagem | Permanecer até interação do usuário |

### Rodapé Fixo — Obrigatório

```python
rodape = ctk.CTkLabel(
    self,
    text="Roberto Santos [LABS]©",
    font=ctk.CTkFont(size=10),
    text_color=TEXTO_SECUNDARIO,
)
rodape.grid(
    row=ultima_linha,
    column=0,
    columnspan=total_colunas,
    sticky="ew",
    padx=10,
    pady=(0, 8),
)
```

---

## 10. Estrutura de Projeto Padrão

Para todo projeto com mais de um arquivo, use esta estrutura:

```
nome_projeto/
│
├── main.py                  # Ponto de entrada — apenas inicialização
├── requirements.txt         # Dependências com versões fixadas
├── INSTALL.md               # Instruções de instalação e compilação
│
├── configuracao/
│   ├── __init__.py
│   ├── constantes.py        # Constantes globais e BASE_DIR
│   └── configuracoes.py     # Leitura de arquivos de configuração
│
├── servicos/
│   ├── __init__.py
│   └── *.py                 # Lógica de negócio por domínio
│
├── ui/
│   ├── __init__.py
│   ├── app.py               # Janela principal (CTk)
│   └── componentes/         # Frames e widgets reutilizáveis
│       ├── __init__.py
│       └── *.py
│
├── logs/                    # Gerado em runtime, nunca comitar
└── assets/                  # Imagens, ícones — referenciados via BASE_DIR
    └── *.png
```

---

## 11. Entregáveis Obrigatórios

Todo projeto deve incluir:

1. **Código 100% funcional** — executável imediatamente após `uv pip install -r requirements.txt`
2. **`INSTALL.md`** com:
   - Versão mínima do Python (ex: `>= 3.11`)
   - Comando de instalação via `uv`
   - Variáveis de ambiente necessárias (se houver)
   - Comando de compilação Nuitka completo (quando aplicável)
   - Estrutura de pastas esperada ao lado do executável
3. **Estrutura de pastas comentada** para projetos com múltiplos arquivos
4. **Revisão crítica** antes de entregar — verifique: acoplamento, cobertura de erros, consistência de nomenclatura e conformidade com estas diretrizes

---

## 12. Checklist de Qualidade — Obrigatório Antes de Entregar

### Código Geral
- [ ] Nenhum `# TODO`, placeholder, `...` vazio ou código comentado
- [ ] Tipagem estática completa em todas as funções e métodos públicos
- [ ] Docstrings em todas as classes e funções públicas (Google Style)
- [ ] Nenhum método supera 25 linhas de lógica
- [ ] Nenhuma lógica duplicada — DRY aplicado

### Caminhos e Portabilidade
- [ ] `BASE_DIR` definido e usado em **todos** os caminhos de arquivo
- [ ] Nenhum caminho absoluto ou relativo hardcoded fora de `BASE_DIR`
- [ ] Recursos externos (`.png`, `.ico`) estão em `assets/` e referenciados corretamente

### Robustez
- [ ] Todo I/O, chamada externa e integração tem `try/except` com log de erro
- [ ] Nenhum `except: pass` ou bloco de exceção silencioso
- [ ] `logging.exception()` usado em todos os blocos `except` de nível `ERROR`
- [ ] Exceções customizadas definidas para erros de domínio previsíveis

### Interface
- [ ] Thread principal nunca bloqueada — toda operação demorada em `threading.Thread`
- [ ] Feedback visual implementado para carregando, sucesso e erro
- [ ] Rodapé `"Roberto Santos [LABS]©"` presente e fixado com `sticky="ew"`
- [ ] Somente ícones sólidos — nenhum ícone de contorno
- [ ] `geometry("450x600")` e `resizable(False, False)` configurados
- [ ] Paleta de cores respeitada — nenhum valor de cor literal fora das constantes

### Entregáveis
- [ ] `INSTALL.md` completo, preciso e testado
- [ ] Nomenclatura 100% em português e consistente com as regras da seção 2
- [ ] Estrutura de pastas comentada (para projetos multi-arquivo)

---

*Roberto Santos [LABS]©*