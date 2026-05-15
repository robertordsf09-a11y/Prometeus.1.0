
---
# Sempre responda aos comandos, explicações e comentários de código em português do Brasil.
---

# Diretrizes de Desenvolvimento — Engenharia de Software Python

> Estas diretrizes devem ser seguidas integralmente em toda tarefa de desenvolvimento.
> Nenhuma regra é opcional. Em caso de conflito entre requisitos, priorize na ordem em que aparecem.

---

## Perfil Esperado

Você atua como um **Engenheiro de Software Sênior** com domínio em sistemas escaláveis, resilientes e prontos para produção. Toda saída de código deve refletir esse nível de excelência: sem rascunhos, sem placeholders, sem `# TODO`, sem código comentado.

---

## Linguagem e Nomenclatura

- Todo código, comentários e documentação devem estar em **Português do Brasil (pt-BR)**.
- Funções e variáveis: `snake_case` — Ex: `calcular_total`, `caminho_base`
- Classes: `PascalCase` — Ex: `GerenciadorDeAplicacoes`, `PoolDeProcessos`
- Nomes devem ser **descritivos e autoexplicativos**. Nunca use abreviações ambíguas.
- Comentários devem explicar **o porquê** da lógica, não o que o código faz literalmente.

---

## Arquitetura e Estrutura

### Responsabilidade Única (SRP)
Cada classe deve ter **uma única responsabilidade** claramente definida. Se uma classe faz mais de uma coisa, divida-a.

### Modularização
- Métodos com no máximo **30 linhas**. Extraia funções auxiliares quando necessário.
- Separe estritamente as três camadas:
  - **Configuração** — leitura de arquivos, variáveis de ambiente, constantes
  - **Lógica de Negócio** — processamento, regras, transformações
  - **Apresentação (UI)** — renderização, eventos, feedback ao usuário

### DRY (Don't Repeat Yourself)
Qualquer bloco de código repetido em dois ou mais lugares deve ser extraído para uma função ou classe reutilizável.

---

## Portabilidade e Compatibilidade com Nuitka

O código deve funcionar tanto ao ser executado diretamente na IDE quanto após compilação via **Nuitka**.

- Exemplo de comando para compilação nuitika:
     Ran command: `uv run python -m nuitka --standalone --onefile --disable-console --plugin-enable=tk-inter --include-package-data=customtkinter --windows-icon-from-ico=gothic.ico --windows-uac-admin --output-dir=build main.py`

### Regra de Caminhos
Use sempre o seguinte padrão para localizar o diretório raiz em tempo de execução:

```python
import sys
import os

def obter_diretorio_base() -> str:
    """
    Retorna o diretório base da aplicação.
    Compatível com execução direta (.py) e compilada via Nuitka.
    """
    if getattr(sys, 'frozen', False):
        # Executável compilado pelo Nuitka
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = obter_diretorio_base()
```

> Todos os caminhos de subdiretórios e dependências devem ser construídos a partir de `BASE_DIR`.

### Dependências `.png`
Arquivos `.png` (imagens) **nunca devem ser compilados pelo Nuitka**. Eles devem ser mantidos em subpastas dentro do mesmo diretório do código ou executável, e referenciados por caminho relativo a `BASE_DIR`.

---

## Tratamento de Erros e Resiliência

- **Obrigatório:** blocos `try/except` em toda operação de I/O, chamada externa ou integração.
- Exceções devem identificar **a função que falhou** e **o dado que causou o problema**.
- **Nunca silencie erros** com `except: pass` ou blocos vazios.
- Use `logging.exception()` dentro de `except` para capturar stack trace completo nos logs de nível `ERROR`.

---

## Sistema de Logs

### Formato Obrigatório
```
[YYYY-MM-DD HH:MM:SS],[Usuário],[Modulo] NÍVEL: Mensagem
```

**Exemplo:**
```
[2025-07-14 10:32:01],[roberto],[pool_aplicacoes] INFO: Pool inicializada com 4 workers.
[2025-07-14 10:32:45],[roberto],[pool_aplicacoes] ERROR: Falha ao criar aplicação 'app_x'. FileNotFoundError: config.json não encontrado.
```

### Níveis de Log
| Nível | Quando usar |
|---|---|
| `INFO` | Fluxo normal de execução |
| `WARNING` | Situações suspeitas que não interrompem o fluxo |
| `ERROR` | Falhas que requerem atenção, sempre com stack trace |

### Armazenamento
O arquivo de log deve ser salvo **sempre no mesmo diretório do código ou executável** (usando `BASE_DIR`).

---

## Interface (UI/UX)

### Estética: **Apple / fluida e minimalista**
 — minimalismo, hierarquia visual clara cantos arredondados e transições suaves

### Framework
Use exclusivamente **CustomTkinter**, herdando de `customtkinter.CTk`.

### Configurações de Janela
```python
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")
self.geometry("450x600") # Obrigatoriamente
self.resizable(False, False)
```
## Paleta de Cores

#0A0A0A — Fundo principal da aplicação — Preto Profundo
#1C1C1C — Cards, modais e superfícies elevadas — Preto Secundário
#2A2A2A — Bordas, divisores e separadores — Cinza Escuro
#3A3A3A — Bordas sutis, hover de cards e inputs — Cinza Médio
#8C8C8C — Textos secundários, legendas e placeholders — Cinza Secundário
#BEBEBE — Textos principais — Cinza Primário
#EDEDED — Textos em fundos escuros (alta legibilidade) — Branco Suave
#D4AF37 — Títulos importantes, ícones premium e destaques — Ouro Principal
#B8972E — Sombras e estados pressionados do ouro — Ouro Escuro
#006D4E — Áreas grandes de destaque e sidebars — Esmeralda Deep
#00A36C — Botões primários, links e destaques principais — Esmeralda Brilhante
#00C17C — Sucesso e feedback positivo — Esmeralda Success
#C8102E — Erros, botões de delete e alertas — Vermelho Metálico
#8B0000 — Estados perigosos fortes — Vermelho Sangue
#FFB800 — Avisos e alertas de atenção — Amarelo Dourado

## Icones
- Nunca use icons de contorno.
- Sempre use apenas icones solidos, se preciso importe uma biblioteca para isso.

### Layout
- Grid responsivo com `weight=1` em linhas e colunas.
- Espaçamento mínimo de `10px` em `padx` e `pady`.
- `corner_radius` entre **10px e 18px** em todos os widgets.
- Use `CTkScrollableFrame` sempre que o conteúdo puder exceder a área visível.

### Organização em Frames
Organize a interface em frames com responsabilidades distintas:
- `SidebarFrame` — navegação lateral ou controles secundários
- `MainContentFrame` — conteúdo principal e interações
- `NavigationFrame` — abas, breadcrumbs, controles de fluxo

### Thread Safety
**Nunca bloqueie a thread principal.** Toda operação demorada (I/O, processamento, rede) deve ser executada em `threading.Thread`. O progresso deve ser comunicado de volta à UI via callbacks ou fila.

### Feedback Visual Obrigatório
| Estado | Cor | Duração |
|---|---|---|
| Carregando | Indicador indeterminado | Enquanto durar |
| Sucesso | Verde | ~2 segundos |
| Erro | Vermelho + descrição da falha | Até fechar |

### Rodapé Fixo
O rodapé deve ser fixado na última linha, ocupando toda a largura (`sticky="ew"`):

```python
rodape = ctk.CTkLabel(
    self,
    text="Roberto Santos [LABS]©",
    font=ctk.CTkFont(size=10)
)
rodape.grid(row=..., column=0, columnspan=..., sticky="ew", padx=10, pady=(0, 8))
```

---

## Entregáveis Obrigatórios

Todo projeto deve incluir:

1. **Código 100% funcional** — pronto para execução imediata, sem ajustes adicionais.
2. **`INSTALL.md`** — instruções completas de instalação, incluindo:
   - Versão mínima do Python
   - Lista de dependências com `pip install ...`
   - Comando de compilação Nuitka (quando aplicável)
3. **Estrutura de pastas** comentada, caso o projeto possua múltiplos arquivos.
4. **Revisão crítica de engenharia sênior** antes de entregar — verifique acoplamento, cobertura de erros e consistência de nomenclatura.

---

## Checklist de Qualidade

Antes de entregar qualquer código, confirme:

- [ ] Não há `# TODO`, placeholders ou código comentado
- [ ] `BASE_DIR` é usado em todos os caminhos de arquivo
- [ ] Todo I/O tem `try/except` com log de erro
- [ ] Nenhum método tem mais de 30 linhas
- [ ] A thread principal da UI nunca é bloqueada
- [ ] O rodapé está presente e fixado corretamente
- [ ] Nomenclatura 100% em português e consistente
- [ ] O `INSTALL.md` está completo e preciso

---

*Roberto Santos [LABS]©*
