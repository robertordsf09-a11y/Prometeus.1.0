---
noteId: "186cabf050af11f192eea110685456dd"
tags: []
name: "Regras Críticas"
description: "Regras inegociáveis de desenvolvimento que devem permanecer no QWEN.md principal"
type: "feedback"

---

## Regras Críticas a Manter no QWEN.md Principal

1. **Perfil Esperado** - Engenheiro de Software Sênior com código pronto para produção
2. **Linguagem e Nomenclatura** - snake_case, PascalCase, Português do Brasil
3. **Tipagem Estática** - Obrigatório em todo código
4. **Arquitetura** - SRP, modularização estrita, DRY, injeção de dependências
5. **Dependências** - Use uv como gerenciador
6. **Portabilidade** - Compatibilidade com Nuitka e regra de caminhos
7. **Tratamento de Erros** - Todo I/O exige try/except, sem silenciar erros
8. **Logs** - Formato obrigatório e fábrica de logger
9. **Interface** - CustomTkinter, thread safety, feedback visual
10. **Estrutura de Projeto** - Organização padrão
11. **Entregáveis** - Código funcional, INSTALL.md, checklist
12. **Qualidade** - Revisão crítica antes de entregar

**Why:** Essas regras são o DNA do projeto e devem ser verificadas em cada interação
**How to apply:** Manha-las sempre visíveis no QWEN.md principal, não particionar