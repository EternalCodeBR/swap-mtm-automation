# Status do Projeto e Roadmap de Transição

Este documento serve como a memória técnica do projeto "Atualizador de Swaps", consolidando o estado atual e a estratégia futura definida em maio de 2026.

## 1. Estado Atual (VBA)
*   **Correções Realizadas:**
    *   Sincronização de diretórios para a nova estrutura `Ano > Mês > Dia`.
    *   Refatoração para caminhos dinâmicos usando `%USERPROFILE%` (compatibilidade multi-usuário).
    *   Correção de erros de compilação no `Mod5_Calculadoras` (constantes e sintaxe).
    *   Implementação de placeholder `movarqu` para evitar interrupções no fluxo.
*   **Gargalo Identificado:** Atualização manual diária do SOFR6M e tempo de manutenção ao adicionar novas calculadoras.

## 2. Decisão Estratégica: Migração para Java
*   **Arquitetura:** Spring Boot (Back-end) + HTML/JS (Dashboard Web) + Apache POI (Manipulação Excel).
*   **Modelo de Implantação:** Centralizado em workstation local (Custo Zero), acessível via IP.
*   **ROTIME (Return on Time):**
    *   Investimento: ~25 horas de desenvolvimento.
    *   Economia: ~7 horas/mês.
    *   **Payback: 3,5 a 4 meses.**

## 3. Próximos Passos (Futuro)
1.  **Prova de Conceito (PoC):** Criar um serviço Java simples que abra uma calculadora via Apache POI e valide o status da célula de Check.
2.  **Módulo SFTP:** Implementar o download nativo do INETX em Java para eliminar a dependência do WinSCP.
3.  **Dashboard:** Desenvolver a interface web para gerenciamento dinâmico das 14 calculadoras.

## 4. Referências Técnicas
*   Modelo de Pastas: `arquitetura-java-futura.md` (detalha o `pom.xml` e a estrutura Maven).
