# Atualizador de Swaps — Orquestração de Marcação a Mercado (VBA)

Automação do fechamento diário de MTM/PU de uma carteira de contratos de swap, substituindo um processo manual, planilha por planilha, por uma
rotina orquestrada com fontes de dados integradas (Bloomberg, FRED API, provedor
de curvas via SFTP), calendário de dias úteis, log auditável para escalar o número de contratos sem alterar código.

> Projeto publicado como amostra de código.
> Nomes de cliente, contraparte, empregador, credenciais e o layout visual das
> planilhas foram removidos ou substituídos por valores fictícios. Nenhum dado
> real de mercado, cliente ou contraparte está presente neste repositório.

## Problema

O fechamento diário do MTM de cada contrato de swap era feito **manualmente,
calculadora por calculadora**: baixar as curvas do dia via SFTP, localizar os
arquivos de mercado da Bloomberg, abrir cada planilha, colar os dados na aba
certa, atualizar a taxa de CDI e conferir se o "Check" fechava em OK. Isso
tinha custos concretos:

- **Tempo**: minutos por calculadora, multiplicado por cada contrato ativo,
  todos os dias úteis.
- **Conhecimento tribal**: cada calculadora tinha sua própria célula de data
  e de validação (`N6`, `M6` ou `P6`, dependendo do contrato) — só quem já
  tinha mexido naquela planilha sabia onde colar o quê.
- **Escalabilidade zero**: adicionar um novo contrato exigia alterar código
  VBA hardcoded, em vez de apenas cadastrar o novo ativo.
- **Erro operacional silencioso**: sem log estruturado, uma curva não
  encontrada ou uma data errada só aparecia quando alguém abria a planilha e
  reparava manualmente.

## Solução

Uma rotina VBA orquestrada em módulos, com separação clara de
responsabilidades e arquitetura **data-driven** (a lista de contratos vive
numa aba de configuração, não no código):

- **Orquestrador central** (`FluxoCompleto`): dirige o fluxo do fim ao fim —
  resolve D-1/D-2 pelo calendário de dias úteis, localiza/copia a pasta MTM
  do dia, baixa e valida as curvas, dispara a atualização de cada contrato e
  grava um log físico (CSV) auditável, espelhado numa aba de Log.
- **Calendário de dias úteis** próprio, alimentado por uma aba de feriados.
- **Download automático de curvas** via SFTP (script gerado dinamicamente
  para o WinSCP) — credenciais **sempre lidas de parâmetro em runtime**,
  nunca hardcoded no código-fonte.
- **Integração com Bloomberg e FRED API**: localização automática dos
  arquivos de câmbio/curva do dia (estrutura de pastas ano/mês/dia) e
  consulta HTTP para a taxa SOFR diária, com parser de JSON tolerante a
  locale (evita o bug clássico do VBA de interpretar `.` como separador de
  milhar em `pt-BR`).
- **Atualização das calculadoras com auto-detecção**: a macro lê a fórmula
  da célula de validação de cada planilha para descobrir sozinha onde fica a
  célula de data, em vez de depender de um mapeamento hardcoded por
  contrato. Override manual disponível para os casos que fogem do padrão.
- **Cadastro assistido** (`IncluirCalculadora`): adicionar um novo contrato
  é preencher um assistente em tela, não editar código — a macro de
  atualização já processa qualquer linha nova da configuração.
- **Sincronização de código** (`SincronizarModulosVBA`): reimporta os módulos
  `.bas` do disco para dentro do workbook com um clique, eliminando o passo
  manual de Alt+F11 → remover módulo → importar arquivo a cada alteração.

## Resultado

- Fechamento de todos os contratos: de um processo manual, calculadora por
  calculadora, para **uma única macro** (`FluxoCompleto`).
- **Zero alteração de código** para escalar: novos contratos entram por um
  assistente de cadastro; a arquitetura já suportava a carteira crescer sem
  tocar em VBA.
- **Log auditável**: cada etapa (curva localizada, contrato atualizado,
  falha de parsing, curva ausente) fica registrada em CSV, permitindo
  reconstruir o que aconteceu em qualquer dia de fechamento.
- **Erros operacionais visíveis na hora**: farol de status por contrato
  (aba de configuração) mostra imediatamente quem fechou OK e quem precisa
  de atenção, em vez de descobrir o problema só quando alguém abre a
  planilha.
- **Credenciais fora do código-fonte**: SFTP e chave de API passaram a ser
  parâmetros de configuração, não constantes hardcoded (correção aplicada
  especificamente para esta publicação, documentada em `docs/roadmap.md`).
- Roadmap de evolução já definido para migrar a orquestração para um serviço
  Java/Spring Boot centralizado, com dashboard web (ver `docs/`).

## Estrutura do repositório

```text
vba/
  Mod0_SyncModulos.bas          # Sincroniza os .bas do disco com o workbook
  Mod1_Orquestrador.bas         # Fluxo principal, log, resolução de datas
  Mod2_DiasUteis.bas            # Calendário de dias úteis (feriados custom)
  Mod3_PastasMTM.bas            # Localização/cópia das pastas MTM D-1/D-2
  Mod4_INETX.bas                # Localização do arquivo de curvas separadas
  Mod5_Calculadoras.bas         # Loop de atualização + auto-detecção de célula
  Mod6_Bloomberg.bas            # Integração Bloomberg (câmbio/curva) + FRED API
  Mod7_IncluirCalculadora.bas   # Assistente de cadastro de novos contratos
  Mod8_DownloadINETX.bas        # Download das curvas via SFTP (WinSCP)
  Modulo_BaixaArquivosINETX.bas # Versão standalone anterior (referência histórica)
docs/
  roadmap.md                    # Estado do projeto e decisão de migração
  arquitetura-java-futura.md    # Estrutura Maven planejada para a v2 (Java)
```

## Tecnologias

VBA (Excel) · SFTP/WinSCP · Bloomberg (arquivos de mercado) · FRED API
(REST/JSON) · calendário de dias úteis customizado 

## O que foi intencionalmente omitido

- Nomes reais de cliente, contraparte, empregador e contratos (substituídos
  por nomes fictícios de forma consistente em todo o código).
- Credenciais de SFTP e chave de API (removidas; a rotina exige que sejam
  configuradas em runtime).
- O arquivo `.xlsm` em si e qualquer script que definisse o layout visual da
  calculadora (cores, posição de células, painel de KPIs) — o repositório
  mostra a lógica de automação, não a planilha proprietária da empresa.
