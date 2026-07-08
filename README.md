# Automação do fechamento diário de MTM de swaps

**Um projeto de controle e redução de risco operacional**, desenvolvido para
aumentar a precisão da coleta e autalização das informações. 

> Este é um estudo de caso publicado para fins de portfólio de trabalho.
> Por confidencialidade, nomes de cliente, contraparte e empregador não são
> citados, e nenhum código-fonte, dado real ou material proprietário está
> incluído aqui. O objetivo é apresentar o raciocínio e lógica utilizada
> por trás da solução.

## O risco que existia

O fechamento diário da marcação a mercado (MTM) de uma carteira de swaps
dependia de um processo manual, repetido planilha por planilha: baixar dados
de mercado do dia, localizar o arquivo certo, copiar os valores na aba
certa, atualizar a taxa e conferir se o resultado batia. Do ponto de vista
de controle, isso significava três riscos concretos:

- **Risco de pessoa-chave.** Cada planilha tinha suas particularidades, e
  só quem já tinha operado aquela planilha específica sabia exatamente onde
  cada dado deveria entrar. O processo dependia do conhecimento de uma
  pessoa, não de uma regra documentada e repetível.
- **Risco de erro silencioso.** Sem um registro estruturado de cada etapa,
  uma taxa desatualizada ou um dado colado no lugar errado só era percebido
  quando alguém abria a planilha depois — nunca no momento em que o erro
  acontecia.
- **Risco de escala.** Cada novo contrato adicionado à carteira significava
  mais um processo manual repetido todos os dias úteis, sem nenhum ganho de
  eficiência com o crescimento do volume.

Em uma instituição regulada, um processo de marcação a mercado que depende
de memória individual e não deixa rastro auditável é, por definição, um
ponto cego de controle.

## O que foi construído

Uma rotina automatizada que assume o fechamento diário do início ao fim,
substituindo a execução manual por um processo padronizado e verificável:

- **Padronização do processo**: a rotina segue sempre a mesma sequência de
  validações, independente de quem a executa ou de quantos contratos
  existam na carteira — elimina a dependência de conhecimento individual.
- **Coleta automática dos dados de mercado**: busca as fontes externas
  necessárias (curvas de juros, taxas de câmbio, taxas de referência) sem
  intervenção manual, reduzindo o risco de usar um dado errado ou
  desatualizado.
- **Registro auditável de cada execução**: cada etapa do fechamento fica
  registrada — o que foi buscado, o que foi atualizado, o que falhou — em
  um log que permite reconstruir exatamente o que aconteceu em qualquer dia,
  a qualquer momento.
- **Sinalização imediata de exceções**: em vez de descobrir um problema
  quando alguém abre a planilha por acaso, a rotina evidencia na hora quais
  contratos fecharam corretamente e quais precisam de atenção.
- **Escala sem esforço adicional**: novos contratos são incorporados por um
  cadastro simples, sem exigir alterações manuais recorrentes no processo a
  cada crescimento da carteira.
- **Credenciais de acesso fora do processo**: nenhuma senha ou chave de
  acesso a sistemas externos fica fixa dentro da rotina — são informadas em
  ambiente controlado, separado do código.

## Resultado

- O fechamento deixou de ser um conjunto de tarefas manuais repetidas por
  contrato e passou a ser **uma rotina única, padronizada e auditável**.
- **Rastro de auditoria completo**: qualquer divergência pode ser
  reconstruída depois, com evidência de qual dado foi usado e quando.
- **Erros passaram a ser visíveis no momento em que ocorrem**, não semanas
  depois, quando já é tarde para corrigir a causa.
- **Dependência de pessoa-chave eliminada**: o processo não depende mais de
  alguém lembrar onde cada dado vai em cada planilha.
- Um plano de evolução já foi desenhado para migrar essa automação para uma
  plataforma centralizada, com painel de acompanhamento web — passo natural
  para consolidar o controle em um único ponto de governança.

## Sobre esta publicação

Por se tratar de um processo real de uma instituição financeira, o
código-fonte, a planilha e qualquer material que identifique cliente,
contraparte ou empregador não são publicados aqui. Para uma conversa sobre
a lógica de risco e controle por trás da solução — ou para uma
demonstração — fico à disposição.
