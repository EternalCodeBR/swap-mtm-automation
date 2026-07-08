# Automação do fechamento diário de MTM de Swaps

**Um projeto de automação de fluxo e redução de risco operacional**, desenvolvido para
aumentar a precisão da coleta e atualização das informações. 

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
  cada dado deveria entrar. 
- **Risco de erro silencioso.** Sem um registro estruturado de cada etapa,
  uma taxa desatualizada ou um dado colado no lugar errado só era percebido
  quando alguém abria a planilha depois.
- **Risco de escala.** Cada novo contrato adicionado à carteira significava
  mais um processo manual repetido todos os dias úteis, sem nenhum ganho de
  eficiência com o crescimento do volume.

## O que foi construído

Uma rotina automatizada que assume o fechamento diário do início ao fim,
substituindo a execução manual por um processo padronizado e verificável:

- **Padronização do processo**: a rotina segue sempre a mesma sequência de
  validações, independente de quem a executa ou de quantos contratos
  existam na carteira.
- **Coleta automática dos dados de mercado**: busca as fontes externas
  necessárias (curvas de juros, taxas de câmbio, taxas de referência) sem
  intervenção manual, reduzindo o risco de usar um dado errado ou
  desatualizado.
- **Registro auditável de cada execução**: cada etapa do fechamento fica
  registrada: o que foi buscado, o que foi atualizado, o que falhou. 
  Todas as informações em um log que permite reconstruir exatamente o que aconteceu em qualquer dia, a qualquer momento.
- **Sinalização imediata de exceções**: em vez de descobrir um problema
  quando alguém abre a planilha por acaso, a rotina evidencia na hora quais
  contratos fecharam corretamente e quais precisam de atenção.
- **Escala sem esforço adicional**: novos contratos são incorporados por um
  cadastro simples, sem exigir alterações manuais recorrentes no processo a
  cada crescimento da carteira. O projeto é escalável. 
- **Credenciais de acesso fora do processo**: nenhuma senha ou chave de
  acesso a sistemas externos fica fixa dentro da rotina, são informadas em
  ambiente controlado, separado do código.

## Resultado

- O fechamento deixou de ser um conjunto de tarefas manuais repetidas por
  contrato e passou a ser **uma rotina única, padronizada e pronta para auditoria, se necessário**.
- **Erros passaram a ser visíveis no momento em que ocorrem**, não semanas
  depois, quando já é tarde para corrigir.
- **Dependência de pessoa-chave eliminada**: o processo não depende mais de
  alguém lembrar onde cada dado vai em cada planilha.
