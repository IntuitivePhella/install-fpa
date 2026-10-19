No MVP Excel UNTITLED, são extraídos \*\*quatro relatórios do Esferas\*\*. RRA e Contas a Pagar são consolidados de todas as produtoras; Contas a Receber e Comissões são filtrados para a UNTITLED na origem.

Aqui, \*\*sanitizar significa preparar os dados para carga no Excel, preservando valores, identificadores, fórmulas aplicáveis e rastreabilidade\*\*. As exclusões devem seguir critérios explícitos e deixar registro.

\*\*Período comum às quatro extrações\*\*

Todas usam o período \*\*acumulado de 1º de janeiro até o último dia do mês apurado\*\*, no campo de data correspondente ao relatório.

Por exemplo, para apurar setembro de 2026:

\- Início: \*\*01/01/2026\*\*.  
\- Fim: \*\*30/09/2026\*\*.  
\- Competência do fechamento: \*\*09/2026\*\*.

Extrair apenas setembro produziria um recorte mensal, insuficiente para a população acumulada prevista na DRE.

| Relatório | Campo de data | Recorte na origem | Formato bruto |  
|---|---|---|---|  
| Rentabilidade Analítica — RRA | Aprovação | Consolidado das produtoras | XLS recebido por e-mail |  
| Contas a Pagar — CP | Data competência | Consolidado das produtoras | CSV baixado |  
| Contas a Receber — CR | Data competência | Localidade da UNTITLED | CSV baixado |  
| Real x Orçado Comissões | Aprovação | Localidade da UNTITLED | XLSX baixado |

\*\*1. Rentabilidade Analítica — RRA\*\*

\*\*Extração\*\*

1\. Acessar o ambiente \[Papaki Esferas\](https\://papaki.esferas.com.br/index.php/home), com a sessão autenticada.  
2\. Entrar em \*\*Consolidado das produtoras → Jobs\*\*. A listagem de projetos abre em uma nova aba.  
3\. Selecionar \*\*Competência \= Aprovação\*\*.  
4\. Preencher o período acumulado.  
5\. Clicar em \*\*Pesquisar\*\*.  
6\. Selecionar \*\*Todos\*\* e conferir filtros, datas e resultados.  
7\. Abrir a \*\*seta à direita de Relatórios\*\*. A rotina registra que clicar diretamente no texto pode levar a uma página 404\.  
8\. Selecionar \*\*Rentabilidade analítica uma única vez\*\*.  
9\. Confirmar o aviso de envio por e-mail e registrar o horário real da solicitação.

O RRA \*\*não baixa diretamente nessa etapa\*\*: chega como anexo ao Gmail configurado na rotina.

Antes da solicitação, devem ser registrados os IDs das mensagens candidatas já existentes. Depois, identifica-se a mensagem nova vinculada ao disparo e salva-se o \*\*anexo original XLS\*\*, preservando seus bytes. A rotina prevê monitoramento por até 20 minutos, sem solicitar outro relatório por atraso.

O bruto permanece como XLS. Sua conversão acontece na sanitização.

\*\*Sanitização global\*\*

O tratamento deve:

\- Converter efetivamente \*\*XLS → XLSX\*\*; trocar apenas a extensão não converte o arquivo.  
\- Validar as \*\*49 colunas, A:AW\*\*, seus cabeçalhos e sua ordem.  
\- Preservar os registros das diferentes produtoras.  
\- Preservar valores e fórmulas.  
\- Remover apenas linhas técnicas reconhecidas por critérios determinísticos, como cabeçalhos auxiliares ou totalizadores identificados.  
\- Preservar o \*\*\`\#DIV/0\!\` esperado\*\* e bloquear erros inesperados.

Uma linha financeira não deve ser excluída apenas porque contém um erro esperado de cálculo percentual.

\*\*Preparação da BASE UNTITLED\*\*

Depois da sanitização global:

1\. Manter somente \*\*\`Núcleo \= Projetos \- Untitled\`\*\*.  
2\. Excluir linhas cujo \*\*\`Título do job\` contenha \`PESQUISA E TRATAMENTO\`\*\*.  
3\. Preservar as 49 colunas \*\*A:AW\*\*.  
4\. Reendereçar referências relativas das fórmulas quando a mudança de linha exigir.  
5\. Registrar a linha de origem, a linha de destino e o motivo de cada exclusão.

A resultante alimenta \*\*\`BASE\!A5:AW\`\*\* na cópia do modelo anual.

\*\*Ponto de atenção encontrado:\*\* o preparo legado em \`core.py\` calcula certas fórmulas e transforma alguns percentuais com denominador zero em \*\*“Não calculável”\*\*. Esse comportamento diverge do contrato vigente de preservação das fórmulas e do \`\#DIV/0\!\` esperado; portanto, não deve ser apresentado como a sanitização contratual concluída.

\*\*2. Contas a Pagar — CP\*\*

\*\*Extração\*\*

1\. No Papaki Esferas, acessar \*\*Consolidado das produtoras → Contas a pagar\*\*.  
2\. Na nova aba, preencher \*\*Data competência\*\* com o período acumulado.  
3\. Selecionar \*\*Status \= Todos\*\*.  
4\. Selecionar \*\*Pesquisar em \= Individuais e agrupadas\*\*.  
5\. Clicar em \*\*Pesquisar\*\* e conferir o resultado.  
6\. Inventariar os arquivos correspondentes já existentes em Downloads.  
7\. Abrir as opções pela \*\*seta à direita de Exportar\*\*.  
8\. Selecionar \*\*Excel uma única vez\*\*.

Apesar do nome “Excel” na opção de exportação, a fonte esperada pela rotina é um \*\*CSV\*\*, normalmente com nome iniciado por \`Contas a pagar\`.

O aviso de mais de 200 registros ou de demora significa que a geração já começou: confirma-se apenas o aviso e aguarda-se. Se o clique no corpo de \*\*Exportar\*\* já disparou a exportação, ele conta como o único disparo.

\*\*Sanitização global\*\*

Primeiro, validar o CSV, os cabeçalhos e as \*\*41 colunas, A:AO\*\*. Depois, aplicar esta regra:

\> \*\*Excluir a linha quando \`Nº job\` estiver preenchido e \`Job\` não contiver \`PESQUISA E TRATAMENTO\`.\*\*

| Nº job | Conteúdo de Job | Resultado |  
|---|---|---|  
| Vazio | Qualquer conteúdo | Manter |  
| Preenchido | Contém \`PESQUISA E TRATAMENTO\` | Manter |  
| Preenchido | Não contém \`PESQUISA E TRATAMENTO\` | Excluir |

Essa exclusão é aplicada à \*\*base consolidada inteira\*\*, sem filtro por Localidade nessa etapa.

O tratamento também deve:

\- Converter datas reconhecidas para datas reais do Excel.  
\- Converter valores monetários para números, respeitando separadores e sinais.  
\- Preservar identificadores textuais, inclusive zeros à esquerda.  
\- Distinguir vazio de zero.  
\- Remover somente rodapés ou totalizadores identificados.  
\- Rastrear cada exclusão.  
\- Gerar o consolidado sanitizado em \*\*XLSX\*\*.

Não se acrescentam filtros de investimento, fornecedores ou outras regras de runbooks antigos por inferência.

\*\*Preparação de DESPESAS UNTITLED\*\*

Após a sanitização global:

1\. Filtrar \*\*\`Localidade \= UTLD PRODUÇÃO DE FILMES LTDA\`\*\*.  
2\. Preparar as 41 colunas \*\*A:AO\*\*.  
3\. Carregar em \*\*\`DESPESAS\!A3:AO\`\*\*.  
4\. Preservar e estender as fórmulas existentes em \*\*AP:AQ\*\*, conforme o modelo anual apto.  
5\. Verificar cobertura e resultado dessas fórmulas em todas as linhas carregadas.

\*\*\`Criado em\` não integra a seleção do período.\*\* A classificação financeira em AP:AQ é uma etapa posterior à limpeza e à seleção dos registros.

\*\*3. Contas a Receber — CR\*\*

\*\*Extração\*\*

1\. No Papaki Esferas, acessar \*\*Consolidado das produtoras → Contas a receber\*\*.  
2\. Trocar o critério \*\*Venc. alterado\*\* por \*\*Data competência\*\*.  
3\. Preencher o período acumulado, sem deixar datas conflitantes em outros campos de competência.  
4\. Selecionar \*\*Status \= Todos\*\*.  
5\. Selecionar exatamente \*\*\`Localidade \= UTLD PRODUÇÃO DE FILMES LTDA\`\*\*.  
6\. Clicar em \*\*Pesquisar\*\* e conferir a seleção.  
7\. Inventariar os arquivos existentes em Downloads.  
8\. Acionar \*\*Exportar uma única vez\*\*, usando a seta se houver menu.  
9\. Confirmar eventual aviso e aguardar o CSV.

A opção \*\*\`UTLD PRODUCAO\` não deve ser selecionada\*\*.

Há uma exigência de proveniência importante: o CSV não comprova a Localidade em uma coluna. Por isso, \*\*a seleção visível no Esferas precisa ser registrada\*\*. Um consolidado sem essa prova não pode ser tratado como CR filtrado da UNTITLED.

\*\*Sanitização e preparação de RECEITA\*\*

Como o relatório já vem filtrado na origem, o tratamento deve:

1\. Validar os cabeçalhos e as \*\*28 colunas, A:AB\*\*.  
2\. Interpretar corretamente a codificação e o delimitador do CSV.  
3\. Tipar datas e valores monetários.  
4\. Preservar textos, identificadores e campos vazios.  
5\. Preservar \*\*todas as linhas financeiras\*\*, inclusive as que venham a receber a classificação \*\*\`Desconsiderar\`\*\*.  
6\. Registrar a correspondência entre registro de origem e linha de destino.  
7\. Gerar a base tratada e carregar em \*\*\`RECEITA\!A2:AB\`\*\*.

\*\*“Desconsiderar” é uma classificação, não uma ordem para apagar a linha.\*\*

A coluna \*\*AC\*\* deve ser completada pelo mecanismo demonstrado no modelo anual apto e pela definição financeira aplicável. O sanitizador não deve inventar classificação com base em históricos. Ausência ou ambiguidade na classificação impede a publicação conforme o contrato.

\*\*4. Real x Orçado Comissões\*\*

\*\*Extração\*\*

1\. Acessar o ambiente \[UNTITLED Esferas\](https\://untitled.esferas.com.br/index.php/home).  
2\. Navegar por \*\*JOBS → Projetos → Projetos\*\*.  
3\. Selecionar \*\*Competência \= Aprovação\*\*.  
4\. Preencher o período acumulado.  
5\. Selecionar exatamente \*\*\`Localidade \= UTLD PRODUÇÃO DE FILMES LTDA\`\*\*.  
6\. Clicar em \*\*Pesquisar\*\*.  
7\. Selecionar \*\*Todos\*\* e conferir filtros e resultados.  
8\. Inventariar os arquivos correspondentes em Downloads.  
9\. Em \*\*Relatórios\*\*, selecionar \*\*Real x Orçado Comissões uma única vez\*\*.

O relatório baixa diretamente em \*\*XLSX\*\*. Não há espera por e-mail.

Também aqui, a Localidade selecionada precisa ficar registrada: o arquivo isoladamente não comprova o filtro aplicado na tela.

\*\*Sanitização\*\*

Validar as \*\*14 colunas, A:N\*\*, na ordem esperada. Aplicar os critérios:

1\. \*\*Manter somente\*\* linhas com    
   \*\*\`Modelo de Orçamento \= Untitled \- Publicidade Padrão\`\*\*.  
2\. \*\*Excluir\*\* linhas cujo    
   \*\*\`Título do Job\` contenha \`PESQUISA E TRATAMENTO\`\*\*.  
3\. \*\*Excluir\*\* linhas cuja coluna \*\*H\*\* corresponda a uma destas seis opções:

   \- \`Seguro Produção\`  
   \- \`Taxa da produtora\`  
   \- \`Comissão total 1\`  
   \- \`Comissão total 2\`  
   \- \`Impostos\`  
   \- \`Taxa ADM\`

\*\*Todas as outras opções de H devem ser preservadas, inclusive opções novas.\*\* A regra é uma lista explícita de exclusões.

A saída deve:

\- Preservar as 14 colunas e seus dados.  
\- Trazer os cabeçalhos em \*\*A1:N1\*\*, com dados a partir da linha 2\.  
\- Preservar as fórmulas de origem, reendereçando referências relativas quando as linhas mudarem.  
\- Registrar contagens e motivos de exclusão.  
\- Permitir rastrear cada linha mantida até o XLSX bruto.

A carga corresponde a \*\*\`REAL X ORÇADO COMISSÕES\!A1:N\`\*\*, incluindo cabeçalhos; as fórmulas do modelo em \*\*O:Q\*\* são estendidas às linhas carregadas. A aba já deve existir vazia no modelo anual preparado pelo FP\&A.

