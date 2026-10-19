# MVP Papaki — Importação das bases do Esferas (Untitled)

Data: 2026-10-09 · Status: rascunho para revisão

## 1. Objetivo e escopo

**Objetivo do MVP:** validar que os dados do Esferas são **extraídos, sanitizados e alimentados** corretamente, no banco (Supabase) e na cópia do Modelo DRE. O princípio central: o sistema é orientado pelos dados e pelas regras de negócio, não pelas limitações físicas de uma planilha preexistente, e funciona **independentemente da data da extração** (fechamento mensal ou demonstrativo parcial).

**Dentro do escopo (MVP, produtora Untitled):**
- Extrair 4 relatórios do Esferas: RRA, CP, CR e Real x Orçado Comissões.
- Sanitizar conforme `instrucoes_relatorios_ base.md` (o `.md` é a regra vigente; o gabarito de 09/2026 foi montado com filtros mais amplos e **não** é fonte de verdade linha a linha).
- Gravar as bases tratadas no Supabase.
- Gerar uma **cópia** do Modelo DRE da produtora e carregar os dados **somente** nas abas de base.
- Emitir um relatório de conferência da importação.

**Fora do escopo (responsabilidade do FP&A):** conferir ou alterar fórmulas das demais abas, como a DRE é montada, inventário de fórmulas, views em SQL das análises e painel web. O sistema **não interfere** em nenhuma aba que não seja de base. O desenho não impede essa evolução futura.

**Escalabilidade prevista:** ~18 produtoras; o MVP só implementa a Untitled, mas todo dado carrega `produtora_id`.

## 2. Decisões já tomadas

| Tema | Decisão |
|---|---|
| Regra de verdade | O `.md`; satélites (SOMA, Vision, Plena, Serviços Compartilhados) ficam fora do MVP. |
| Supabase | Projeto único, `produtora_id` em todas as tabelas. |
| Execução | Claude Code local (computador ligado, Chrome logado no Esferas). Agendado no **2º dia útil** do mês seguinte + disparo manual a qualquer momento. |
| Reimportação | Cada importação é uma **foto completa** do período. A nova substitui a anterior na leitura; a anterior fica no histórico. Sem chave por linha no MVP; comparação entre extrações é evolução futura. |
| Fórmulas das bases | Preservadas como texto e **reendereçadas** quando a linha muda. `AC` (RECEITA) e `AP:AQ` (DESPESAS) são do modelo e apenas **estendidas** às novas linhas. |
| BUDGET | Entrada manual anual, já pré-preenchida no modelo; não é tocada. A aba oculta "Quadro de folha-desp" é ignorada. |
| E-mail do RRA | O sistema é agnóstico ao usuário: lê o Gmail conectado ao Claude de quem executa. |
| Divergências de contagem | Diferença entre tela e arquivo é **alerta registrado, nunca erro**. |

## 3. Arquitetura em cinco camadas

1. **Recepção:** caixa de entrada de arquivos. A automação do Chrome é *uma* fonte; subir arquivos manualmente usa o mesmo caminho.
2. **Validação, sanitização e padronização:** funções puras por relatório.
3. **Armazenamento:** Supabase (bruto referenciado, tratado, exclusões, execuções, cargas).
4. **Regras e cálculos:** fora do MVP (ficam no Excel do FP&A).
5. **Geração do demonstrativo:** no MVP, a cópia do Modelo DRE com as bases carregadas.

## 4. Modelo de dados (Supabase)

Todas as tabelas de dados levam `importacao_id`, `produtora_id` e `competencia` (ex.: `2026-09`); `data_extracao`, `origem`, período e gatilho ficam em `importacoes` (uma vez por foto).

- `produtoras` — cadastro e Localidade que a define no Esferas.
- `importacoes` — uma por rodada/foto: gatilho (manual/agendado), período (início e fim vindos do arquivo), status, horários, `vigente` (bool); campos de proveniência (`data_extracao`, `origem`, `motivo_falha`) e `conferencia` (`jsonb`, o relatório de conferência da rodada).
- `arquivos_brutos` — nome, hash, tamanho, relatório, filtros vistos na tela, prova (print) e total declarado na tela.
- `rra_base`, `cp_despesas`, `cr_receita`, `comissoes` — dado tratado: todas as colunas do relatório (49, 41, 28 e 14) em um `jsonb` com valores tipados (datas ISO-8601, números como número, vazio = `null`), mais chave promovida (`cod` no CP; `num_job` no RRA e nas Comissões) e rastreio (`linha_origem`, `linha_destino`, `importacao_id`). Colunas tipadas fixas podem ser geradas depois como views.
- `exclusoes` — relatório, linha de origem, `tipo` (`tecnica` = linha vazia ou rodapé de TOTAL; `negocio` = regra da §6), `motivo` (código) e `regra` (texto). As duas listas são gravadas, para refazer bruto → técnicas → exclusões → mantidas só com o banco.
- `cargas_excel` — aba, intervalo escrito e mapeamento linha de origem → linha de destino.

Segurança: RLS ativado; chave de serviço fora do código; projeto dedicado e restrito (há CPF/CNPJ e dados bancários de fornecedores no CP). Bruto **não** vai para controle de versão.

Reimportação: mesmo hash de arquivo = ignorado. Foto nova da mesma produtora e período → nova `importacao` marcada `vigente`, anterior desmarcada, nada apagado.

## 5. Extração (fonte: automação do Chrome)

Período acumulado: 01/01 até o fim do período apurado (parâmetro; padrão = último dia do mês de fechamento).

| Relatório | Caminho | Localidade | Formato bruto |
|---|---|---|---|
| CP | papaki.esferas → Consolidado → Contas a pagar; Data competência; Status=Todos; Pesquisar em=Individuais e agrupadas; Exportar → "Excel" (1 vez) | **Todas** | CSV UTF-8 com BOM, `;`, 41 colunas |
| CR | papaki.esferas → Consolidado → Contas a receber; Data competência (trocar "Venc. alterado"); Status=Todos; Exportar (1 vez) | **Somente `UTLD PRODUÇÃO DE FILMES LTDA`** (nunca `UTLD PRODUCAO`) | CSV **Windows-1252**, `;`, 28 colunas |
| Comissões | untitled.esferas → JOBS → Projetos; Aprovação; aba Todos; Relatórios → Real x Orçado Comissões (1 vez) | `UTLD PRODUÇÃO DE FILMES LTDA` | XLSX, 14 colunas (autoFilter inválido, ver §6) |
| RRA | papaki.esferas → Consolidado → Jobs; Aprovação; aba Todos; selecionar todos; Relatórios (seta) → Rentabilidade analítica (1 vez) | **Todas** | XLS OLE2 (`Plan1`), 49 colunas, chega por e-mail |

**Travas antes de exportar:** (1) conferir a Localidade na tela, porque o Esferas abre com a última usada; (2) conferir critério e datas após pesquisar (o formulário já recarregou sem as datas); (3) conferir o total de registros (referência independente: MCP do Esferas, tolerância calibrada por relatório); (4) um único disparo por relatório, registrado; (5) salvar print dos filtros como prova de proveniência.

**RRA por e-mail:** identificar a mensagem por remetente `notas@papaki.com.br`, assunto "Relatório de rentabilidade gerado com sucesso" e o horário "solicitado em dd/mm/aaaa hh:mm:ss" no corpo (tolerância de poucos segundos em relação ao disparo registrado). Registrar os IDs de mensagens candidatas antes do disparo. Espera de até 20 minutos, sem novo disparo por atraso. Anexo baixado pelo Chrome, preservando os bytes originais.

**Falhas:** execução marcada `falhou` com motivo; nada é carregado; retomada sem re-disparar o que já baixou.

## 6. Sanitização

Etapas comuns, nesta ordem: ler com a codificação correta por arquivo → validar cabeçalho e contagem de colunas (**diverge = para**) → remover só linhas técnicas por regra (vazia, `VALOR TOTAL`, `TOTAL PÁGINA`/`TOTAL GERAL`, `TOTAL`, títulos) → tipar (datas reais, valores numéricos, identificadores como texto com zeros à esquerda, vazio ≠ zero) → aplicar exclusões de negócio, registrando cada uma em `exclusoes`.

| Relatório | Regra de negócio |
|---|---|
| RRA | Manter `Núcleo = Projetos - Untitled`; excluir `Título do job` contendo `PESQUISA E TRATAMENTO`; preservar `#DIV/0!` esperado, bloquear erro inesperado; **não** calcular nem trocar por "Não calculável". |
| CP | Excluir se `Nº job` preenchido **e** `Job` sem `PESQUISA E TRATAMENTO`; depois filtrar `Localidade = UTLD PRODUÇÃO DE FILMES LTDA`. Sem outros filtros. |
| CR | Sem exclusão de linhas. "Desconsiderar" é classificação, não motivo para apagar. |
| Comissões | Manter `Modelo de Orçamento = Untitled - Publicidade Padrão`; excluir `PESQUISA E TRATAMENTO`; excluir da coluna H apenas: Seguro Produção, Taxa da produtora, Comissão total 1, Comissão total 2, Impostos, Taxa ADM (opções novas são preservadas). Cabeçalho na linha 3 do bruto (título na 1); fórmulas L:N preservadas e reendereçadas. |

Particularidades do bruto: o XLSX de Comissões tem `autoFilter` com referência inválida (leitores comuns falham); o RRA precisa de conversão XLS → XLSX real.

## 7. Carga no Excel (somente abas de base)

Origem: cópia do modelo anual da produtora preparado pelo FP&A.

| Aba | Intervalo | Observação |
|---|---|---|
| `BASE` | `A5:AW` | 49 colunas do RRA da Untitled. **Linha de total recriada** (ver abaixo). |
| `DESPESAS` | `A3:AO` | 41 colunas; estender fórmulas `AP:AQ` a todas as linhas. |
| `RECEITA` | `A2:AB` | 28 colunas; estender a coluna `AC`. |
| `REAL X ORÇADO COMISSÕES` | `A1:N` | Nome da aba é o do modelo (vale o do FP&A). Inclui cabeçalho; fórmulas de origem em L:N reendereçadas; **estender as fórmulas do modelo em `O:Q`** (Mês, Orçado único, Orçado por diretor) a todas as linhas carregadas. |

**Linha de total da `BASE`:** o modelo de origem não a traz; o sistema a recria. Posição: última linha de dados + 2 (uma linha em branco entre os dados e o total), como no modelo de setembro. Conteúdo: rótulo `TOTAL` na coluna A e `=SUM(X4:X<última linha de dados + 1>)` nas mesmas 35 colunas do modelo de setembro (corrigido em 09/10/2026: a lista sempre teve 35 colunas, conferido no gabarito de setembro): `J K L N P Q R S T U V W X Y Z AA AB AC AD AE AF AG AH AI AJ AK AL AM AN AO AP AQ AR AS AU`. As linhas 1 a 4 da aba (título, cabeçalho) pertencem ao modelo e não são alteradas; o intervalo somado começa na linha 4 (vazia), como no modelo de setembro.

**Tipo das células:** preservar o tipo original da célula de origem (um `Núm. JOB` numérico no XLS continua numérico; textos como `75131 #004` continuam texto). Não converter inteiros em texto como `75132.0`.

**Não interferir nas demais abas.** Método preferido: editar **somente o XML das abas de base** dentro do `.xlsx`, deixando o restante do arquivo intacto, com `fullCalcOnLoad` para o Excel recalcular ao abrir. Teste de aceite do método: após a carga, todas as partes do arquivo fora das 4 abas (e dos itens estritamente necessários, como strings compartilhadas) devem ser idênticas ao modelo de origem. Alternativa se o método falhar: automação do próprio Excel. O sistema **não** confere fórmulas de outras abas.

## 8. Relatório de conferência (por importação)

- ~~Soma das linhas lidas contra o total de controle do arquivo~~ **Removido em 09/10/2026:** as linhas de total (`VALOR TOTAL`, `TOTAL PÁGINA`/`TOTAL GERAL`, `TOTAL`) são apenas removidas como técnicas e não são usadas como conferência nem bloqueio.
- Contagens: bruto → após remoções técnicas → após exclusões (por motivo) → mantidas → carregadas na aba; bruto = mantidas + técnicas + exclusões e mantidas = carregadas, senão **bloqueia**. Erro de contrato (cabeçalho/colunas divergentes, registro com nº de colunas inesperado, erro inesperado de célula, classificação ausente) também deixa a importação **bloqueada**, registrada com o motivo; outras exceções a deixam **falhou**.
- Cobertura das fórmulas **apenas das abas de base** (`AP:AQ` em DESPESAS, `AC` em RECEITA, `L:N` e `O:Q` em comissões, total da `BASE`): as fórmulas valem para **todas** as linhas carregadas dos relatórios sanitizados; nenhuma linha carregada sem fórmula.
- Alertas de diferença tela × arquivo (nunca erro).
- Classificação ausente em `AC`/`AP:AQ` (modelo sem fórmula para estender, ou linha carregada sem a fórmula) → **bloqueia a publicação** (importação `bloqueada`, código de saída 2).
- Prova de Localidade (print) anexada.

## 9. Dados de teste e critérios de aceite

Duas extrações reais de 09/2026 do mesmo período: `amostras/2026-09/bruto` (minha, 08/10) e `bruto_papaki` (09/10). Contagens esperadas para a extração de 08/10: BASE 72 jobs, DESPESAS 1.110, RECEITA 484, Comissões 649. Para a de 09/10, as contagens serão calculadas na implementação. Diferenças entre as duas extrações **não** são defeito: servem para testar reimportação e foto vigente.

Aceite do MVP: (a) as 4 bases importam sem intervenção para Supabase e para a cópia do modelo; (b) o relatório de conferência fecha nas contagens (bruto → técnicas → exclusões → carregadas); (c) as abas fora das bases ficam inalteradas; (d) reimportar o mesmo arquivo não duplica; (e) importar a segunda extração torna-a vigente sem apagar a primeira.

## 10. Riscos e pontos em aberto

Decididos na revisão de 09/10/2026: a aba de comissões usa o nome e o layout do modelo (`REAL X ORÇADO COMISSÕES`); a linha de total da `BASE` é recriada pelo sistema; não há pendências de sessões anteriores.

1. **Anexo do RRA pelo Chrome** (download a partir do Gmail): ainda não testado, será validado no primeiro teste de ponta a ponta. Hoje o arquivo foi salvo manualmente.
2. **Escrita direta no XML** exige teste (strings compartilhadas, estilos, fórmulas compartilhadas). Alternativa: automação do Excel.
3. **Diferença de 36.997 (arquivo) × 37.030 (tela) no CP:** causa não confirmada; tratada como alerta.
4. **CR sem identificador por linha:** irrelevante para a foto vigente, relevante só para comparação futura entre extrações.
5. **Sessão do Esferas expirada no meio da rodada:** checagem no início; falha limpa.
6. **`.md` de instruções:** atualizado em 09/10/2026 para a nomenclatura vigente (`REAL X ORÇADO COMISSÕES`, com O:Q estendidas).
7. **Tabelas dinâmicas** do modelo mantêm o cache antigo até o FP&A atualizar.

### Decisões de 09/10/2026 (revisão do plano)

- Cabeçalhos são escritos pelo sistema a partir do bruto (o modelo-base os traz em branco).
- Fórmulas `O:Q` das comissões: usar as do modelo de setembro/2026.
- Dados tratados no banco: `jsonb` com valores tipados + chave promovida (ver Plano 3, Task 7).
- Totais de controle: **eliminados** (decisão de 09/10/2026, revista após a revisão do Plano 1). Linhas de total são só removidas como técnicas, reconhecidas por rótulo e posição exatos. Registro com nº de colunas inesperado (que não seja rodapé reconhecido) para a rodada.
- **Supabase no plano gratuito** (500 MB; pausa automática após ~7 dias sem atividade). Migrar de plano é decisão do ambiente Papaki.
- **Supabase pausado** (confirmado em 09/10/2026): no MVP o `importar` apenas **detecta** o banco inacessível/pausado antes de gravar qualquer coisa, avisa "banco pausado: restaure o projeto e rode de novo" e sai com código 3. A restauração automática do projeto antes de importar foi construída no **Plano 4** (opção `--restaurar-banco`, com a chave pessoal `SUPABASE_ACCESS_TOKEN`; só restaura com banco inacessível e pausa confirmada pela API).
- **Rodadas bloqueadas** (confirmado em 09/10/2026): a importação é criada no banco antes de sanitizar; um erro de contrato (sanitização ou montagem do Excel) a deixa `bloqueada` com o motivo, sem gravar linhas nem `arquivos_brutos`/proveniência (código 2). Uma rodada bloqueada pela conferência (contagens que não fecham, mantidas ≠ carregadas, linha sem fórmula do modelo) grava linhas e exclusões no banco **sem** torná-las vigentes (código 2).
- **Classificação** (confirmado em 09/10/2026): as fórmulas de classificação/cálculo do modelo (`AC`, `AP:AQ` e as demais da §8) precisam abranger todos os registros carregados; só a classificação **ausente** é detectada e bloqueia. A noção de classificação "ambígua", que constava da §8 até esta data, ficou **fora do MVP** (registro histórico).

### Decisões de 13/10/2026 (Plano 3 construído)

- **Alerta de #N/D:** o `importar` recalcula uma **cópia temporária** do arquivo no Excel (instância própria, só leitura, cópia apagada no fim) e conta os `#N/D` em `DESPESAS!AP:AQ`. Só gera **alerta** (classificação sem correspondência no De-Para; regra da Papaki), nunca bloqueia. Se o recálculo falhar, vira aviso na conferência. `--sem-recalculo` pula a contagem e o relatório registra "contagem de #N/D não executada". O arquivo entregue nunca é recalculado.
- **Relatório de conferência:** além das contagens, mostra por aba `constantes_substituidas` (`overwritten_constants`), `formulas_unificadas` (`unified_formula_rows`, linhas cuja fórmula de dados de `DESPESAS!AP` foi unificada para a **fórmula completa** do modelo, decisão 12 do HANDOFF) e `autofiltro_estendido` (`autofilter_extended`). São **informativos**: aparecem em "Informações" e não alertam nem bloqueiam.
- **Arquivo de saída e conferência:** o Excel é montado num arquivo de preparo oculto na pasta de saída e só é trocado pelo nome final ao fim, depois de gravar o banco. Se a rodada falhar, o arquivo e a conferência anteriores ficam intactos, ou a conferência é regravada como `falhou`: a pasta nunca mostra conferência "ok" de uma rodada que não está no banco. Rodada bloqueada gera o arquivo com sufixo `_BLOQUEADO`.
- **Segurança do banco:** RLS ligado em todas as tabelas e **sem políticas** (só a chave de serviço acessa); privilégios de `anon` e `authenticated` revogados nas tabelas e sequences; a promoção a foto vigente é uma **função SQL atômica** (`promover_importacao`) com trava por produtora e competência, só executável pela `service_role`. Um índice único garante no máximo uma foto vigente por produtora e competência.
- **Banco pausado:** detectado antes de gravar qualquer coisa; avisa "banco pausado: restaure o projeto e rode de novo" e sai com **código 3** (a restauração automática foi construída no Plano 4, com `--restaurar-banco`).

## 11. Replicação no ambiente Papaki (requisito)

Depois de validado aqui, o sistema precisa poder ser **instalado e reexecutado no ambiente da Papaki**, por qualquer usuário do FP&A que tenha licença do Claude e login no Esferas (o sistema é agnóstico ao e-mail do usuário). Entregáveis previstos (**Plano 5**, após os Planos 1–4):

1. **Pacote de instalação:** repositório com `pyproject.toml`, `.env.example`, README de instalação e verificação do ambiente (Windows, Excel, Python, Chrome, extensão do Claude, acesso ao Supabase).
2. **Configuração por produtora** (`PRODUTORAS` em `config.py` hoje; arquivo de configuração depois): nome, Localidade no Esferas, modelo anual, regras específicas, para escalar das 18 produtoras sem mexer em código.
3. **Prompts e runbooks** da extração pelo Chrome (os 4 relatórios, as travas de Localidade/datas/total, o e-mail do RRA), escritos para o Claude Code executar sem depender do contexto desta conversa.
4. **Rotina agendada** (2º dia útil do mês seguinte) e comando de disparo manual sob demanda.
5. **Checklist de validação** na primeira execução no ambiente Papaki (contagens, conferência manual do arquivo no Excel).
6. Projeto Supabase da Papaki (dedicado, restrito) e aplicação da migração.
