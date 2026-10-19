# Plano 4 — Automação mensal da extração (desenho)

Data: 2026-10-14 · Status: aprovado; alinhada ao implementado em 15/10/2026 (Tasks 1–11) · Complementa a spec do MVP (§5, §8, §10, §11) e não reabre as decisões 1–13 do HANDOFF.

## 1. Objetivo

Rodar todo mês, sem intervenção, o caminho Esferas (Chrome) → arquivos brutos → `importar`, e permitir o disparo manual a qualquer momento. O Claude dirige o Chrome seguindo roteiros; o código cuida do que precisa ser exato e testável. Não inclui o pacote de instalação na Papaki nem a configuração das 18 produtoras (Plano 5).

## 2. Decisões do usuário (14/10/2026)

1. **Banco pausado:** acordar automaticamente com a chave pessoal do Supabase (`SUPABASE_ACCESS_TOKEN`, só no `.env`). Sem a chave, vale o comportamento atual (aviso e código 3). A chave foi gerada pelo usuário, que a insere no `.env`; ela nunca é lida nem impressa.
2. **Agendamento:** tarefa agendada do aplicativo Claude (2º dia útil) **mais** disparo manual.
3. **E-mail do RRA:** chega no Gmail do Chrome do usuário do Esferas; a mensagem é achada e o anexo baixado pelo Chrome, preservando os bytes.
4. **O sistema não envia e-mail.** O desfecho é um resumo na pasta da competência e um aviso do Claude.
5. **Dia útil** = segunda a sexta sem os feriados nacionais (inclui Sexta-feira Santa e Finados). Carnaval e Corpus Christi contam como dia útil. Os feriados ficam em arquivo de configuração editável.
6. **Pastas fora do projeto e definidas pelo ambiente.** O projeto guarda só código, testes e roteiros. Onde ficam modelos, extrações e resultados é uma configuração (`PAPAKI_DADOS_DIR` no `.env`), escolhida na instalação orientada pelo Claude junto ao usuário Papaki: reaproveitando uma estrutura de pastas que já exista lá ou criando uma nova, conforme as respostas dele. Nada de dados dentro do repositório.
7. **Modelo DRE base por produtora:** um arquivo anual por produtora, mantido (substituído ou ajustado) pela equipe FP&A **no mesmo local**. O sistema guarda o caminho exato desse arquivo na configuração da produtora e nunca o altera. O nome pode mudar de ano para ano; o local não.

## 3. Fluxo mensal

1. A tarefa roda **todo dia útil de manhã** e pergunta ao programa (`agenda`) se hoje é o 2º dia útil **ou um dia útil seguinte** (até o 5º) em que a competência anterior ainda não tem conferência `ok`/`com_alertas` de uma importação real na pasta de apuração (é o que o `agenda` confere; o gatilho não entra). Se não, termina em silêncio. Assim uma falha é refeita no dia seguinte.
2. Garante o banco ativo (§4.2).
3. Extrai os 4 relatórios (§4.3), sempre registrando antes de clicar em Exportar.
4. RRA por e-mail (§4.4).
5. `importar --extracao <pasta> --gatilho agendado`.
6. Escreve `RESULTADO.md` na pasta da competência e avisa o usuário. Falhas e bloqueios também avisam.

## 4. Peças

### 4.1 `calendario.py`
`dias_uteis(ano, mes)`, `segundo_dia_util(ano, mes)`, `competencia_alvo(hoje)` (mês anterior) e `deve_rodar(hoje, ja_importada)`. Feriados: fixos mais Sexta-feira Santa (calculada a partir da Páscoa), mais os extras do arquivo de configuração. Comando `agenda --produtora P --hoje AAAA-MM-DD` imprime competência e se deve rodar: só do 2º ao 5º dia útil, e não se a pasta de apuração da competência já tem conferência `ok`/`com_alertas` de uma importação real com período até o fim do mês. Teste de referência: novembro de 2026 tem o dia 2 feriado (Finados) e o 2º dia útil é 04/11.

### 4.2 `supabase_admin.py`
Consulta o estado do projeto e pede restauração pela API de gerenciamento do Supabase, com `SUPABASE_ACCESS_TOKEN` e o ref do projeto (derivado de `SUPABASE_URL`, sem imprimir). Espera o projeto ficar ativo, consultando a cada 15 s, com **prazo global de 9 minutos** (do início da restauração até desistir; uma vez por execução). Integra-se ao `importar` só com a opção `--restaurar-banco` (a rotina mensal a passa; o uso manual comum não restaura sozinho): quando `verificar_acesso` falha e há chave, restaura, espera e testa de novo (o cliente do banco e a restauração leem o mesmo `.env` e conferem o mesmo projeto); sem chave ou com falha, código 3 como hoje. Todas as mensagens passam por `mensagem_segura` e nunca citam URL, ref nem chaves. Restaurar só ocorre com banco inacessível **e** estado "pausado" confirmado pela API (nunca em outro tipo de falha).

### 4.3 `extracao.py` (registro e travas)
Cada extração tem uma pasta (§4.7: `<DADOS>/extracoes/<competencia>/<produtora>/<AAAAMMDD-HHMM>/`) com `estado.json`. Comandos, chamados pelo Claude:
- `extracao iniciar --produtora --competencia --periodo-fim`: cria a pasta; **falha se há extração aberta** daquela produtora e competência (para continuar, use `--retomar`; a que tem disparo sem arquivo tem prioridade). `--nova` começa outra extração e **abandona em definitivo** a aberta (ela não aceita mais nada); é recusado se algum relatório dela já foi disparado sem arquivo. O Claude só usa `--nova` se o usuário pedir.
- `extracao tela --relatorio R --localidade ... --data-inicio ... --data-fim ... --status ... --total N --print arquivo`: grava o que foi visto na tela e **valida na hora**: Localidade (CP e RRA = "Todas"; CR e Comissões = `UTLD PRODUÇÃO DE FILMES LTDA`, nunca `UTLD PRODUCAO`), datas (01/01 do ano até o fim do período), Status = Todos nos relatórios que têm esse campo, print existente. Qualquer falha devolve erro e o Claude **não exporta**.
- `extracao disparar --relatorio R`: grava horário do disparo **antes** do clique; recusa se esse relatório já tem disparo (um disparo por relatório, também numa retomada). **A tela registrada vale 15 minutos**: passado esse tempo entre `tela` e `disparar`, o disparo é recusado e a tela precisa ser conferida e registrada de novo. No RRA, recusa também sem os ids de `rra-ids-antes`. **Não existe comando para abandonar nem redisparar um relatório**: essa decisão é do usuário.
- `extracao arquivo --relatorio R --caminho ...`: confere extensão e assinatura do formato (CSV, XLSX, XLS OLE2), calcula hash e tamanho, copia o arquivo para a pasta com **nome fixo** (`cp.csv`, `cr.csv`, `rra.xls`, `comissoes.xlsx`); um arquivo anterior ao disparo ou igual ao de outro relatório é recusado.
- `extracao fechar`: só passa se os 4 relatórios têm tela válida, disparo único e arquivo; gera `extracao.json` (manifesto) com filtros vistos, totais, prints, horários e hashes.
Retomada (`iniciar --retomar`): pula o que já tem arquivo; um relatório disparado e sem arquivo nunca é disparado de novo (o Claude procura o arquivo em Downloads, ou, no caso do RRA, no e-mail; se não achar, a rodada falha com motivo).

### 4.4 `rra_email.py`
Função pura `escolher_mensagem(candidatas, ids_antes, disparo, tolerancia_s)`. Aceita só: remetente `notas@papaki.com.br`, assunto "Relatório de rentabilidade gerado com sucesso", id que não estava em `ids_antes` e horário "solicitado em dd/mm/aaaa hh:mm:ss" (Brasília) entre 10 s antes e 180 s depois do disparo registrado (o disparo é gravado antes do clique). Zero ou mais de uma candidata → não escolhe (erro explícito). Comando `extracao rra-ids-antes` guarda os ids existentes (ids de **mensagem** do Gmail, no mesmo formato nos dois passos); no roteiro ele roda **antes** da tela do RRA, para tela e disparo ficarem próximos dentro dos 15 minutos. O `rra-escolher` sai com `0` (escolheu), `4` ("aguarde": ainda não chegou) ou `3` (erro, por exemplo mais de uma candidata). Espera de até 20 min, sem novo disparo por atraso. Após baixar, `extracao arquivo` confere que é XLS OLE2.

### 4.5 `importar --extracao <pasta>`
Lê `extracao.json`: caminhos dos 4 brutos, totais da tela (alimenta `--tela`), filtros vistos e caminho do print (preenche `arquivos_brutos.filtros_tela` e `prova_caminho`, que hoje ficam nulos), horário real da extração (da hora do primeiro disparo; ISO-8601 com fuso; sem fuso, assume America/Sao_Paulo) e origem. As opções atuais por arquivo continuam valendo e a nova é uma alternativa; passar as duas é erro de uso. Manifesto inválido (tela sem validação, hash que não confere com o arquivo) → código 3 sem gravar nada. Nas rodadas por `--extracao`, o `importar` escreve na pasta de apuração o **`RESULTADO.md`** (resumo curto em português simples, que nunca contradiz a conferência; uma rodada nova apaga o da rodada anterior; arquivos idênticos já importados não geram um novo). O **`PARADA.md`** é escrito pelo Claude, nunca pelo programa, quando a rotina para antes do `importar` (etapa, motivo sem dados, como retomar); o Claude nunca escreve em `RESULTADO.md`.
Limpeza de órfãs: no início de uma rodada com banco, importações `em_andamento` da mesma produtora e competência com mais de 6 horas são marcadas `falhou` ("rodada interrompida"). Os argumentos do `importar` e dos outros comandos não aceitam abreviações (`allow_abbrev=False`).

### 4.7 `caminhos.py` (onde ficam os arquivos)
Uma só peça sabe a estrutura de pastas; o resto do código pede o caminho a ela. A raiz vem de `PAPAKI_DADOS_DIR` (`.env`); sem ela, os comandos que precisam de pasta param com a mensagem "defina PAPAKI_DADOS_DIR". Fora da raiz, só o modelo base, cujo caminho vem da configuração da produtora: até o Plano 5, a variável `PAPAKI_MODELO_<PRODUTORA>` do `.env` (ex.: `PAPAKI_MODELO_UNTITLED`), decisão 7. Estrutura relativa à raiz (formato padrão; a instalação pode mapear para pastas já existentes da Papaki):

```
<DADOS>/extracoes/<AAAA-MM>/<produtora>/<AAAAMMDD-HHMM>/   4 brutos, prints, estado.json, extracao.json
                                                     (CR e Comissões são filtrados por produtora; CP e RRA
                                                     consolidados poderão ser reaproveitados no Plano 5)
<DADOS>/apuracoes/<produtora>/<AAAA-MM>/             Modelo DRE resultante, conferencia.md/json,
                                                     RESULTADO.md, work\ (RRA convertido)
<DADOS>/ensaios/<data>/                                  saídas de ensaio (criada ao usar --sem-banco --saida)
```

- Os nomes de subpastas (`extracoes`, `apuracoes`) são os padrões; ficam num único lugar da configuração para a instalação poder trocá-los por nomes já usados na Papaki.
- Várias rodadas da mesma competência continuam na mesma pasta de apuração; vale a conferência mais recente (regra atual do README). Um `RESULTADO.md` registra qual importação do banco é a vigente.
- `importar` passa a aceitar `--modelo` e `--saida` **opcionais** quando a produtora está configurada e `PAPAKI_DADOS_DIR` existe (modelo da configuração; saída em `apuracoes/<produtora>/<competencia>`). Passá-los explicitamente continua funcionando e vale sobre o padrão (testes e uso avulso).
- No desenvolvimento, `PAPAKI_DADOS_DIR` aponta para uma pasta **fora** do repositório, escolhida pelo usuário; as pastas `amostras/` e `out/` atuais seguem só como material de teste.
- O Plano 5 acrescenta a pergunta de instalação ("onde a Papaki quer guardar modelos, extrações e resultados? existe uma estrutura a reaproveitar?") e grava a resposta na configuração; o Plano 4 já deixa o código pronto para isso.

### 4.6 Roteiros (`runbooks/`)
`extracao-esferas.md` (passo a passo dos 4 relatórios, com as armadilhas: Localidade pré-preenchida, seta ao lado de "Relatórios" para não cair em 404, opção "Excel" que gera CSV, aviso de mais de 200 registros, CR em Windows-1252, `UTLD PRODUCAO` proibido) e `rotina-mensal.md` (fluxo da §3, é também o texto da tarefa agendada). Escritos para um Claude sem o contexto desta conversa. A regra de parada é explícita: qualquer trava que falhe → não exporta, registra o motivo em `PARADA.md`, avisa. Clique que não aconteceu depois de um `disparar` ok: não clicar de novo, não redisparar, parar e avisar. No RRA a ordem é: conferir a tela, `rra-ids-antes`, print, `tela`, `disparar` e clique imediato (janela de 180 s). Se o `py` do PATH não tiver as bibliotecas, os roteiros mandam usar o caminho completo do Python do projeto (a instalação do Plano 5 fixará isso).

## 5. Segurança e dados sensíveis

- O token pessoal vive só no `.env`, nunca em log, resposta ou teste. Em testes entra um valor falso.
- Prints de tela e arquivos brutos ficam sempre em `PAPAKI_DADOS_DIR`, fora do repositório (e o `importar` recusa uma raiz de dados que esteja dentro da pasta do projeto). Prints podem mostrar nomes de clientes e de usuário do Esferas: não vão para respostas nem chats.
- Mensagens de erro novas passam por `mensagem_segura`.
- Gravar no Supabase com dados reais continua exigindo o "pode importar" do usuário. Acordar o projeto também só com autorização para o primeiro teste real (leitura do estado).

## 6. Testes

Tudo com dados sintéticos: calendário (incluindo novembro de 2026, Páscoa de vários anos, viradas de ano), `supabase_admin` com API falsa (pausado → restaurando → ativo; tempo esgotado; sem chave; erro que não é pausa), `extracao` (cada trava, disparo duplicado, retomada, hash divergente), `rra_email` (zero, uma, duas candidatas, horário fora da tolerância, id antigo), `importar --extracao` (manifesto válido e inválido, fuso, órfãs). Os 401 testes rápidos e os 20 reais continuam passando.

## 7. Rodada assistida e tarefa agendada (não testáveis sem o Chrome)

1. **Rodada assistida** com o usuário olhando o Chrome. Primeiro com `--sem-banco`, comparando contagens com a extração de setembro. Gravar no banco só com o "pode importar"; como setembro já está no banco, uma extração nova da mesma competência viraria a foto vigente, então o usuário decide na hora se usa competência de ensaio ou fica sem banco.
2. **Teste curto da tarefa agendada** (abrir Esferas e Gmail, sem exportar) para saber se o Chrome logado e as permissões de site funcionam sem alguém na tela. A tarefa agendada só é criada depois disso e com autorização do usuário.
3. **Restauração real:** não se pausa o projeto para testar. A peça é validada com API falsa e, com autorização, por uma leitura do estado do projeto.

## 8. Riscos e pontos em aberto

- Tarefa agendada pode não conseguir usar o Chrome sem permissões já concedidas (teste da §7.2; plano B: disparo manual com o mesmo roteiro).
- O Esferas pode mudar a tela: os roteiros descrevem o objetivo de cada passo, não só cliques, e a rodada para quando algo não bate.
- A chave pessoal do Supabase enxerga todos os projetos da conta; por isso só é usada para ler estado e restaurar o projeto do ref configurado.
- Total de registros da tela é só alerta, como na spec (não há MCP do Esferas disponível como referência independente).
- Débitos do Plano 3 que ficam fora (o texto de erro dos sanitizadores que repetia célula já foi corrigido na auditoria): `exclusoes` repetindo o texto da regra, lixo acumulado de rodadas falhadas (500 MB), view da foto vigente.
