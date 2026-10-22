# papaki-dre

Pega as quatro bases do Esferas de uma produtora (RRA, CP, CR e Real x Orçado Comissões), limpa os
dados, guarda tudo no Supabase e gera uma **cópia** do Modelo DRE preenchida só nas abas de base.
MVP: produtora **Untitled**.

## Instalar na sua máquina (prompt padrão)

O **Sistema Central Papaki FP&A** é instalado por quem é do FP&A, na própria máquina, com o próprio acesso ao Esferas e ao
e-mail. Quem instala **não precisa saber programar**: abra o Claude (Claude Code) em qualquer pasta e cole:

> Clone o repositório `https://github.com/IntuitivePhella/install-fpa` numa pasta do meu computador e me ajude, passo a passo, em todo o processo
> de configuração e testes até o Sistema Central Papaki FP&A funcionar na minha máquina. Siga o arquivo `INSTALAR.md`.

O Claude conduz tudo (confere a máquina, cria o Python próprio do projeto, pergunta onde guardar modelos, extrações e
resultados, ajuda a configurar o banco, faz o ensaio e, só se você autorizar, cria a rotina agendada). Outra produtora:
`NOVA-PRODUTORA.md`. Regras fixas para o Claude: `CLAUDE.md`.

## Atenção: dados sensíveis

Os relatórios têm CPF, CNPJ e dados bancários. **Nunca** coloque no git as pastas `amostras/` e `out/`
nem o arquivo `.env` (o `.gitignore` já cobre os três, além de `*.xls`, `*.xlsx` e `*.csv`). Não cole
linhas desses relatórios em e-mails, chats ou logs.

## Requisitos

- Windows com Microsoft Excel instalado (converte o XLS do RRA e conta os `#N/D`).
- Python 3.11 ou mais novo (a instalação cria um ambiente próprio, a pasta `.venv`).
- Git e Chrome com a extensão do Claude, logado no Esferas e no Gmail do usuário.
- O projeto Supabase central da Papaki (um só para todos).

## Instalação

Use o prompt padrão do começo deste arquivo: o Claude segue o `INSTALAR.md`. Em resumo (para quem quiser fazer à mão):

    py -3.12 -m venv .venv
    .venv\Scripts\python.exe -m pip install -e ".[dev]"
    .venv\Scripts\python.exe -m papaki_dre configurar --dados-dir "<pasta de dados>" --criar-pastas ^
      --modelo untitled="<Modelo DRE base>.xlsx" --usuario-esferas "<e-mail>"

Depois abra o `.env` no Bloco de Notas e cole `SUPABASE_URL` e `SUPABASE_SERVICE_KEY`. Confira com
`.venv\Scripts\python.exe -m papaki_dre verificar`. **Use sempre o Python da `.venv`**: o `py` do PATH pode abrir um
Python sem as bibliotecas do projeto.

## Banco de dados (Supabase)

Isto é feito **uma vez só, por quem administra o banco da Papaki**; quem instala o sistema numa máquina só recebe a
URL e a chave.

1. Crie um projeto Supabase dedicado (não reaproveite o de outro sistema).
2. Aplique as migrações, **nesta ordem**: `supabase/migrations/0001_schema.sql` e `supabase/migrations/0002_reservas.sql`
   (painel, SQL Editor, ou CLI do Supabase). A 0002 serve a várias máquinas: guarda as reservas de rodada e o nome da
   máquina que importou. O desenho para as 18 produtoras está em `IMPLANTACAO-PAPAKI.md`.
3. A segurança é por RLS (regra de segurança por linha do Supabase) ligado em todas as tabelas e **sem
   nenhuma política**: só a chave de serviço (`service_role`, a chave de administrador do projeto) consegue
   ler e gravar. Por isso o `.env` leva só essa chave e ele **nunca vai
   para o git**. Não use a chave pública (`anon`) aqui.

### Foto vigente, reimportação e histórico

Cada importação é uma "foto" daquela produtora e competência. A foto mais recente concluída vira a
**vigente**. As anteriores continuam guardadas, com `vigente = false`: nada é apagado.

- Arquivos idênticos (mesmo conteúdo) aos de uma importação já concluída: o comando avisa e não faz
  nada. Use `--forcar` para repetir mesmo assim.
- Arquivos diferentes para a mesma competência (uma nova extração): a nova vira a vigente e a antiga
  fica no histórico.
- Rodada **bloqueada** ou que **falhou** fica registrada, mas não vira vigente.

### Banco pausado

No plano gratuito o projeto pausa depois de cerca de 7 dias sem uso. O comando testa o acesso antes de
gravar qualquer coisa. Se o banco estiver pausado, mostra "banco pausado: restaure o projeto e rode de
novo" e termina com código `3`, sem ter gravado nada. Restaure no painel do Supabase, espere ficar
ativo e rode de novo.

Com `--restaurar-banco` e `SUPABASE_ACCESS_TOKEN` no `.env` (chave pessoal do Supabase: Conta, Access
Tokens), o comando primeiro confirma, na API do Supabase, que o projeto está mesmo pausado; só então
pede a restauração e espera. O prazo total é de **9 minutos** (do início da restauração até desistir) e
isso só acontece uma vez por execução. Se o banco estiver inacessível por outro motivo, ele **não**
restaura nada. Sem a opção, continua só avisando.

## Uso

Os exemplos estão para o Prompt de Comando (cmd).

Com os quatro arquivos na mão:

    .venv\Scripts\python.exe -m papaki_dre importar --produtora untitled --competencia 2026-09 ^
      --cp "CP.csv" --cr "CR.csv" --rra "RRA.xls" --comissoes "Comissoes.xlsx" ^
      --modelo "Modelo DRE - Untitled - base.xlsx" --saida out\2026-09 ^
      --tela cp=100,cr=50,rra=200

No PowerShell, o mesmo comando usa o acento grave no fim da linha em vez de `^`:

    .venv\Scripts\python.exe -m papaki_dre importar --produtora untitled --competencia 2026-09 `
      --cp "CP.csv" --cr "CR.csv" --rra "RRA.xls" --comissoes "Comissoes.xlsx" `
      --modelo "Modelo DRE - Untitled - base.xlsx" --saida out\2026-09 `
      --tela cp=100,cr=50,rra=200

Ou, depois de uma extração fechada (veja "Rotina mensal"), com modelo e saída vindos do ambiente:

    .venv\Scripts\python.exe -m papaki_dre importar --produtora untitled --competencia 2026-09 --extracao <pasta da extração>

### Opções

| Opção | O que faz |
|---|---|
| `--produtora` | Produtora cadastrada (hoje: `untitled`). Obrigatória. |
| `--competencia AAAA-MM` | Mês da competência. Obrigatória. |
| `--cp`, `--cr`, `--rra`, `--comissoes` | Os quatro arquivos brutos. Obrigatórios, a não ser que use `--extracao`. |
| `--extracao <pasta>` | Pasta de uma extração **fechada**: no lugar dos quatro arquivos, de `--tela`, `--data-extracao` e `--origem` (passar os dois jeitos é erro). Lê os arquivos, os totais da tela, os filtros, o print e a hora real da extração. Grava também o `RESULTADO.md`. |
| `--restaurar-banco` | Se o banco estiver inacessível e a API do Supabase confirmar que o projeto está pausado, acorda o projeto (até 9 min, uma vez por execução). Exige `SUPABASE_ACCESS_TOKEN` e `SUPABASE_URL` no `.env`. |
| `--modelo` | Modelo anual vazio (`.xlsx`). Nunca é alterado: o sistema grava numa cópia. Opcional quando `PAPAKI_MODELO_<PRODUTORA>` está no `.env`. |
| `--saida` | Pasta de saída (criada se não existir). Opcional quando `PAPAKI_DADOS_DIR` está no `.env` (padrão: `<PAPAKI_DADOS_DIR>\apuracoes\<produtora>\<competência>`). |
| `--tela cp=N,cr=N,rra=N,comissoes=N` | Quantos registros a tela do Esferas mostrava. Diferença é só alerta. |
| `--periodo-fim AAAA-MM-DD` | Demonstrativo parcial. Padrão: último dia da competência (ou o da extração). |
| `--sem-banco` | Não usa o Supabase: só limpa, confere e gera o Excel. |
| `--sem-recalculo` | Não abre o Excel para contar `#N/D` (ver abaixo). |
| `--forcar` | Importa mesmo que os arquivos já tenham sido importados. |
| `--gatilho manual\|agendado` | Quem disparou a rodada. Padrão: `manual`. |
| `--origem` | Texto livre sobre a origem dos arquivos. Padrão: `arquivos locais`. |
| `--data-extracao` | Data/hora da extração (formato ISO-8601, ex.: `2026-10-08T16:34:00-03:00`). Sem fuso, vale horário de Brasília. Padrão: agora. |

### Exemplos

| Para quê | Comando (resumido) |
|---|---|
| Importação normal | `.venv\Scripts\python.exe -m papaki_dre importar --produtora untitled --competencia 2026-09 --cp ... --cr ... --rra ... --comissoes ... --modelo ... --saida out\2026-09` |
| Conferir os totais da tela | acrescente `--tela cp=100,cr=50,rra=200` |
| Testar sem tocar no banco | acrescente `--sem-banco` |
| Mais rápido, sem abrir o Excel para contar `#N/D` | acrescente `--sem-recalculo` |
| Repetir arquivos já importados | acrescente `--forcar` |
| Mês ainda em andamento | acrescente `--periodo-fim 2026-09-15` |

Os números acima são só de exemplo.

## Rotina mensal

A rotina roda todo mês, do 2º ao 5º dia útil (até a competência ficar importada), e também sob demanda. A tarefa agendada roda **todo dia útil de
manhã** e é o `agenda` que decide se hoje é dia de rodar (se não for, ela termina em silêncio). Quem dirige o Chrome é o **Claude**,
seguindo roteiros em texto (`runbooks/`); o programa só registra o que aconteceu e **barra** o que estiver
errado antes de exportar. Em resumo:

1. `agenda` diz se hoje é dia de rodar (2º ao 5º dia útil, e a competência anterior ainda sem importação
   concluída). Ele consulta o banco central: se **qualquer máquina** já importou a competência, não manda rodar.
2. O Claude abre a extração (`extracao iniciar`) e, para cada um dos quatro relatórios do Esferas,
   confere a tela, tira um print, registra (`extracao tela`), grava o disparo (`extracao disparar`),
   clica em Exportar e registra o arquivo baixado (`extracao arquivo`). O RRA vem por e-mail.
3. `extracao fechar` gera o registro final da extração; `importar --extracao` lê esse registro.
4. O resultado vira um `RESULTADO.md` curto e um aviso do Claude. **O sistema não envia e-mail.**

Roteiros (o texto que o Claude segue): `runbooks/extracao-esferas.md` e `runbooks/rotina-mensal.md`.

### Comandos novos

**`agenda --produtora untitled [--hoje AAAA-MM-DD]`** responde em uma linha de JSON (texto estruturado que outros programas leem) se deve rodar e qual é
a competência (o mês anterior). Dia útil é segunda a sexta sem feriado nacional (Carnaval e Corpus Christi
contam como dia útil). Fora do 2º ao 5º dia útil, ou com a competência já importada, responde "não rodar".
O `agenda` consulta o banco central da Papaki para saber se **alguma máquina** já importou a competência (e, se ela
já foi importada, não manda rodar). Opções: `--sem-banco` (decide só pela pasta local, sem consultar o banco) e
`--restaurar-banco` (se o banco estiver pausado, acorda o projeto antes de consultar; pode levar até cerca de 9 minutos).
Com `--reservar`, se for dia de rodar e o banco responder, ele reserva a produtora e a competência para esta máquina (nome
em `PAPAKI_MAQUINA`); se outra máquina já reservou, responde "não rodar". A reserva vale 6 horas e o `importar` a libera.

**`reserva obter --produtora <id> --competencia AAAA-MM`** faz a mesma reserva à mão (rodadas manuais) e diz qual máquina a
tem. **`produtora situacao [--competencia AAAA-MM]`** mostra, para cada produtora cadastrada, se a competência já foi
importada (até que dia e por qual máquina) e se há reserva ativa.

**`extracao <subcomando>`** (usado pelo Claude, raramente por você):

| Subcomando | O que faz |
|---|---|
| `iniciar` | Cria a pasta da extração. **Falha se há extração aberta** daquela produtora e competência: nesse caso use `--retomar`. |
| `iniciar --retomar` | Reabre a extração aberta e segue de onde parou. |
| `iniciar --nova` | Começa outra extração e **abandona em definitivo** a anterior (ela não aceita mais nada). É recusado se algum relatório dela já foi disparado e ficou sem arquivo. |
| `tela` | Registra o que a tela do Esferas mostra (Localidade, datas, status, total, print) e **valida na hora**. Se algo não bate, não se exporta. |
| `disparar` | Grava o horário do disparo **antes** do clique em Exportar. |
| `arquivo` | Registra o arquivo baixado: confere o formato e guarda uma cópia com nome fixo. |
| `rra-ids-antes`, `rra-escolher` | Só para o RRA: guarda os e-mails que já existiam e escolhe, sem ambiguidade, o e-mail novo do relatório. |
| `fechar` | Só passa com os 4 relatórios completos. Gera o registro que o `importar --extracao` lê. |
| `estado` | Mostra o que já foi feito (por relatório: tela, disparo, arquivo), para retomar. |

Regras que o programa impõe:

- **Um disparo por relatório.** Um relatório disparado e sem arquivo **nunca** é disparado de novo. Não
  existe comando para abandonar um disparo nem para redisparar: essa decisão é sua (do usuário).
- **A tela vale 15 minutos.** Se passar disso entre `extracao tela` e `extracao disparar`, o `disparar`
  recusa e a tela precisa ser conferida e registrada de novo.
- **Nomes fixos** dos arquivos na pasta da extração: `cp.csv`, `cr.csv`, `rra.xls`, `comissoes.xlsx`.
- **RRA por e-mail:** só vale o e-mail de `notas@papaki.com.br` cujo "solicitado em" cair entre 10 segundos
  antes e 180 segundos depois do disparo, e que não estivesse na lista de antes. Zero ou mais de uma
  candidata: o programa não escolhe.
- **Falha ou trava que recusa:** o Claude não exporta, não redispara, escreve o motivo num `PARADA.md` na
  pasta da apuração e avisa você. Quem decide o que fazer é você.

### Comandos de instalação

Nenhum deles mostra valores do `.env` (só nomes e respostas sim/não):

| Comando | O que faz |
|---|---|
| `configurar` | Grava no `.env` as respostas da instalação (pasta de dados, nomes das subpastas, Modelo DRE base, usuário do Esferas). Nunca grava chaves do Supabase: quem cola é o usuário. Com `--criar-pastas`, cria a pasta de dados e as duas subpastas. |
| `verificar` | Confere o ambiente item a item (OK, AVISO ou FALHA): Python, bibliotecas, `.env`, pastas, modelo, Excel e banco (`--sem-banco` pula o banco). |
| `produtora listar` \| `produtora mostrar --produtora <id>` \| `produtora situacao` | Lista as produtoras cadastradas, mostra as regras de uma (arquivo `produtoras/<id>.json`) ou mostra a situação de todas no banco central. |
| `instalacao mostrar` | Mostra o que o Claude precisa saber desta instalação: usuário do Esferas, nomes das subpastas e se há modelo configurado. |

### Variáveis do `.env` (só os nomes; os valores ficam só no seu `.env`)

| Nome | Para quê |
|---|---|
| `PAPAKI_DADOS_DIR` | Pasta onde ficam modelos, extrações e resultados. **Fica fora do projeto**: o programa recusa uma pasta de dados dentro do projeto. |
| `PAPAKI_MODELO_<PRODUTORA>` (ex.: `PAPAKI_MODELO_UNTITLED`) | Caminho exato do Modelo DRE base do ano daquela produtora. O sistema nunca o altera. O nome do arquivo pode mudar de um ano para o outro. |
| `SUPABASE_ACCESS_TOKEN` | Chave pessoal do Supabase (Conta, Access Tokens: uma senha de acesso à sua conta, gerada por você). Só serve para acordar o projeto pausado com `--restaurar-banco`. Nunca é impressa nem gravada. |
| `PAPAKI_PASTA_EXTRACOES` / `PAPAKI_PASTA_APURACOES` | Nomes das subpastas de extrações e de apurações dentro da pasta de dados (opcionais; padrão `extracoes` e `apuracoes`). Escolhidos na instalação. |
| `PAPAKI_USUARIO_ESFERAS` | E-mail do usuário do Esferas e do Gmail desta máquina. O Claude o usa para escolher o Chrome certo. |
| `PAPAKI_MAQUINA` | Nome curto desta máquina (ex.: `principal`, `grupo-a`). Vai para as reservas de rodada e para cada importação. Gravado por `configurar --maquina`. |
| `SUPABASE_URL`, `SUPABASE_SERVICE_KEY` | O banco central da Papaki. Quem cola é o usuário; nunca vão para o git. |

Sem `PAPAKI_DADOS_DIR`, os comandos que precisam de pasta param com a mensagem "defina PAPAKI_DADOS_DIR".

### Estrutura de pastas dentro de `PAPAKI_DADOS_DIR`

    <PAPAKI_DADOS_DIR>\extracoes\<AAAA-MM>\<produtora>\<AAAAMMDD-HHMM>\   4 brutos (nomes fixos), prints, estado.json, extracao.json
    <PAPAKI_DADOS_DIR>\apuracoes\<produtora>\<AAAA-MM>\                  Modelo DRE gerado, conferencia.md/json, RESULTADO.md, PARADA.md (se houve parada), work\
    <PAPAKI_DADOS_DIR>\ensaios\<data>\                                    saídas de ensaio com --sem-banco

Os nomes `extracoes` e `apuracoes` são os padrões; a instalação pode trocá-los por nomes que a Papaki já usa (`PAPAKI_PASTA_EXTRACOES`, `PAPAKI_PASTA_APURACOES`). `ensaios` é só a convenção
dos roteiros.

**A pasta de dados fica fora do projeto.** Prints e brutos têm nomes de clientes, CPF/CNPJ e dados
bancários: nunca vão para o git, para chats ou para logs.

### RESULTADO.md, PARADA.md e o que o `importar` devolve

- **`RESULTADO.md`**: resumo curto em português simples, escrito pelo `importar --extracao` na pasta da
  apuração. Nunca contradiz a conferência. Uma rodada nova apaga o resumo da rodada anterior.
- **`PARADA.md`**: escrito pelo Claude (nunca pelo programa) quando a rotina para antes do `importar`:
  etapa, motivo (sem dados) e como retomar.
- Arquivos idênticos aos já importados: o comando diz isso (e se é ou não a foto vigente) e não gera
  `RESULTADO.md` novo.

## O que a conferência diz

A conferência é feita por **contagens**: bruto, menos remoções técnicas (linhas vazias e rodapés de
TOTAL), menos exclusões de negócio (por motivo), igual às mantidas, que devem ser iguais às carregadas
na aba. Os valores das linhas de TOTAL dos relatórios não são lidos nem comparados.

### Status

- **ok**: tudo fechou e não há alertas.
- **com_alertas**: tudo fechou, mas há algo para olhar. O arquivo é publicado normalmente.
- **bloqueada**: a rodada **não** vira a vigente. Bloqueio pela conferência (contagens, fórmulas) gera
  o arquivo com o sufixo `_BLOQUEADO`; bloqueio por **erro de contrato** não gera arquivo Excel,
  só a conferência (`conferencia.md`/`json`) com o motivo.

O que **bloqueia**:

- as contagens não fecham (bruto diferente de mantidas + técnicas + exclusões);
- mantidas diferentes de carregadas na aba;
- linha carregada sem a fórmula do modelo (inclui o TOTAL da BASE);
- erro de contrato: cabeçalho ou colunas diferentes do esperado, registro com número de colunas
  inesperado, erro de célula inesperado, classificação ausente, modelo "sujo".

Outras falhas (arquivo aberto, erro do banco etc.) deixam a importação como **falhou**, com o motivo.

O que só **alerta** (não bloqueia):

- total visto na tela (`--tela`) diferente do arquivo;
- avisos dos sanitizadores;
- fórmulas do modelo sobrando abaixo dos dados carregados (mantidas);
- datas gravadas em célula sem formato de data;
- **`#N/D` em `DESPESAS!AP:AQ`**: classificação sem correspondência no De-Para. O sistema pede ao
  Excel um recálculo completo **numa cópia temporária** (o arquivo entregue não é recalculado nem
  alterado) e conta as células com `#N/D`. A classificação é regra da Papaki: o sistema só avisa. Se o
  Excel falhar, vira um alerta de que não foi possível contar. Com `--sem-recalculo` essa contagem
  não roda e a conferência registra "contagem de #N/D não executada".

**Campos informativos** (aparecem em "Informações" e nunca mudam o status): quantas constantes ou
fórmulas de outra família foram **substituídas** pela fórmula do modelo, quantas linhas tiveram a
**fórmula de dados da AP unificada** para a fórmula completa do modelo e se o **autofiltro foi
estendido** até a última linha carregada.

## Saídas

Na pasta de `--saida`:

- `Modelo DRE - <Produtora> - MM_AAAA.xlsx` (rodada ok ou com alertas) ou
  `Modelo DRE - <Produtora> - MM_AAAA_BLOQUEADO.xlsx` (rodada bloqueada pela conferência; erro de
  contrato não gera Excel);
- `conferencia.md` (para ler) e `conferencia.json` (para outros programas);
- `work\`: arquivos de trabalho (RRA convertido para XLSX).

Os arquivos são montados em um arquivo de preparo oculto e só trocados no fim, depois que o banco
aceitou a rodada. Assim, uma falha não deixa a pasta com um resultado "ok" que não está no banco. Se
o arquivo final estiver aberto no Excel, o comando pede para fechá-lo e rodar de novo.

**Atenção:** um `*_BLOQUEADO.xlsx` e um `.xlsx` "ok" de rodadas antigas podem existir juntos na mesma
pasta. Vale a **conferência mais recente** (`conferencia.md`/`json`), não o arquivo mais novo à vista.

Rodada com `--sem-banco` não é registrada no Supabase: a conferência diz isso (`"banco": false` e
"rodada sem banco: não registrada no Supabase"). Para só **regenerar o Excel** de uma extração já
importada, use `--sem-banco` numa **pasta de saída diferente**: o comando recusa `--sem-banco` numa
pasta que já tem a saída de uma importação registrada no banco, para nunca sobrescrevê-la. Também
recusa uma saída que cairia sobre o próprio arquivo de `--modelo`.

## Diagnóstico

Em caso de falha o comando mostra só `FALHA: <motivo>`, sem URL, chave ou conteúdo de linhas. Com
`set PAPAKI_DEBUG=1` (PowerShell: `$env:PAPAKI_DEBUG = "1"`) ele imprime também o rastro completo do erro, que pode trazer **detalhes do
banco** (inclusive trechos das linhas recusadas): use só com dados sintéticos e não cole a saída em
chats ou chamados.

## Códigos de saída

| Código | Significado |
|---|---|
| `0` | ok, com alertas, ou arquivos idênticos já importados (nada a fazer) |
| `2` | bloqueada (não publicada como vigente) |
| `3` | falha, banco pausado ou inacessível, ou argumento inválido |
| `4` | "aguarde": só no `extracao rra-escolher` (o e-mail do RRA ainda não chegou) |

O `importar` usa `0`, `2` e `3`. Os subcomandos `extracao` usam `0` (ok), `3` (erro ou trava recusada) e
`4` (aguarde, só no `rra-escolher`). O `agenda` usa `0` (e escreve a resposta em JSON) ou `3`.

## Testes

Rápidos (não precisam de Excel, dados reais nem banco; todos passam):

    .venv\Scripts\python.exe -m pytest tests -q -m "not real and not excel"

**Reais** (usam o Excel e/ou os brutos reais de `amostras\2026-09`; ignorados se os arquivos faltarem;
levam alguns minutos e não devem rodar durante uma extração). Os 20 de
`tests/test_cli_real.py` e `tests/test_build_real.py` passam:

    .venv\Scripts\python.exe -m pytest tests/test_cli_real.py tests/test_build_real.py
    .venv\Scripts\python.exe -m pytest tests -q -m "real"
    .venv\Scripts\python.exe -m pytest tests -q -m "excel"

Teste **ao vivo** e opcional, contra o Supabase real do `.env`. Usa só dados sintéticos e limpa o que
criou no fim. No Prompt de Comando:

    set PAPAKI_LIVE=1
    .venv\Scripts\python.exe -m pytest tests/test_db_live.py

No PowerShell:

    $env:PAPAKI_LIVE = "1"
    .venv\Scripts\python.exe -m pytest tests/test_db_live.py

## Regras

Ver `instrucoes_relatorios_ base.md` e `docs/superpowers/specs/`. O sistema só escreve nas abas
`BASE`, `DESPESAS`, `RECEITA` e `REAL X ORÇADO COMISSÕES`; as demais abas não são tocadas. As tabelas
dinâmicas do modelo não são atualizadas: isso é do FP&A.
