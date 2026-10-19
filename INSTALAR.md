# Roteiro de instalação (o Claude executa, com o usuário ao lado)

Leia `CLAUDE.md` antes: as regras de lá valem em todos os passos. Pasta do projeto: a pasta onde está este arquivo.
Fale em português simples, **um passo de cada vez**, e confirme o resultado antes de seguir.
Os comandos com barras normais (`.venv/Scripts/python.exe`) valem no PowerShell e no Git Bash, não no cmd puro.

**Combine com o usuário no início** (diga em 3 linhas): a instalação leva de 30 a 60 minutos; é preciso internet, o
Microsoft Excel, o Chrome com o Esferas e o Gmail dele já logados e os dados do banco central da Papaki (o
administrador passa a URL e a chave do Supabase). Você vai pedir o "OK" dele antes de instalar programas, criar
pastas, capturar a tela (prints do Esferas no ensaio e nas extrações), acordar o banco pausado, criar a tarefa agendada
e gravar no banco.

## Passo 1: conferir a máquina

| O quê | Como conferir | Se faltar |
|---|---|---|
| Windows 10 ou 11 | já está rodando | o sistema não funciona em outro sistema |
| Git | `git --version` | peça o OK e instale com `winget install -e --id Git.Git` |
| Python 3.11 ou mais novo | `py -0p` lista os Pythons instalados | peça o OK e instale com `winget install -e --id Python.Python.3.12`, depois abra um terminal novo |
| Excel | o comando `verificar` (passo 6) confere | peça ao usuário para instalar/ativar o Excel |
| Chrome com a extensão do Claude | `list_connected_browsers` mostra o navegador | peça ao usuário para abrir o Chrome do perfil que tem o Esferas e o Gmail e instalar/ativar a extensão do Claude |
| Esferas e Gmail logados | abra `papaki.esferas.com.br`, `untitled.esferas.com.br` e o Gmail no Chrome certo e veja se há tela de login | **nunca digite senha**: peça ao usuário para entrar e avise quando terminar |

## Passo 2: Python próprio do projeto (`.venv`)

Na pasta do projeto, escolhendo a versão mais nova que o `py -0p` mostrou (3.11 ou mais; vale 3.14). O `3.12` abaixo é só
exemplo; use a mais nova que o `py -0p` listar:

```
py -3.12 -m venv .venv
.venv/Scripts/python.exe -m pip install --upgrade pip
.venv/Scripts/python.exe -m pip install -e ".[dev]"
.venv/Scripts/python.exe -m papaki_dre verificar --sem-banco
```

Se o `py` não existir ou o `py -0p` falhar (Python instalado sem o launcher), procure o `python.exe` instalado (`where python`
ou a pasta do Python) e use o caminho completo no lugar de `py -3.12`: `"<caminho do python.exe>" -m venv .venv`.

O último comando ainda vai acusar `FALHA` de configuração (pasta de dados, modelo): é esperado. Se der erro de rede no
`pip`, avise o usuário e tente de novo depois. **Dali em diante, todo comando usa `.venv/Scripts/python.exe`.**

## Passo 3: entrevista (uma pergunta por vez)

1. **Onde guardar modelos, extrações e resultados?** "A Papaki já tem uma pasta (no computador, no OneDrive ou em rede)
   onde vocês guardam o Modelo DRE e os relatórios do Esferas? Se sim, qual? Se não, eu crio uma nova; sugiro
   `C:\Users\<usuário>\papaki_dados`." Regras: tem de ficar **fora da pasta do projeto**, com espaço livre (as extrações
   somam dezenas de MB por mês) e, de preferência, fora de pasta sincronizada em tempo real (o OneDrive pode travar
   arquivo aberto). O sistema cria, dentro dela, uma subpasta de extrações (padrão `extracoes`) e outra de apurações
   (padrão `apuracoes`). Se a Papaki já usa outros nomes de pasta, use os nomes dela (`--pasta-extracoes` e
   `--pasta-apuracoes`).
2. **Modelo DRE base da Untitled:** "Qual é o arquivo do Modelo DRE base do ano (o vazio, sem os dados do mês)? Quem o
   mantém? Ele fica sempre no mesmo lugar?" Anote o caminho completo do `.xlsx`. O sistema nunca altera esse arquivo:
   trabalha numa cópia.
3. **Usuário do Esferas:** "Qual é o e-mail do usuário do Esferas e do Gmail desta máquina (é no Gmail dele que chega o
   relatório RRA)?"
4. **Rotina mensal automática:** "Esta máquina será a **responsável** por rodar a rotina automática da Untitled todo
   mês? Só **uma** máquina por produtora deve agendar." (Se for outra, pule o passo 9: nesta máquina só o disparo manual.)
5. **Banco central:** "O administrador do Supabase da Papaki já te passou a URL e a chave de serviço?" Se não, peça a ele.
   Se o banco ainda não existe: pare, explique que alguém da Papaki precisa criar o projeto (você só ajuda, com o OK dele:
   nome e organização, e a migração `supabase/migrations/0001_schema.sql`).

## Passo 4: gravar as respostas

```
.venv/Scripts/python.exe -m papaki_dre configurar --dados-dir "<pasta de dados>" --criar-pastas --modelo untitled="<caminho completo do .xlsx>" --usuario-esferas "<e-mail>"
```

Acrescente `--pasta-extracoes "<nome>"` e `--pasta-apuracoes "<nome>"` só se o usuário quis nomes diferentes dos
padrões. O comando mostra só os **nomes** do que gravou. **Nunca abra o `.env`** para conferir.

## Passo 5: as chaves do Supabase (o usuário cola)

Diga: "Vou abrir o arquivo `.env` no Bloco de Notas para você. Eu não vou ler nem mostrar nada dele." Abra sem bloquear a conversa: no PowerShell, `Start-Process notepad .env`; no Git Bash, `start notepad .env` (na pasta do
projeto). Se o `.env` já existia sem as linhas `SUPABASE_URL=` e `SUPABASE_SERVICE_KEY=`, o usuário deve acrescentá-las. Peça ao usuário para colar, **depois do sinal de `=`**, a URL em `SUPABASE_URL` e a chave de
serviço em `SUPABASE_SERVICE_KEY` (e, se tiver, a chave pessoal em `SUPABASE_ACCESS_TOKEN`, que só serve para acordar o
banco pausado), salvar (Ctrl+S) e fechar. Avise que essa chave dá acesso total ao banco: nunca vai para e-mail, chat
ou GitHub. Espere ele dizer "pronto".

## Passo 6: verificar

```
.venv/Scripts/python.exe -m papaki_dre verificar
```

Leia cada linha e explique, em português simples, o que for `FALHA` ou `AVISO`. Só siga quando não houver `FALHA`.
Banco pausado (plano gratuito pausa após cerca de 7 dias sem uso): peça ao usuário para clicar em "Restore project" no
painel do Supabase e rode `verificar` de novo.

## Passo 7: ensaio sem gravar no banco (a primeira execução)

Escolha a competência do mês anterior já fechado. Com o usuário **olhando o Chrome**, siga `runbooks/extracao-esferas.md`
até o `extracao fechar` e depois:

```
.venv/Scripts/python.exe -m papaki_dre importar --produtora untitled --competencia <AAAA-MM> --extracao "<pasta da extração>" --sem-banco --saida "<pasta_apuracao>/ensaios/<AAAA-MM-DD>"
```

(`<pasta_apuracao>` e `<pasta da extração>` vêm do `extracao iniciar`.) Nada é gravado no banco nesta rodada.

### Checklist da primeira execução

Faça junto com o usuário e relate cada item:

1. O comando terminou com `Status: ok` ou `Status: com_alertas` (nunca `bloqueada`).
2. Abra `conferencia.md` e leia as contagens (BASE, DESPESAS, RECEITA, Comissões): bruto = técnicas + exclusões +
   mantidas, e mantidas = carregadas. A diferença entre o total da tela e o do arquivo é **alerta esperado**, não erro.
3. Abra o `.xlsx` gerado no Excel: não pode aparecer aviso de reparo, e as quatro abas de base têm os dados.
4. O usuário do FP&A confere por amostragem algumas linhas do arquivo contra o Esferas.
5. As demais abas (DRE etc.) são do FP&A: o sistema não toca nelas. As tabelas dinâmicas não são atualizadas pelo sistema.
6. `#N/D` em `DESPESAS!AP:AQ` aparece como alerta quando a classificação não tem correspondência no De-Para: a
   classificação é regra da Papaki; o sistema só avisa.

Se algum item falhar, **pare**, explique ao usuário e não grave nada no banco.

## Passo 8: gravar no banco central

Só com o "**pode importar**" explícito do usuário:

```
.venv/Scripts/python.exe -m papaki_dre importar --produtora untitled --competencia <AAAA-MM> --extracao "<pasta da extração>" --gatilho manual --restaurar-banco
```

(`--restaurar-banco` acorda o banco se estiver pausado, e isso também está coberto pelo "pode importar". Pode levar até
cerca de 9 minutos: use timeout longo ou segundo plano e não cancele.) Repita o mesmo comando uma vez: deve responder "Arquivos idênticos já importados" e não gravar nada.
Diga ao usuário qual importação ficou como a vigente.

## Passo 9: tarefa agendada (só na máquina responsável e com o OK do usuário)

Crie a tarefa no **Claude desta máquina**, a partir de `runbooks/tarefa-agendada.md`, trocando `<PASTA_DO_PROJETO>`,
`<PRODUTORA>` e `<EMAIL_DO_USUARIO>`. Avise o usuário de que a tarefa só roda com o computador ligado, o aplicativo do
Claude aberto e o Chrome (com a extensão) aberto às 08:00. Avise também que a tarefa, por desenho, **acorda o banco pausado e grava a importação todo mês SEM pedir novo "sim"**: o
OK dele ao criar a tarefa vale para isso. **Peça esse OK de forma explícita antes de criar.** Mostre o nome, o horário e como desligar. **Risco conhecido:**
se **duas máquinas** agendarem a mesma produtora, o `agenda` só evita a repetição depois de uma importação gravada; duas
máquinas acordando no mesmo dia antes disso poderiam extrair ao mesmo tempo. Por isso, uma máquina por produtora.

## Passo 10: fechar

Resuma ao usuário em 3 linhas o que foi feito e lembre: para rodar um mês à mão, peça "rode o mês AAAA-MM" numa conversa
com o Claude nesta pasta; para atualizar o sistema, `git pull` e depois
`.venv/Scripts/python.exe -m pip install -e ".[dev]"` e `verificar`; resultados ficam na pasta de apurações. Outra
produtora: `NOVA-PRODUTORA.md`.

## Se algo der errado

| Situação | O que fazer |
|---|---|
| `pip` sem rede ou bloqueado | avise; tente de novo; pergunte se há proxy ou bloqueio da empresa |
| `verificar` acusa banco inacessível | banco pausado: restaurar no painel; senão conferir URL e chave (o usuário, no `.env`) |
| `verificar` acusa "não está em UTF-8" ao configurar | o usuário abre o `.env` no Bloco de Notas e usa Salvar como > UTF-8 |
| Excel ausente | sem Excel o sistema não converte o RRA nem confere; peça para instalar |
| Esferas ou Gmail pedem login | não digite senha; peça ao usuário para entrar |
| Qualquer coisa fora do roteiro | pare e pergunte; nunca improvise com dados reais |
