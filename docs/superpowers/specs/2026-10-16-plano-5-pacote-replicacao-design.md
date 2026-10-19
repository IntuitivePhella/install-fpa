# Plano 5 — Pacote de replicação: "Sistema Central Papaki FP&A" (desenho)

Data: 2026-10-16 · Status: aprovado pelo usuário em 16/10/2026 (conversa de brainstorming) · Complementa a spec do MVP (§11) e a do Plano 4; não reabre as decisões 1–14 do HANDOFF.

## 1. Objetivo

Transformar o que foi construído e ensaiado para o Felipe num produto **agnóstico ao usuário**: qualquer membro do FP&A da Papaki, com licença do Claude, login no Esferas e o próprio e-mail, instala o sistema na própria máquina com **um prompt padrão**, e o Claude conduz toda a instalação, os testes e a criação da rotina agendada daquele usuário. Tudo sai de um **repositório público no GitHub do Felipe**.

O MVP continua sendo **só a Modelo DRE da Untitled**, mês a mês, a partir dos relatórios das duas URLs do Esferas (`papaki.esferas.com.br` e `untitled.esferas.com.br`). Outras produtoras (SOMA etc.) **não** são cadastradas neste plano.

## 2. Decisões do usuário (16/10/2026)

1. **Outras produtoras:** quando uma entrar, o Claude entrevista o usuário Papaki (extração, limpeza, pastas, Modelo base). O plano entrega a **mecânica** e o roteiro da entrevista; só a Untitled fica preenchida.
2. **Rotina agendada e e-mail do RRA são do usuário que instala:** a tarefa nasce no Claude dele e o RRA é lido no Gmail dele.
3. **Banco central:** um único projeto Supabase da Papaki (o "Central Papaki FP&A"). O administrador cria o projeto uma vez; cada máquina recebe as chaves no próprio `.env`, coladas pelo usuário (o Claude nunca as vê).
4. **Repositório público**, publicado com **histórico novo e limpo** (um commit inicial); o histórico completo continua só na pasta local. Sem licença acrescentada pelo Claude (decisão de vocês).
5. **Duas máquinas e a rotina:** só o `agenda` consulta o banco central (se a competência já foi importada por qualquer máquina, não roda). Risco residual conhecido: duas máquinas acordando no mesmo dia **antes** de qualquer importação ainda poderiam extrair ao mesmo tempo. Fica registrado no roteiro de instalação.

## 3. Desenho

### 3.1 Duas camadas de configuração

- **No repositório (igual para todos):** `produtoras/<id>.json`, um arquivo por produtora: `nome`, `localidade` (nome no Esferas; vale para a tela do CR e das Comissões e para o filtro do CP), `esferas` (endereço do consolidado e da produtora), `regras` (`nucleo_rra`, `modelo_comissoes`, `comissoes_excluir_h`) e `linha_mestre` (linha do modelo com a fórmula completa de `DESPESAS!AP`). Os valores da Untitled saem de `config.py` **sem mudar nenhum resultado** (os testes e os 20 testes reais provam isso).
- **Na máquina (`.env`, nunca vai para o GitHub):** `PAPAKI_DADOS_DIR`, `PAPAKI_PASTA_EXTRACOES` e `PAPAKI_PASTA_APURACOES` (opcionais; nomes das subpastas, padrão `extracoes` e `apuracoes`), `PAPAKI_MODELO_<ID>` (caminho do Modelo base de cada produtora), `PAPAKI_USUARIO_ESFERAS` (e-mail do usuário do Esferas/Gmail, para o Claude escolher o Chrome certo) e as chaves do Supabase.

Tolerâncias: hoje não existe tolerância por produtora (diferença tela × arquivo é só alerta; a janela do e-mail do RRA e a validade da tela de 15 min valem para todos). Por isso **não** entram no arquivo da produtora (YAGNI); entram quando existir uma regra que dependa delas.

### 3.2 Python fixo

A instalação cria `.venv` dentro da pasta do projeto. Roteiros e tarefa agendada passam a chamar `.venv/Scripts/python.exe -m papaki_dre ...` (caminho relativo à pasta do projeto, barras normais: funciona no PowerShell e no Git Bash). Nada de caminho escrito à mão nem de `py`. Um atalho `.cmd` foi descartado: o `cmd.exe` estraga argumentos com acento (`--localidade "UTLD PRODUÇÃO ..."`).

### 3.3 Comandos novos

| Comando | Para quê |
|---|---|
| `configurar` | Grava no `.env`, **sem mostrar o arquivo**, só as chaves permitidas (pasta de dados, nomes das subpastas, modelo por produtora, usuário do Esferas). Nunca grava segredos. `--criar-pastas` cria a pasta de dados e as duas subpastas. Cria o `.env` com as linhas dos segredos vazias se não existir. |
| `verificar` | Confere o ambiente e imprime `OK/AVISO/FALHA` por item, **sem valores**: Windows, Python e `.venv`, bibliotecas, Excel registrado, nomes do `.env`, pasta de dados (fora do projeto, gravável), modelos, e (sem `--sem-banco`) banco central e tabelas. |
| `produtora listar` / `produtora mostrar` | Imprime a configuração (sem segredos) para os roteiros usarem em vez de texto fixo. |
| `instalacao mostrar` | Imprime só o que o Claude precisa saber da instalação local: usuário do Esferas, nomes das subpastas, se há modelo (sim/não). Nunca caminhos nem chaves. |
| `agenda` (alterado) | Se o calendário manda rodar, consulta o banco central; com `--restaurar-banco` acorda o projeto pausado antes de consultar. Sem banco configurado ou inacessível, vale só a pasta local e o JSON diz isso. |

O `importar` também passa a garantir a linha da produtora na tabela `produtoras` do banco (hoje só a Untitled vem da migração), com os dados do arquivo da produtora.

### 3.4 Instalação conduzida pelo Claude

O repositório traz `CLAUDE.md` (regras fixas para qualquer sessão que abra a pasta), `INSTALAR.md` (roteiro) e `NOVA-PRODUTORA.md` (entrevista). O roteiro de instalação:

1. confere Windows, Excel, Git, Python 3.11+ (instalar só com o OK do usuário) e Chrome com a extensão do Claude, logado no Esferas e no Gmail do usuário;
2. cria `.venv` e instala o projeto;
3. **entrevista**: onde guardar modelos, extrações e resultados (reaproveitar uma estrutura que a Papaki já tem ou criar uma nova), qual é o Modelo DRE base da Untitled, qual é o e-mail do usuário do Esferas;
4. `configurar` grava as respostas; o usuário **cola as chaves do Supabase** direto no `.env`;
5. `verificar`; ensaio sem banco com a **checklist da primeira execução**; gravação no banco só com o "pode importar";
6. cria a tarefa agendada no Claude daquele usuário, só com o OK dele, a partir do modelo `runbooks/tarefa-agendada.md`.

O **prompt padrão** (no topo do `README.md`): "Clone o repositório `<URL>` numa pasta do meu computador e me ajude, passo a passo, em todo o processo de configuração e testes até o Sistema Central Papaki FP&A funcionar na minha máquina. Siga o arquivo `INSTALAR.md`."

### 3.5 Nova produtora

`NOVA-PRODUTORA.md`: o Claude **não adivinha**. Entrevista o usuário Papaki sobre extração (Localidade em cada relatório, endereços, telas), limpeza (Núcleo do RRA, modelo de comissões, exclusões), pastas, Modelo base (abas, linhas de cabeçalho e de dados, linha-mestre) e cria `produtoras/<id>.json`. Regra que cabe nos quatro campos atuais entra só por arquivo; regra de **tipo novo** (outro filtro, outra coluna) vira um plano próprio, com testes. Layout de modelo diferente do da Untitled também (hoje `config.DEST` é único).

### 3.6 Publicação

Repositório público: antes de subir, limpeza (e-mails pessoais, caminhos da máquina, código do projeto Supabase, HANDOFF fora do repositório) e uma **auditoria automática** (`scripts/auditar_publicacao.py`, também chamada por um teste) que falha ao achar e-mail fora da lista permitida, CPF/CNPJ fora de `tests/`, chave do Supabase, caminho `C:\Users\<nome>` ou arquivo `.xls/.xlsx/.csv/.env` rastreado. Publicação em pasta separada com um único commit; **só com o OK do usuário e a URL do repositório vazio**, e o push usa as credenciais do Git já logadas na máquina (o Claude nunca digita senha).

### 3.7 Prova

Instalação do zero numa pasta de teste desta máquina, seguindo só o `INSTALAR.md` e o prompt padrão. Limite honesto: quem executa já conhece o contexto; a prova definitiva é a primeira instalação real na Papaki. Depois, com o OK do usuário, a instalação atual do Felipe é atualizada (`.venv`, `.env`, texto da tarefa agendada).

## 4. Fora do escopo

Cadastrar outras produtoras; mudar regras da Untitled; criar banco novo; acrescentar licença; publicar sem aprovação; evitar por completo a corrida entre duas máquinas (só a consulta ao banco).

## 5. Riscos

- Instalador conhece o contexto: a prova do zero é parcial (ver 3.7).
- Rotina em duas máquinas (ver decisão 5).
- Repositório público expõe regras de negócio e nomes internos da Papaki (o usuário escolheu mesmo assim); não expõe dados dos relatórios nem chaves, e a auditoria impede o contrário.
- `pip install` e `winget` precisam de internet; falha de rede para a instalação e é relatada.

## 6. Testes

Dados sempre inventados. Os 908 testes rápidos continuam passando; os 20 reais (usam o Excel e as amostras) provam que a Untitled gera o mesmo resultado depois de a configuração sair do código. Novos: carregamento e validação do arquivo da produtora, `configurar` (nunca imprime valores nem grava segredos), `verificar` com dublês, `agenda` com banco falso, nomes de subpastas, regras passadas aos sanitizadores, texto dos roteiros e documentos (comandos citados existem), auditoria de publicação.
