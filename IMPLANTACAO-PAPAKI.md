# Implantação na Papaki: o desenho e o roteiro de condução

**Leia este arquivo primeiro** (depois o `CLAUDE.md`). Ele diz como o sistema está desenhado para o grupo Papaki inteiro,
o que já foi decidido e em que ordem o Claude conduz a implantação. Para cada passo ele aponta o roteiro detalhado.
Fale com o usuário em português simples, uma pergunta por vez, e **não adivinhe regra de negócio**.

## 1. Objetivo final

Deixar o **Sistema Central Papaki FP&A** funcionando no ambiente da Papaki, conduzido pelo Claude junto com o usuário do
FP&A: as **18 produtoras** do grupo, cada uma com o Modelo DRE do ano preenchido todo mês a partir das bases do Esferas, em
uma **máquina principal** (que roda todas) e, se a Papaki quiser, em **máquinas de grupo** (que rodam só algumas).
Hoje só a produtora **Untitled** está cadastrada e ensaiada; as outras entram uma a uma (`NOVA-PRODUTORA.md`).

## 2. Como o sistema está desenhado

| Assunto | Decisão |
|---|---|
| Código | **Um só** para todas as produtoras. Nada de cópia do sistema por produtora. |
| Regras de cada produtora | Um arquivo `produtoras/<id>.json` no repositório (igual para todos): Localidade no Esferas, endereços, Núcleo do RRA, modelo de comissões, opções da coluna H que não entram, linha mestre da fórmula. |
| O que é da máquina | O `.env` (nunca vai para o GitHub): pasta de dados, caminho do Modelo DRE de cada produtora (`PAPAKI_MODELO_<ID>`), usuário do Esferas, **nome da máquina** (`PAPAKI_MAQUINA`) e as chaves do Supabase. |
| Modelo DRE | As 18 produtoras usam o **mesmo molde** (as abas `BASE`, `DESPESAS`, `RECEITA` e `REAL X ORÇADO COMISSÕES`). O Modelo base de cada uma nunca é alterado: o sistema grava numa cópia. Layout diferente do molde exige **plano próprio**. |
| Banco | **Um único** projeto Supabase da Papaki, um único esquema. Todas as tabelas têm `produtora_id`. A linha da produtora nasce sozinha na primeira importação. |
| Pasta de dados | **Compartilhada** entre as máquinas (disco de rede ou sincronização confiável). Cada máquina aponta o próprio `PAPAKI_DADOS_DIR` para a mesma pasta. Quem decide se a rotina roda é o **banco**, nunca só a pasta (a sincronização pode atrasar). |
| Rotina agendada | Uma tarefa **por produtora** em cada máquina que a roda (`runbooks/tarefa-agendada.md`), com **horários escalonados**. |
| Banco indisponível | A rodada agendada **segue e avisa** o usuário (decide só pela pasta local). |

## 3. Várias máquinas, sem rodar duas vezes

- **Máquina principal:** roda todas as produtoras. **Máquinas de grupo:** rodam só algumas, **além** da principal. Uma mesma
  produtora pode, portanto, ter tarefa em mais de uma máquina: isso é intencional. A regra antiga "uma máquina por
  produtora" **não vale mais**.
- **Horários:** a máquina de grupo agenda mais cedo (ex.: 07:00) e a principal mais tarde (ex.: 10:00). Se o grupo já
  importou, o `agenda` da principal vê a competência importada no banco e termina em silêncio; se o grupo falhou, a principal
  assume. Entre produtoras diferentes na mesma máquina, horários diferentes (comece com 30 minutos).
- **Reserva no banco (`0002_reservas.sql`):** a tarefa usa `agenda ... --reservar`. A máquina que chega primeiro reserva a
  produtora e a competência; a outra recebe `rodar: false` (`reserva: de_outra_maquina`). A reserva vale **6 horas**, é
  renovada pela mesma máquina e liberada pelo `importar` (importada, bloqueada, idêntica ou falha). Se o processo morrer, ela
  expira sozinha. Numa **parada** (`PARADA.md`) ela **não** é liberada de propósito: outra máquina não extrai do zero uma
  produtora cuja extração já foi disparada.
- **Sem nome da máquina ou sem banco:** a rodada segue e o Claude **avisa o usuário**; só o horário protege.
- **Rodada manual:** antes de extrair, `reserva obter --produtora <id> --competencia AAAA-MM`. Se outra máquina tem a
  reserva, avise o usuário (o JSON diz qual máquina e até quando); quem decide é ele.
- **Visão do grupo:** `produtora situacao` mostra, para cada produtora, se a competência está importada (até que dia e por
  qual máquina) e se há reserva ativa. Use no fim da implantação e todo mês.

## 4. Ordem de condução da implantação

Siga nesta ordem. Em cada fase, termine dizendo ao usuário, em 2 ou 3 linhas, o que passou, o que não passou e o que mudou.

1. **Pré-requisitos:** Windows, Excel, Git, Python 3.11+, Chrome com a extensão do Claude, logado no Esferas e no Gmail do
   usuário. Detalhes no `INSTALAR.md` (Passos 1 e 2).
2. **Banco central (uma vez só, por quem administra o Supabase da Papaki):** criar o projeto dedicado e aplicar **as duas
   migrações, nesta ordem**: `supabase/migrations/0001_schema.sql` e depois `supabase/migrations/0002_reservas.sql` (painel,
   SQL Editor, ou CLI). Quem instala as máquinas só recebe a URL e a chave de serviço, que o **usuário cola** no `.env`
   (o Claude nunca as vê). Criar projeto e aplicar migração só com o "sim" do usuário.
3. **Máquina principal** (`INSTALAR.md`, Passos 3 a 8): entrevista, `configurar` (inclui `--maquina principal`), `verificar`
   (deve mostrar `OK reservas`), ensaio **sem banco** da Untitled com o usuário olhando, e só então a gravação no banco com o
   "pode importar".
4. **Demais produtoras, uma por vez** (`NOVA-PRODUTORA.md`): entrevista, `produtoras/<id>.json`, `configurar --modelo`, ensaio
   sem banco, gravação com o "pode importar". Regra de **tipo novo** (outro filtro, outra coluna) ou layout de modelo diferente:
   **pare e abra plano próprio com testes**. Nunca invente a regra.
5. **Máquinas de grupo (se houver):** repetir `INSTALAR.md` com o nome da máquina (ex.: `grupo-a`) e só as produtoras dela
   (`PAPAKI_MODELO_<ID>` só dessas).
6. **Tarefas agendadas:** só com o OK explícito do usuário (a tarefa acorda o banco e grava toda vez sem novo "sim"). Uma por
   produtora em cada máquina, com horários escalonados (`runbooks/tarefa-agendada.md`). Meça o tempo real de uma extração e
   confirme a validade da reserva (6 h) e o intervalo entre tarefas.
7. **Aceite:** `verificar` sem FALHA em cada máquina; `produtora situacao` mostra a competência de cada produtora; uma rodada
   agendada real por produtora conferida pelo usuário; `scripts/auditar_publicacao.py` e os testes passando antes de
   qualquer commit.

## 5. Quando parar e perguntar

Dúvida de regra de negócio; regra de tipo novo; layout de modelo diferente; login pedido pelo Esferas, Gmail, GitHub ou
Supabase (nunca digite senha); extração já disparada sem arquivo (nunca redispare; `--nova` só se o usuário pedir);
conflito de reserva com outra máquina numa rodada manual; qualquer gravação no Supabase real, criação de projeto, instalação
de programa ou alteração de tarefa agendada sem o "sim" do usuário no chat.

## 6. Comandos de referência

Todos com `.venv/Scripts/python.exe -m papaki_dre <comando>`, na pasta do projeto.

| Comando | Para quê |
|---|---|
| `configurar` | Grava no `.env` as respostas (pasta de dados, modelo por produtora, usuário do Esferas, `--maquina`). Nunca grava segredos. |
| `verificar` | Confere o ambiente (OK, AVISO ou FALHA), sem mostrar valores. |
| `produtora listar` · `produtora mostrar` · `produtora situacao` | Produtoras cadastradas, regras de uma, situação de todas no banco. |
| `agenda --produtora <id> --restaurar-banco --reservar` | Diz se hoje é dia de rodar e reserva a rodada. |
| `reserva obter` | Reserva à mão (rodada manual) e diz quem tem a reserva. |
| `extracao ...` | Registro da extração do Esferas (usado pelo Claude). |
| `importar` | Sanitiza, grava no banco e gera a cópia do Modelo DRE. |

## 7. Pendências conhecidas

- Validade da reserva (6 h) e intervalo entre tarefas (30 min) **a confirmar** com o tempo real de uma extração.
- Sem banco ou sem nome da máquina, só o horário evita duas máquinas extraírem juntas.
- Desenho completo e decisões datadas: `docs/superpowers/specs/2026-10-22-plano-6-varias-maquinas-design.md`; plano de
  execução: `docs/superpowers/plans/2026-10-22-plano-6-varias-maquinas.md`.
