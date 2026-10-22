# Plano 6 — Várias produtoras e várias máquinas (desenho)

Data: 2026-10-22 · Status: **aprovado pelo usuário em 22/10/2026** (decisões da seção 8 respondidas) · Complementa as specs do MVP e dos Planos 4 e 5; não reabre as decisões 1–14 do HANDOFF.

## 1. Objetivo

O grupo Papaki tem 18 produtoras, todas com o **mesmo molde** de Modelo DRE. Este plano deixa o sistema pronto para rodar todas elas, em mais de uma máquina, sem duplicar código, banco ou rotina, e sem duas máquinas extraírem a mesma produtora e competência ao mesmo tempo.

## 2. Decisões do usuário (22/10/2026)

1. As 18 produtoras usam o mesmo molde de Modelo DRE (o layout único de `config.DEST` e `headers.json` serve).
2. Há **uma máquina principal que roda todas as produtoras**. Pode haver **máquinas de grupo**, que rodam só algumas, **além** da principal. A sobreposição é intencional.
3. As tarefas agendadas **não precisam ser simultâneas**: podem ter defasagem de horário entre si.

Consequência: a regra "uma máquina por produtora" (INSTALAR.md, pergunta 4 e Passo 9; `runbooks/tarefa-agendada.md`) deixa de valer e precisa ser reescrita.

## 3. O que o código faz hoje (lido em 22/10/2026)

- O banco só vê a rodada quando o `importar` chama `criar_importacao` (`cli.py`, linha ~579), **depois** de toda a extração. A janela de corrida entre duas máquinas é, portanto, a extração inteira.
- O `agenda` só pergunta ao banco se existe foto vigente completa (`vigente_completa`). Não existe "alguém está extraindo agora".
- `origem` em `importacoes` é a constante `"Esferas (Chrome)"`: não diz qual máquina importou.
- Não existe identificação da máquina em lugar nenhum.
- `encerrar_orfas` (6 h) só limpa importações `em_andamento` criadas pelo `importar`; não cobre a fase de extração.
- A saída (cópia do DRE, `RESULTADO.md`) fica na pasta de dados **de quem importou**.

## 4. Desenho

### 4.1 Parte A — só texto, sem código (entrega imediata)

- **Escalonar horários.** Máquinas de grupo agendam mais cedo e a principal mais tarde no mesmo dia: se o grupo já importou, o `agenda` da principal vê a foto vigente e termina em silêncio; se o grupo falhou, a principal assume sozinha (rede de segurança). Dentro de uma mesma máquina, cada produtora tem a própria tarefa, com horários diferentes (intervalo a definir pelo tempo real de uma extração).
- Reescrever a regra de máquinas e a orientação de horários em `INSTALAR.md`, `runbooks/tarefa-agendada.md` e `NOVA-PRODUTORA.md` (Passo 6.3: perguntar em quais máquinas a produtora roda).
- Limite conhecido: sem a Parte B, duas máquinas que acordem enquanto uma ainda extrai (antes do `importar`) podem extrair juntas. O escalonamento reduz o risco, não o elimina.

### 4.2 Parte B — reserva atômica no banco

- Migração `supabase/migrations/0002_reservas.sql`: tabela `reservas_rodada` (`produtora_id`, `competencia`, `maquina`, `criada_em`, `expira_em`, `status` = `ativa` | `liberada` | `expirada`), **índice único parcial** em (`produtora_id`, `competencia`) onde `status = 'ativa'`, RLS ligado sem políticas e `revoke` para `anon`/`authenticated`, como na 0001.
- Função SQL `reservar_rodada(p_produtora, p_competencia, p_maquina, p_horas)` com o mesmo cuidado de `promover_importacao` (trava consultiva por produtora e competência): expira reservas vencidas, e então reserva ou devolve quem já a tem. A mesma máquina **renova**. O tempo é o do **banco** (`now()`), nunca o relógio das máquinas. Função `liberar_rodada` equivalente.
- **Identidade:** chave `PAPAKI_MAQUINA` no `.env` (nome curto, não é segredo), gravada por `configurar --maquina`; `verificar` confere. Nova coluna `importacoes.maquina` mostra quem importou (a `origem` não muda).
- **Fluxo:**
  1. `agenda --reservar` (as tarefas agendadas passam a usá-lo): com `rodar = true` e banco consultado, tenta reservar. JSON ganha `reserva`: `obtida` | `de_outra_maquina` | `nao_consultada`. Se `de_outra_maquina`, `rodar = false` com o motivo "rodada em andamento em outra máquina". Sem `--reservar`, o `agenda` continua só consultando.
  2. Rodada **manual**: o roteiro reserva pelo mesmo comando; se a reserva for de outra máquina, o Claude avisa e quem decide é o usuário.
  3. O `importar` libera a reserva da própria máquina ao terminar (importada, bloqueada ou falha). Se o processo morrer, a reserva **expira sozinha** (padrão 6 h, como `encerrar_orfas`).
  4. Parada com `PARADA.md`: a reserva **não** é liberada; expira, e a mesma máquina a renova ao retomar no dia seguinte. Assim a principal não extrai do zero uma produtora cuja extração já foi disparada em outra máquina.
- **Banco indisponível:** como hoje, decide a pasta local e avisa o usuário (a reserva fica "não consultada"). A defesa passa a ser só o escalonamento da Parte A.

### 4.3 Parte C — visão do grupo

- `produtora situacao` (somente leitura no banco): para cada produtora cadastrada, a competência alvo, se há foto vigente (e o fim do período), qual máquina importou e se há reserva ativa.
- `verificar`: uma máquina "tem" uma produtora quando há `PAPAKI_MODELO_<ID>` no `.env`; o `verificar` confere, para cada uma delas, o arquivo `produtoras/<id>.json` e o modelo. Não existe lista nova de produtoras por máquina.

## 5. Fora do escopo

Regras de limpeza de tipo novo (cada uma vira plano próprio), layouts de modelo diferentes, orquestrador central das tarefas, mudar as decisões 1–14, e qualquer gravação no Supabase real sem o "sim" do usuário.

## 6. Riscos

- Reserva órfã: coberta pela expiração; o limite de 6 h precisa ser maior que uma extração real (incluindo a espera do e-mail do RRA). Confirmar no ensaio.
- A lógica SQL não é testável só com dublês. Os testes Python provam as chamadas e as decisões; a função SQL é validada num banco **de teste** (projeto ou branch separado) antes de entrar no banco da Papaki, com o OK do usuário.
- A cópia do DRE fica na pasta de dados de quem importou. Se a principal for a rede de segurança de produtoras que o grupo normalmente importa, ela só terá as cópias do que ela própria importou (ver decisão a).

## 7. Testes

Dados sempre inventados. Banco falso para: reserva obtida, negada, renovada pela mesma máquina, expirada, liberada ao fim do `importar` (nos três desfechos), `agenda --reservar` com banco indisponível, `produtora situacao`, `verificar` com várias produtoras, `configurar --maquina`. Teste dos documentos: os comandos citados existem e a regra antiga ("uma máquina por produtora") não aparece mais.

## 8. Decisões do usuário (22/10/2026)

a. A pasta de dados (extrações e apurações) será **compartilhada** entre as máquinas. Cada máquina aponta o próprio `PAPAKI_DADOS_DIR` para ela. Consequências: todas enxergam as cópias do DRE e o estado das extrações; uma extração aberta por uma máquina pode ser retomada por outra depois que a reserva expirar; e duas máquinas **nunca** podem escrever no estado da mesma extração ao mesmo tempo (é para isso que a reserva existe). O `verificar` deve avisar se a pasta de dados estiver dentro do projeto (já faz) e o runbook deve recomendar uma pasta compartilhada com sincronização confiável.
b. Com o banco indisponível, a rodada agendada **segue e avisa** (como hoje).
c. Validade da reserva: **6 h**, a confirmar no ensaio com o tempo real de uma extração.
d. Intervalo entre tarefas na mesma máquina: **a medir no ensaio**.
e. Validação da função SQL da reserva num banco de teste: **autorizada** (a criação do projeto ou branch de teste e qualquer gravação continuam pedindo confirmação no momento de executar).

Risco novo trazido por (a): serviços de sincronização de pasta (nuvem) podem atrasar a chegada de arquivos entre máquinas; por isso a decisão de rodar ou não vem do **banco** (reserva e foto vigente), nunca só da pasta.

## 9. Ordem de entrega

A (textos) → B (reserva, identidade e migração, com ensaio no banco de teste) → C (situação e verificação). Cada parte passa nos testes rápidos e na auditoria de publicação (`scripts/auditar_publicacao.py`) antes do commit.
