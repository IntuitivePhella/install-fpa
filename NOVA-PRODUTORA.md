# Roteiro: cadastrar uma nova produtora (o Claude entrevista o usuário da Papaki)

Leia `CLAUDE.md` e `IMPLANTACAO-PAPAKI.md` antes. **Não adivinhe nada**: toda regra de extração, limpeza, pasta e modelo vem do usuário da Papaki
(registre as respostas). O que já existe hoje (Untitled) mostra o formato: `.venv/Scripts/python.exe -m papaki_dre
produtora mostrar --produtora untitled`. Faça **uma pergunta por vez**, em português simples.

## 1. Identificação

- Nome da produtora e um id curto, minúsculo, sem acento nem espaço (ex.: `soma`).
- Nome **exato** da Localidade da produtora no Esferas. Pergunte se existe uma Localidade de nome parecido que **não**
  pode ser usada (na Untitled, `UTLD PRODUCAO` sem cedilha).
- Endereço do consolidado (hoje `papaki.esferas.com.br`) e endereço do Esferas da produtora (ex.: `soma.esferas.com.br`).

## 2. Extração

"Os quatro relatórios (CP, CR, RRA e Real x Orçado Comissões) são extraídos pelo mesmo caminho de telas que a Untitled?
Em qual Localidade cada um?" Padrão da Untitled: CP e RRA = **Todas as localidades**; CR e Comissões = a Localidade da
produtora. Peça para o usuário mostrar a tela de cada relatório. Se algo for diferente (menu, campo, formato do arquivo),
anote e adapte `runbooks/extracao-esferas.md` para essa produtora. Se o **formato do arquivo** mudar (colunas
diferentes das do `src/papaki_dre/headers.json`), isso exige código novo: **plano próprio**.

## 3. Limpeza (regras de negócio)

Pergunte cada regra e escreva a resposta nas palavras do usuário:

- **RRA:** qual é o Núcleo que entra? (`nucleo_rra`; na Untitled, `Projetos - Untitled`). Há jobs a excluir
  (`PESQUISA E TRATAMENTO`)?
- **CP:** entra só a Localidade da produtora? Excluem-se linhas com número de job preenchido e sem `PESQUISA E TRATAMENTO`?
- **CR:** alguma exclusão de linha? (na Untitled, nenhuma: "Desconsiderar" é classificação, não motivo para apagar.)
- **Comissões:** qual Modelo de Orçamento entra (`modelo_comissoes`) e quais opções da coluna H **não** entram
  (`comissoes_excluir_h`)?

As quatro regras acima cabem no arquivo da produtora. Uma regra de **tipo novo** (outro filtro, outra coluna) **não**
cabe: pare, explique ao usuário e abra um **plano próprio**, com testes. Nunca invente a regra.

## 4. Pastas e Modelo DRE base

- As pastas de extrações e de apurações são as da instalação (`instalacao mostrar`); a produtora ganha subpastas
  próprias automaticamente.
- "Qual é o Modelo DRE base do ano dessa produtora? Quem o mantém? Fica no mesmo lugar?" Anote o caminho completo do
  `.xlsx`. Peça para abrir o arquivo e conferir com você o layout: abas `BASE`, `DESPESAS`, `RECEITA` e
  `REAL X ORÇADO COMISSÕES`; linha do cabeçalho e da primeira linha de dados de cada aba; colunas de fórmula (`AP:AQ`,
  `AC`, `L:N`, `O:Q`) e a **linha do modelo com a fórmula completa de classificação de `DESPESAS!AP`** (`linha_mestre`).
  Compare com `src/papaki_dre/config.py` (`DEST`). Se o layout for diferente, isso exige código novo: **plano próprio**.

## 5. Criar o arquivo da produtora

Crie `produtoras/<id>.json` neste formato (troque os valores pelos da entrevista; `linha_mestre` pode ficar vazio, `{}`,
se a produtora não usa a fórmula completa):

```json
{
  "id": "<id>",
  "nome": "<Nome>",
  "localidade": "<Localidade exata no Esferas>",
  "esferas": {"consolidado": "papaki.esferas.com.br", "produtora": "<id>.esferas.com.br"},
  "regras": {
    "nucleo_rra": "<Núcleo>",
    "modelo_comissoes": "<Modelo de Orçamento>",
    "comissoes_excluir_h": ["<opção 1>", "<opção 2>"]
  },
  "linha_mestre": {"cp": {"AP": <linha>}}
}
```

Depois: `.venv/Scripts/python.exe -m papaki_dre produtora mostrar --produtora <id>` (confira com o usuário) e
`.venv/Scripts/python.exe -m papaki_dre configurar --modelo <id>="<caminho completo do .xlsx>"`. O banco ganha a linha da
produtora sozinho na primeira importação.

**Este repositório é público.** O arquivo da produtora traz só nomes de negócio (Localidade, Núcleo, Modelo): pergunte ao
usuário se podem ficar públicos. Nunca coloque dados de relatórios, e-mails nem caminhos de máquina nele.

## 6. Ensaio, banco e rotina

1. Ensaio **sem banco**, com o usuário olhando: o mesmo do Passo 7 e do "Checklist da primeira execução" do `INSTALAR.md`,
   com `--produtora <id>`. Compare com a DRE que o FP&A já fez à mão do mês anterior **sabendo que ela não é verdade
   linha a linha** (a regra é a que o usuário ditou na entrevista).
2. Gravar no banco: só com o "pode importar" (Passo 8).
3. Pergunte **em quais máquinas** a produtora vai rodar (a principal roda todas; pode haver máquinas de grupo) e crie a
   tarefa agendada **própria** da produtora em cada uma (Passo 9, `runbooks/tarefa-agendada.md`, com horários escalonados),
   só com o OK do usuário. Quem usa cada máquina repete o `configurar --modelo` dela. Depois do primeiro mês, rode
   `.venv/Scripts/python.exe -m papaki_dre produtora situacao` para conferir a produtora entre as demais.
4. Registre as regras novas em `instrucoes_relatorios_ base.md` (seção da produtora), rode a auditoria
   (`.venv/Scripts/python.exe scripts/auditar_publicacao.py`) e os testes, e faça o commit do arquivo da produtora.
