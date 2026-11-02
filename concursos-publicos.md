# Projeto: Monitor de Concursos Públicos

## Objetivo
Script Python que recolhe anúncios de concursos públicos abertos em Portugal e
gera uma tabela Excel simples (sem formatação), para triagem rápida de
oportunidades de obra.

## Âmbito e limites (ler antes de tudo)
- Só apanha procedimentos COM publicidade: concurso público e concurso limitado
  por prévia qualificação. Estes são obrigatoriamente publicados no
  Diário da República (2.ª série, Parte L).
- NÃO apanha ajustes diretos nem consultas prévias: são por convite e não são
  anunciados publicamente. Nenhuma API resolve isto.
- As plataformas eletrónicas (Vortal, SaphetyGov, acinGov, etc.) são onde o
  procedimento tramita, não a fonte primária do anúncio. Não são o alvo inicial.

## Fontes (por ordem de prioridade)
| Fonte | Cobertura | Acesso | Estado |
|---|---|---|---|
| Diário da República (Parte L) | Todos os concursos públicos nacionais | Sem API documentada; RSS / site | A investigar |
| TED (JOUE) | Só acima dos limiares europeus | API oficial v3, pesquisa anónima | Confirmado |
| Portal BASE | Anúncios + contratos | Dataset dados.gov.pt (desatualizado); site sem API documentada | Pouco útil para tempo real |

## Fase 1 — Viabilidade de acesso (fazer primeiro)
1. DR: testar RSS da 2.ª série e ver se dá para filtrar Parte L / anúncios de procedimento.
   Se não, avaliar pesquisa do site (pedidos HTTP) e a robustez disso.
2. TED: chamada à Search API filtrando país = PRT e CPV 45* (obras). Validar campos.
3. BASE: verificar se a pesquisa pública de anúncios tem endpoint utilizável.
4. Para cada fonte registar: URL, método, formato de resposta, campos úteis,
   limites de uso, termos de utilização.

## Fase 2 — Plataformas
Só depois da Fase 1. Objetivo: saber em que plataforma corre cada procedimento
(campo "plataforma eletrónica" já vem no anúncio do DR), não fazer scraping delas.

### Acesso público às peças (testado 2026-11-01, sem login)
- acinGov: o "Link peças" descarrega um ZIP com tudo (pode ter centenas de MB).
- anoGov / ComprasPT: página "List of Documents"; cada ficheiro tem link direto
  `.../decryptservlet?papId=..&docId=..&fichId=..`.
- Vortal: `GET /public/api/PublicTenderDocuments/GetPublicTenderInformation?uniqueIdentifierEncrypted=<fim do link>`
  dá `contractNoticeUrl` (termina no UId, ex. PT1.NTC.3836841); depois
  `GET /public/api/ContractNoticeDetail/GetContractNoticeDocuments?contractNoticeUId=<UId>`
  dá a lista com `downloadUrl` por ficheiro.
- Resumo dos trabalhos: não existe campo estruturado (DR e BASE só repetem o
  título). A "Cláusula 1.ª (Objeto)" do Caderno de Encargos quase sempre também
  só repete o título; alguns cadernos são PDF digitalizado (sem texto) e nem
  sempre se chamam "Caderno" (ex. `..._CE.pdf`, ou dentro do ZIP do projeto).

## Output Excel (colunas)
Data publicação | Nº anúncio | Entidade adjudicante | Objeto |
Tipo de obra (CPV) ("45233140 - Obras em estradas"; CPV distintos dos lotes com ";") |
Preço base (€) | Prazo propostas | Prazo execução | Local ("Concelho (Distrito)") |
Alvará exigido (texto da secção 12 do anúncio, máx. 500 car.) |
Plataforma | Link peças | PDF DR

## Filtros pretendidos
- CPV 45* (trabalhos de construção), 507* (manutenção de instalações de
  edifícios), 5023* (manutenção de estradas e equipamentos afins),
  5051* (bombas, válvulas, torneiras), 50870* (equipamento de recreio
  público), 7731* (espaços verdes), 90650* (amianto), 51* (instalação),
  44211* (pré-fabricados), 44221* (janelas e portas), 34928* (mobiliário
  rodoviário), 90722* (recuperação ambiental), 71* (arquitetura/engenharia),
  92522* (monumentos históricos), 50225* (vias férreas), 50246* (equipamento
  portuário), 50413200 (equipamento contra incêndios), 77211500 (árvores) e
  77320* (campos de desporto);
  códigos confirmados na lista oficial CPV 2008 (ted.europa.eu);
  ou tipo de contrato "Empreitada", qualquer que seja o CPV
- Zona geográfica: Portugal inteiro (sem filtro)
- Preço base: sem mínimo / sem máximo (recolher todas as obras)

## Convenções técnicas
- Python 3.x; requests para HTTP; pandas + openpyxl para o Excel
- Sem formatação no Excel (só cabeçalho e dados), exceto a coluna
  Preço base em formato contabilidade
- Deduplicar por nº de anúncio (o mesmo concurso pode aparecer no DR e no TED)
- Base única concursos_base.xlsx, atualizada a cada execução (deduplicada pelo
  link do PDF do DR); a versão anterior fica em concursos_base.bak.xlsx
- Respeitar termos de uso dos sites; pausas entre pedidos

## Decisões tomadas
- 2026-11-01: removido o mínimo de preço base (era 100 000 €); recolhem-se
  todas as obras, independentemente do valor.
- 2026-11-01: removido o filtro de zona (era Porto, Aveiro, Braga); recolhem-se
  obras de todo o país.
- 2026-11-01: ficheiros diários substituídos por uma base única
  (concursos_base.xlsx); os ficheiros diários antigos são importados na 1.ª execução.
- 2026-11-01: recolha única (recolha_base.py) das obras com prazo aberto
  publicadas nos 180 dias anteriores. A lista vem da pesquisa de anúncios do
  Portal BASE (POST https://www.base.gov.pt/Base4/pt/resultados/,
  type=search_anuncios / detail_anuncios; máx. 50 por página). O detalhe dá o
  link do PDF no DR (campo "reference"). O site do DR não tem listagem de
  diários antigos utilizável (files.diariodarepublica.pt não lista pastas).
  A firewall do BASE (WebKnight, HTTP 999) bloqueia com 1 s entre pedidos;
  usar pelo menos 5 s.

## Questões em aberto
- Frequência de execução (diária?) e onde corre (PC pessoal vs agendado na cloud)