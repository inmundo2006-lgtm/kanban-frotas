"""
sincronizar_frotas.py

⚠️  ATENÇÃO — LEIA ANTES DE USAR
    O KanbanFrotas virou a FONTE ÚNICA do cadastro de frotas. O
    FROTAS_TESTON.xlsx está com os nomes de frente no formato longo
    ("FRENTE 1 - ADEMIR") enquanto a fonte correta dos nomes é o BANCO DE DADOS
    TONELADAS, que usa o nome curto (DANILO, BOCA, ADILIO, MACIEL, ADEMIR).
    Rodar este script com --aplicar sobrescreve frente/CC bons com dados ruins.
    Use SÓ em modo relatório até a planilha ser corrigida.

Compara o FROTAS_TESTON.xlsx (editado manualmente pelo gestor) com o que
está hoje nas listas SharePoint (KanbanCC, KanbanFrentes, KanbanFrotas) e
mostra as diferenças, SEM aplicar nada.

Rode assim primeiro (modo relatório, não mexe em nada):
    python sincronizar_frotas.py FROTAS_TESTON.xlsx

Depois de revisar o relatório, para aplicar as mudanças seguras
(novos CCs, novas frentes, novas frotas, frotas que mudaram de CC/frente,
correções de chassi/ano) rode:
    python sincronizar_frotas.py FROTAS_TESTON.xlsx --aplicar

Frotas que SUMIRAM do Excel (existem no SharePoint mas não aparecem mais
em nenhuma aba) NUNCA são apagadas ou movidas automaticamente — só listadas
no relatório para você decidir (vender, descartar, ou é engano de digitação).
"""
import sys, re, time, json, pathlib, unicodedata
import requests
import pandas as pd

try:
    import tomllib  # Python 3.11+
except ModuleNotFoundError:
    import tomli as tomllib  # pip install tomli --break-system-packages (Python < 3.11)

# ── CONFIGURAÇÃO ─────────────────────────────────────────────────────────────
# As credenciais NÃO ficam mais neste arquivo (evita vazar no GitHub).
# Elas são lidas do mesmo .streamlit/secrets.toml usado pelo app.py.
_SECRETS_PATH = pathlib.Path(__file__).parent / ".streamlit" / "secrets.toml"
if not _SECRETS_PATH.exists():
    print(f"❌ Não encontrei {_SECRETS_PATH}.")
    print("   Crie o arquivo .streamlit/secrets.toml (mesmo formato usado no Streamlit Cloud)")
    print("   com TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_HOST, SITE_NAME, LISTA_CC, LISTA_FRENTES, LISTA_FROTAS.")
    sys.exit(1)

with open(_SECRETS_PATH, "rb") as _f:
    _secrets = tomllib.load(_f)

TENANT_ID     = _secrets["TENANT_ID"]
CLIENT_ID     = _secrets["CLIENT_ID"]
CLIENT_SECRET = _secrets["CLIENT_SECRET"]
SITE_HOST     = _secrets.get("SITE_HOST", "metalcana.sharepoint.com")
SITE_NAME     = _secrets.get("SITE_NAME", "AppKanbanFrotas")
LISTA_CC      = _secrets.get("LISTA_CC",      "KanbanCC")
LISTA_FRENTES = _secrets.get("LISTA_FRENTES", "KanbanFrentes")
LISTA_FROTAS  = _secrets.get("LISTA_FROTAS",  "KanbanFrotas")

# ── AUTH / GRAPH ──────────────────────────────────────────────────────────────
def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
              "client_secret": CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"})
    r.raise_for_status()
    return r.json()["access_token"]

def headers(token):
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

def get_site_id(token):
    r = requests.get(f"https://graph.microsoft.com/v1.0/sites/{SITE_HOST}:/sites/{SITE_NAME}",
                      headers=headers(token))
    r.raise_for_status()
    return r.json()["id"]

def get_lista_id(token, site_id, nome):
    r = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$select=id,name",
                      headers=headers(token))
    r.raise_for_status()
    return next(l["id"] for l in r.json()["value"] if l["name"] == nome)

def listar_items(token, site_id, lista_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista_id}/items?expand=fields&$top=2000"
    r = requests.get(url, headers=headers(token))
    r.raise_for_status()
    return r.json().get("value", [])

def inserir_item(token, site_id, lista_id, fields):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista_id}/items"
    r = requests.post(url, headers=headers(token), json={"fields": fields})
    if r.status_code not in (200, 201):
        print(f"    ❌ Erro ao inserir: {r.status_code} — {r.text[:150]}")
        return None
    return r.json()["id"]

def atualizar_item(token, site_id, lista_id, item_id, fields):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista_id}/items/{item_id}/fields"
    r = requests.patch(url, headers=headers(token), json=fields)
    if r.status_code not in (200, 201):
        print(f"    ❌ Erro ao atualizar: {r.status_code} — {r.text[:150]}")
        return False
    return True

# ── PARSER DO EXCEL ───────────────────────────────────────────────────────────
# A partir da reestruturação do FROTAS_TESTON.xlsx (planilha passou a ter uma
# aba BASE_DADOS consolidada, com colunas nomeadas, em vez de uma aba por
# CC/fazenda com layout de colunas fixo), o parser lê direto essa aba.
# Mapeamento confirmado com o usuário (planilha reestruturada de novo):
#   - "CENTRO DE CUSTO" → Centro de Custo (CC), já no formato "código - nome"
#                          (ex: "038 - AGRO ASTORGA")
#   - "FRENTE/ GRUPO"   → Frente de Corte
# O código do CC na planilha vem às vezes com zero à esquerda e espaçamento
# diferente do que já está salvo no SharePoint (ex: "038 - AGRO ASTORGA" vs
# "38 - AGRO ASTORGA", ou "041 - COCAL" vs "41-COCAL"). Por isso o matching de
# CC usa normalizar_cc_chave() abaixo, que ignora zeros à esquerda e espaços
# ao redor do hífen — evita "CC novo" falso e "mudou de CC" falso por causa
# só de formatação.
SKIP = {'', 'nan', 'NaN', 'None', 'F', 'none', 'NAN', 'DUAL', 'S/I',
        'AGUARDANDO', 'AGUARDANDO DADOS'}

# Situações que NÃO entram no comparador de frotas ativas (equivalente às
# antigas abas VENDIDO/DESCARTE/DESTINADO A VENDA, que eram puladas inteiras).
# Os mesmos valores valem agora para a coluna Status do KanbanFrotas, que
# passou a aceitar "Destinado a venda" e "Descarte" além de "Vendido".
SITUACOES_INATIVAS = {'VENDIDO', 'DESTINADO À VENDA', 'DESTINADO A VENDA',
                      'DESCARTE', 'DESCARTADO', 'BAIXADO'}

def chave_situacao(s):
    """Compara sem acento/caixa: 'Destinado à Venda' == 'DESTINADO A VENDA'."""
    s = unicodedata.normalize("NFKD", str(s or "").strip()).encode("ascii", "ignore").decode()
    return " ".join(s.upper().split())

CHAVES_INATIVAS = {chave_situacao(x) for x in SITUACOES_INATIVAS}

def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    if isinstance(v, float) and v == int(v):
        s = str(int(v))
    else:
        s = str(v).strip()
    return '' if s.upper() in {x.upper() for x in SKIP} else s

# Mapa direto Categoria (nova planilha) → Tipo usado no board (KanbanFrotas).
# Cobre as 24 categorias vistas em BASE_DADOS/Aba Original.
CATEGORIA_TIPO = {
    'Colhedora': 'Colhedora', 'Colheitadeira de Grãos': 'Colhedora',
    'Transbordo': 'Transbordo',
    'Trator': 'Trator', 'Carregadeira': 'Trator', 'Máquina Amarela': 'Trator', 'Pulverizador': 'Trator',
    'Caminhão': 'Caminhão', 'Ônibus': 'Caminhão',
    'Carro': 'Veículo', 'Moto': 'Veículo',
    'Implemento': 'Implemento', 'Carreta': 'Implemento', 'Bazuca': 'Implemento',
    'Tanque': 'Implemento', 'Empilhadeira': 'Implemento',
    'Baú Oficina': 'Apoio', 'Área de Vivência': 'Apoio', 'Motobomba': 'Apoio',
    'Motor Estacionário': 'Apoio', 'Gerador': 'Apoio', 'Drone': 'Apoio', 'Helicóptero': 'Apoio',
    'Motocana': 'Trator',
}

def tipo_eq(categoria, nome):
    """Tipo pela Categoria da planilha; cai para heurística por nome se a
    categoria vier vazia ou fora do mapa (nunca deveria acontecer com a
    BASE_DADOS, mas evita perder frota por categoria não mapeada)."""
    if categoria and categoria in CATEGORIA_TIPO:
        return CATEGORIA_TIPO[categoria]
    n = (nome or '').upper()
    if re.search(r'\b(CH570|CH670|CH950|CH3520|CH3522|JD35\d\d|AGNES|COLHEITADEIRA)\b', n): return 'Colhedora'
    if re.search(r'\b(TRANSBORDO|GIGANTE 22)\b', n): return 'Transbordo'
    if re.search(r'\b(TRATOR|CARREGADEIRA|PULVERIZADOR|AUTOPROPELIDO)\b', n): return 'Trator'
    if re.search(r'\b(CAMINHAO|CAMINHÃO|ONIBUS|ÔNIBUS)\b', n): return 'Caminhão'
    if re.search(r'\b(CARRO|MOTO)\b', n): return 'Veículo'
    if re.search(r'\b(EMPILHADEIRA|CARRETA|BAZUCA|TANQUE|IMPLEMENTO|CULTIVADOR|PLANTADEIRA)\b', n): return 'Implemento'
    if re.search(r'\b(BAU|BAÚ|VIVENCIA|VIVÊNCIA|MOTOBOMBA|GERADOR|DRONE|HELICOPTERO|HELICÓPTERO)\b', n): return 'Apoio'
    return 'Outro'

def extrair_dados(caminho_xlsx):
    """Lê a aba BASE_DADOS (consolidada, colunas nomeadas). Cada linha vira
    uma frota; CC = "Cliente / Fazenda", Frente = "Frente / Grupo"."""
    df = pd.read_excel(caminho_xlsx, sheet_name="BASE_DADOS")
    ccs = set(); frentes = set(); frotas = []
    for _, row in df.iterrows():
        situacao = clean(row.get('Situação', ''))
        if chave_situacao(situacao) in CHAVES_INATIVAS:
            continue  # vendida/destinada a venda/descartada não entra no comparador de ativas

        numero = clean(row.get('Nº Frota', ''))
        descricao = clean(row.get('Descrição (Veículo)', ''))
        nome = f"{numero} - {descricao}" if numero and descricao else (descricao or numero)
        if not nome:
            continue

        cc_nome = clean(row.get('CENTRO DE CUSTO', ''))
        fr_nome = clean(row.get('FRENTE/ GRUPO', ''))
        categoria = clean(row.get('Categoria', ''))

        if cc_nome:
            ccs.add(cc_nome)
        if fr_nome:
            frentes.add(fr_nome)

        frotas.append({
            'nome': nome,
            'tipo': tipo_eq(categoria, descricao),
            'chassi': clean(row.get('Chassi', '')),
            'ano': clean(row.get('Ano Fabricação', '')),
            'cc': cc_nome,
            'frente': fr_nome,
            'obs': clean(row.get('Observação', '')),
        })
    return sorted(ccs), sorted(frentes), frotas


# ── MATCHING (chassi limpo → nome → código numérico do equipamento) ──────────
def chave_chassi(c):
    if not c: return ''
    return str(c).strip().split('/')[0].strip().upper()

def chave_nome(n):
    return re.sub(r'\s+', ' ', str(n).strip().upper())

def normalizar_cc_chave(cc):
    """Chave de comparação para CC que ignora zero à esquerda no código e
    espaçamento ao redor do hífen (ex: '038 - AGRO ASTORGA', '38 - AGRO
    ASTORGA' e '41-COCAL'/'041 - COCAL' caem na mesma chave). Não é o valor
    gravado no SharePoint — só usado para decidir se é CC novo/mudou."""
    s = str(cc).strip()
    m = re.match(r'^0*(\d+)\s*-\s*(.+)$', s)
    if m:
        codigo, nome = m.groups()
        return f"{codigo}-{chave_nome(nome)}"
    return chave_nome(s)

def codigo_prefixo(nome):
    m = re.match(r'^\s*(\d+)', str(nome))
    return m.group(1) if m else None

def montar_indices(frotas, chassi_field, nome_field):
    by_chassi, by_nome, by_codigo = {}, {}, {}
    for f in frotas:
        ck = chave_chassi(f.get(chassi_field, ''))
        if ck:
            by_chassi.setdefault(ck, []).append(f)
        by_nome.setdefault(chave_nome(f.get(nome_field, '')), []).append(f)
        cod = codigo_prefixo(f.get(nome_field, ''))
        if cod:
            by_codigo.setdefault(cod, []).append(f)
    return by_chassi, by_nome, by_codigo

def achar_correspondente(f, by_chassi, by_nome, by_codigo, chassi_field, nome_field):
    ck = chave_chassi(f.get(chassi_field, ''))
    if ck and ck in by_chassi and len(by_chassi[ck]) == 1:
        return by_chassi[ck][0], "chassi"
    nk = chave_nome(f.get(nome_field, ''))
    if nk in by_nome and len(by_nome[nk]) == 1:
        return by_nome[nk][0], "nome"
    cod = codigo_prefixo(f.get(nome_field, ''))
    if cod and cod in by_codigo and len(by_codigo[cod]) == 1:
        return by_codigo[cod][0], "codigo"
    return None, None

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print("Uso: python sincronizar_frotas.py FROTAS_TESTON.xlsx [--aplicar]")
        sys.exit(1)
    aplicar = "--aplicar" in sys.argv

    print("⚠️  KanbanFrotas é a fonte única do cadastro de frotas.")
    print("    Este script serve para CONFERIR o Excel contra o Kanban.")
    print("    Nomes de frente corretos vêm do BANCO DE DADOS TONELADAS (nome curto).\n")

    print("🔐 Autenticando...")
    token = get_token()
    site_id = get_site_id(token)
    print("✅ OK\n")

    id_cc  = get_lista_id(token, site_id, LISTA_CC)
    id_fr  = get_lista_id(token, site_id, LISTA_FRENTES)
    id_ft  = get_lista_id(token, site_id, LISTA_FROTAS)

    print("📂 Lendo listas do SharePoint (estado atual)...")
    ccs_sp_raw     = listar_items(token, site_id, id_cc)
    frentes_sp_raw = listar_items(token, site_id, id_fr)
    frotas_sp_raw  = listar_items(token, site_id, id_ft)

    ccs_sp     = [it["fields"].get("Title", "") for it in ccs_sp_raw]
    frentes_sp = [it["fields"].get("Title", "") for it in frentes_sp_raw]
    frotas_sp  = [{
        "id":     it["id"],
        "nome":   it["fields"].get("Title", ""),
        "chassi": it["fields"].get("Chassi", ""),
        "ano":    it["fields"].get("Ano", ""),
        "obs":    it["fields"].get("Obs", ""),
        "cc":     it["fields"].get("CCNome", ""),
        "frente": it["fields"].get("FrenteNome", ""),
        "status": it["fields"].get("Status", "Ativo"),
    } for it in frotas_sp_raw]
    # só compara frotas ativas — vendidas, destinadas a venda e descartadas
    # ficam de fora do diff (antes só "Vendido" era excluído)
    frotas_sp_ativas = [f for f in frotas_sp
                        if chave_situacao(f.get("status", "Ativo")) not in CHAVES_INATIVAS]

    print(f"   {len(ccs_sp)} CCs | {len(frentes_sp)} frentes | {len(frotas_sp_ativas)} frotas ativas\n")

    print(f"📖 Lendo {sys.argv[1]}...")
    ccs_xl, frentes_xl, frotas_xl = extrair_dados(sys.argv[1])
    print(f"   {len(ccs_xl)} CCs | {len(frentes_xl)} frentes | {len(frotas_xl)} frotas\n")

    # Reconcilia grafia do CC com o que já existe no SharePoint (mesmo CC,
    # formatação diferente: zero à esquerda, espaço no hífen etc.) ANTES de
    # calcular os diffs — assim "038 - AGRO ASTORGA" (Excel) não vira um CC
    # novo/mudança falsa contra "38 - AGRO ASTORGA" (SharePoint).
    ccs_sp_norm = {normalizar_cc_chave(c): c for c in ccs_sp}
    def canon_cc(c):
        return ccs_sp_norm.get(normalizar_cc_chave(c), c)
    for f in frotas_xl:
        f['cc'] = canon_cc(f['cc'])
    ccs_xl = sorted({canon_cc(c) for c in ccs_xl})

    # ── DIFF: CCs e Frentes novos ──────────────
    novos_ccs     = [c for c in ccs_xl if c not in ccs_sp]
    novas_frentes = [f for f in frentes_xl if f not in frentes_sp]

    # ── DIFF: Frotas ────────────────────────────
    by_chassi_sp, by_nome_sp, by_codigo_sp = montar_indices(frotas_sp_ativas, "chassi", "nome")

    # Duplicatas dentro do próprio SharePoint (mesmo chassi em 2+ registros) —
    # atrapalham o matching e são um problema de qualidade de dado à parte.
    duplicatas_sp = {ck: v for ck, v in by_chassi_sp.items() if len(v) > 1}

    novas_frotas, mudou_local, mudou_dados, sem_mudanca = [], [], [], []
    xl_casados_ids = set()
    match_por_metodo = {"chassi": 0, "nome": 0, "codigo": 0}

    for f in frotas_xl:
        m, metodo = achar_correspondente(f, by_chassi_sp, by_nome_sp, by_codigo_sp, "chassi", "nome")
        if not m:
            novas_frotas.append(f)
            continue
        match_por_metodo[metodo] += 1
        xl_casados_ids.add(m["id"])
        mudancas = {}
        if chave_nome(m["cc"]) != chave_nome(f["cc"]):
            mudancas["cc"] = (m["cc"], f["cc"])
        if chave_nome(m.get("frente", "")) != chave_nome(f.get("frente", "")):
            mudancas["frente"] = (m.get("frente", ""), f.get("frente", ""))
        if mudancas:
            mudou_local.append({"sp": m, "xl": f, "mudancas": mudancas})
        dif_dados = {}
        if chave_chassi(m.get("chassi", "")) != chave_chassi(f.get("chassi", "")) and f.get("chassi"):
            dif_dados["chassi"] = (m.get("chassi", ""), f.get("chassi", ""))
        if str(m.get("ano", "")).strip() != str(f.get("ano", "")).strip() and f.get("ano"):
            dif_dados["ano"] = (m.get("ano", ""), f.get("ano", ""))
        if dif_dados:
            mudou_dados.append({"sp": m, "xl": f, "mudancas": dif_dados})
        if not mudancas and not dif_dados:
            sem_mudanca.append(f)

    removidas = [f for f in frotas_sp_ativas if f["id"] not in xl_casados_ids]

    # ── RELATÓRIO ────────────────────────────────
    print("=" * 70)
    print("RELATÓRIO DE COMPARAÇÃO — Excel  vs.  SharePoint (ao vivo)")
    print("=" * 70)

    print(f"\n🆕 Novos Centros de Custo no Excel ({len(novos_ccs)}):")
    for c in novos_ccs: print(f"   - {c}")

    print(f"\n🆕 Novas Frentes no Excel ({len(novas_frentes)}):")
    for f in novas_frentes: print(f"   - {f}")

    print(f"\n🆕 Frotas novas no Excel, não existem no SharePoint ({len(novas_frotas)}):")
    for f in sorted(novas_frotas, key=lambda x: x["cc"])[:200]:
        print(f"   - {f['nome']}  ({f['tipo']})  chassi={f.get('chassi','—')}  → CC: {f['cc']} / {f.get('frente') or '—'}")

    print(f"\n🔀 Frotas que mudaram de CC/Frente ({len(mudou_local)}):")
    for m in mudou_local[:200]:
        f = m["xl"]
        partes = []
        if "cc" in m["mudancas"]:
            de, para = m["mudancas"]["cc"]
            partes.append(f"CC: '{de or '—'}' → '{para or '—'}'")
        if "frente" in m["mudancas"]:
            de, para = m["mudancas"]["frente"]
            partes.append(f"Frente: '{de or '—'}' → '{para or '—'}'")
        print(f"   - {f['nome']}: " + " | ".join(partes))

    print(f"\n✏️  Frotas com chassi/ano diferentes ({len(mudou_dados)}):")
    for m in mudou_dados[:200]:
        f = m["xl"]
        partes = []
        if "chassi" in m["mudancas"]:
            de, para = m["mudancas"]["chassi"]
            partes.append(f"Chassi: '{de}' → '{para}'")
        if "ano" in m["mudancas"]:
            de, para = m["mudancas"]["ano"]
            partes.append(f"Ano: '{de}' → '{para}'")
        print(f"   - {f['nome']}: " + " | ".join(partes))

    print(f"\n⚠️  Frotas que SUMIRAM do Excel — existem no SharePoint mas não foram "
          f"encontradas em nenhuma aba atual ({len(removidas)}):")
    print("    (NÃO alteradas automaticamente — decida manualmente: vendida? descartada? erro de digitação?)")
    for f in sorted(removidas, key=lambda x: x["cc"])[:200]:
        print(f"   - {f['nome']}  (estava em: {f['cc']} / {f.get('frente') or '—'})  chassi={f.get('chassi','—')}")

    print(f"\n✅ Sem mudanças: {len(sem_mudanca)} frotas")
    print(f"   (casadas por chassi: {match_por_metodo['chassi']} | por nome: {match_por_metodo['nome']} | "
          f"por código numérico: {match_por_metodo['codigo']})")

    if duplicatas_sp:
        print(f"\n🚨 ATENÇÃO — chassi duplicado dentro do próprio KanbanFrotas ({len(duplicatas_sp)} chassis, "
              f"envolvendo {sum(len(v) for v in duplicatas_sp.values())} registros):")
        print("    Isso é um problema de qualidade de dado independente deste sync — provavelmente")
        print("    entrou duplicado em alguma migração anterior. Não mexi automaticamente. Registros:")
        for ck, v in list(duplicatas_sp.items())[:30]:
            nomes = ", ".join(f"{x['nome']} (id {x['id']})" for x in v)
            print(f"   - chassi '{ck}': {nomes}")
    print("\n" + "=" * 70)
    print(f"RESUMO: {len(novos_ccs)} CCs novos | {len(novas_frentes)} frentes novas | "
          f"{len(novas_frotas)} frotas novas | {len(mudou_local)} mudaram de lugar | "
          f"{len(mudou_dados)} dados corrigidos | {len(removidas)} sumiram (revisar manualmente)")
    print("=" * 70)

    if not aplicar:
        print("\nℹ️  Modo relatório — nada foi alterado no SharePoint.")
        print("    Para aplicar as mudanças seguras (novos CCs/frentes/frotas, "
              "mudanças de local e correções de dados), rode de novo com --aplicar")
        return

    print("\n⚠️  MODO APLICAR — isso vai gravar no SharePoint (produção).")
    print("    Itens da lista 'sumiram do Excel' NÃO serão tocados — só as demais categorias.")
    print("\n🚨 O KanbanFrotas é a fonte única do cadastro. Se o Excel ainda estiver com")
    print("   nome de frente no formato longo ('FRENTE 1 - ADEMIR'), este comando vai")
    print("   sobrescrever as frentes certas (nome curto: DANILO, BOCA, ADILIO, MACIEL).")
    resp = input('    Digite "SIM, APLICAR" para confirmar: ').strip()
    if resp != "SIM, APLICAR":
        print("Cancelado.")
        return

    # Novos CCs
    print("\n⬆️  Criando novos CCs...")
    ordem_atual = len(ccs_sp)
    for c in novos_ccs:
        ordem_atual += 1
        item_id = inserir_item(token, site_id, id_cc, {"Title": c, "Codigo": "", "Cor": "", "Ordem": ordem_atual})
        if item_id: print(f"   ✅ {c}")
        time.sleep(0.2)

    # Novas frentes
    print("\n⬆️  Criando novas frentes...")
    for f in novas_frentes:
        item_id = inserir_item(token, site_id, id_fr, {"Title": f, "CCNome": ""})
        if item_id: print(f"   ✅ {f}")
        time.sleep(0.15)

    # Novas frotas
    print("\n⬆️  Inserindo frotas novas...")
    for f in novas_frotas:
        inserir_item(token, site_id, id_ft, {
            "Title": f["nome"], "Tipo": f["tipo"], "Chassi": f.get("chassi",""),
            "Ano": f.get("ano",""), "Obs": f.get("obs",""),
            "CCNome": f["cc"], "FrenteNome": f.get("frente",""),
            "Status": "Ativo", "DataVenda": "", "ValorVenda": "",
        })
        time.sleep(0.15)
    print(f"   ✅ {len(novas_frotas)} frotas inseridas")

    # Mudanças de local
    print("\n🔀 Atualizando CC/Frente das frotas que mudaram...")
    for m in mudou_local:
        atualizar_item(token, site_id, id_ft, m["sp"]["id"], {
            "CCNome": m["xl"]["cc"], "FrenteNome": m["xl"].get("frente",""),
        })
        time.sleep(0.15)
    print(f"   ✅ {len(mudou_local)} frotas atualizadas")

    # Correções de chassi/ano
    print("\n✏️  Corrigindo chassi/ano divergentes...")
    for m in mudou_dados:
        campos = {}
        if "chassi" in m["mudancas"]: campos["Chassi"] = m["xl"]["chassi"]
        if "ano" in m["mudancas"]:    campos["Ano"] = m["xl"]["ano"]
        atualizar_item(token, site_id, id_ft, m["sp"]["id"], campos)
        time.sleep(0.15)
    print(f"   ✅ {len(mudou_dados)} frotas corrigidas")

    print(f"\n✅ Sincronização aplicada! {len(removidas)} frotas continuam como estavam "
          f"(revise a lista de 'sumiram do Excel' manualmente no app se precisar vender/descartar).")


if __name__ == "__main__":
    main()