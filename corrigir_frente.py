"""
corrigir_frente.py
Corrige o nome da frente de UMA frota no KanbanFrotas — para os casos em que a
frota ficou com o nome longo ("FRENTE 1 - ADEMIR") em vez do nome curto que o
BANCO DE DADOS TONELADAS usa ("ADEMIR").

Modo relatório (não grava nada):
    python corrigir_frente.py 2616 ADEMIR

Aplicando:
    python corrigir_frente.py 2616 ADEMIR --aplicar

O primeiro argumento é o código no início do Title da frota (ex: 2616) ou um
pedaço do nome; o segundo é o nome curto correto da frente.

Também avisa se a frente de destino não existir na KanbanFrentes e se o nome
antigo ficou órfão (nenhuma frota usando) — nesse caso vale apagar a linha da
KanbanFrentes na mão, senão ela continua aparecendo como opção no app.

⚠️  A rotina diária do HistoricoFrenteFrotas vai enxergar essa correção como
    uma troca de frente na data de hoje. Se isso atrapalhar o rateio/custo,
    ajuste também as linhas do histórico dessa frota.
"""
import sys, re, pathlib
import requests

try:
    import tomllib
except ModuleNotFoundError:
    import tomli as tomllib

_SECRETS_PATH = pathlib.Path(__file__).parent / ".streamlit" / "secrets.toml"
if not _SECRETS_PATH.exists():
    print(f"❌ Não encontrei {_SECRETS_PATH}.")
    sys.exit(1)

with open(_SECRETS_PATH, "rb") as _f:
    _secrets = tomllib.load(_f)

TENANT_ID     = _secrets["TENANT_ID"]
CLIENT_ID     = _secrets["CLIENT_ID"]
CLIENT_SECRET = _secrets["CLIENT_SECRET"]
SITE_HOST     = _secrets.get("SITE_HOST", "metalcana.sharepoint.com")
SITE_NAME     = _secrets.get("SITE_NAME", "AppKanbanFrotas")
LISTA_FRENTES = _secrets.get("LISTA_FRENTES", "KanbanFrentes")
LISTA_FROTAS  = _secrets.get("LISTA_FROTAS",  "KanbanFrotas")


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

def listar(token, site_id, lista):
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista}"
           f"/items?expand=fields&$top=2000")
    r = requests.get(url, headers=headers(token))
    r.raise_for_status()
    return r.json().get("value", [])

def atualizar(token, site_id, lista, item_id, campos):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista}/items/{item_id}/fields"
    r = requests.patch(url, headers=headers(token), json=campos)
    if r.status_code not in (200, 201):
        print(f"   ❌ Erro: {r.status_code} — {r.text[:200]}")
        return False
    return True


def main():
    if len(sys.argv) < 3:
        print("Uso: python corrigir_frente.py <código_ou_trecho_do_nome> <FRENTE_CORRETA> [--aplicar]")
        print("Ex.: python corrigir_frente.py 2616 ADEMIR")
        sys.exit(1)

    alvo      = sys.argv[1].strip()
    frente_ok = sys.argv[2].strip()
    aplicar   = "--aplicar" in sys.argv

    token   = get_token()
    site_id = get_site_id(token)

    frotas  = listar(token, site_id, LISTA_FROTAS)
    frentes = [it["fields"].get("Title", "") for it in listar(token, site_id, LISTA_FRENTES)]

    # casa por código no começo do Title (2616 - ...) ou por trecho do nome
    def bate(nome):
        return bool(re.match(rf"^\s*{re.escape(alvo)}\b", nome)) or alvo.lower() in nome.lower()

    achados = [it for it in frotas if bate(it["fields"].get("Title", ""))]

    if not achados:
        print(f"❌ Nenhuma frota casou com '{alvo}'.")
        return
    if len(achados) > 1:
        print(f"⚠️  {len(achados)} frotas casaram com '{alvo}' — refine o argumento:")
        for it in achados:
            print(f"   - {it['fields'].get('Title','')}  (frente atual: "
                  f"{it['fields'].get('FrenteNome','') or '—'})")
        return

    item      = achados[0]
    nome      = item["fields"].get("Title", "")
    frente_at = item["fields"].get("FrenteNome", "") or ""

    print(f"Frota:  {nome}")
    print(f"Frente: '{frente_at or '—'}'  →  '{frente_ok}'")

    if frente_at == frente_ok:
        print("✅ Já está correta, nada a fazer.")
        return

    if frente_ok not in frentes:
        print(f"⚠️  A frente '{frente_ok}' NÃO existe na {LISTA_FRENTES}. "
              f"Cadastre antes, senão ela não aparece como opção no app.")

    ainda_usam = [it for it in frotas
                  if (it["fields"].get("FrenteNome", "") or "") == frente_at
                  and it["id"] != item["id"]]

    if not aplicar:
        print("\nℹ️  Modo relatório — nada foi gravado. Rode de novo com --aplicar.")
        if frente_at and not ainda_usam:
            print(f"    Obs: depois da troca, nenhuma outra frota fica em '{frente_at}' — "
                  f"vale apagar essa linha da {LISTA_FRENTES}.")
        return

    if atualizar(token, site_id, LISTA_FROTAS, item["id"], {"FrenteNome": frente_ok}):
        print("✅ Frente corrigida.")
        if frente_at and not ainda_usam:
            print(f"   Nenhuma outra frota usa '{frente_at}' — apague essa linha da "
                  f"{LISTA_FRENTES} para ela sumir das opções do app.")
        print("   Lembrete: a rotina diária vai registrar isso como troca de frente hoje "
              "no HistoricoFrenteFrotas.")


if __name__ == "__main__":
    main()
