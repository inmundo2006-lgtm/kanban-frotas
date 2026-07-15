# Kanban Frotas — Teston/Metalcana

## Passo a passo completo

### 1. Criar o App no Azure AD
1. Acesse portal.azure.com → Azure Active Directory → App registrations → New registration
2. Nome: `KanbanFrotas`
3. Após criar, copie o **Application (client) ID**
4. Certificates & secrets → New client secret → copie o **Value**
5. API permissions → Add → Microsoft Graph → Application:
   - `Sites.ReadWrite.All`
   - Clique em **Grant admin consent**

### 2. Criar o site no SharePoint
1. Acesse admin.microsoft.com → SharePoint → Sites → Active sites → Create
2. Tipo: **Team site**
3. Nome: `AppKanbanFrotas`
4. URL será: `metalcana.sharepoint.com/sites/AppKanbanFrotas`

### 3. Importar dados
```bash
pip install requests pandas openpyxl
# Edite setup_sharepoint.py com seu CLIENT_ID, CLIENT_SECRET e TENANT_ID completo
python setup_sharepoint.py FROTAS_TESTON.xlsx
```

### 3.1 Módulo "Entrega Futura" (opcional, uma vez)
Cria a lista à parte que registra frotas alocadas hoje mas que serão
entregues/devolvidas no futuro (não mexe no board principal):
```bash
# Edite CLIENT_SECRET nos dois scripts abaixo
python criar_lista_entrega_futura.py
python importar_entrega_futura.py FROTAS_TESTON.xlsx   # importa a aba "ENTREGA FUTURA" do Excel
```

### 3.2 Sincronizar mudanças manuais feitas no Excel
Quando alguém edita o FROTAS_TESTON.xlsx direto (fora do app) — abas novas,
frota mudou de fazenda/frente, chassi corrigido, etc. — rode o comparador
antes de mexer em qualquer coisa manualmente:
```bash
# Edite CLIENT_SECRET no script
python sincronizar_frotas.py FROTAS_TESTON.xlsx              # só mostra o relatório, não muda nada
python sincronizar_frotas.py FROTAS_TESTON.xlsx --aplicar    # aplica as mudanças seguras, pede confirmação "SIM"
```
O script NUNCA apaga ou marca como vendida uma frota sozinho — equipamentos
que sumiram do Excel só aparecem listados no relatório para você decidir
(vender, descartar, ou foi engano de digitação).

### 4. Deploy no Streamlit Cloud
1. Faça push deste projeto para um repositório GitHub (privado)
2. Acesse share.streamlit.io → New app → selecione o repo
3. Em **Secrets**, adicione:
```toml
TENANT_ID     = "seu-tenant-id-completo"
CLIENT_ID     = "seu-client-id"
CLIENT_SECRET = "seu-client-secret"
SITE_HOST     = "metalcana.sharepoint.com"
SITE_NAME     = "AppKanbanFrotas"
LISTA_CC      = "KanbanCC"
LISTA_FRENTES = "KanbanFrentes"
LISTA_FROTAS  = "KanbanFrotas"
LISTA_ENTREGA_FUTURA = "KanbanEntregaFutura"
```
4. Deploy — a URL gerada pode ser compartilhada com qualquer pessoa

## Como usar
- **Arrastar card** entre colunas = muda o Centro de Custo
- **Clicar em ⚡ frente** no card = abre modal para mudar a Frente de Corte
- **Exportar Excel** = gera relatório com 3 abas
- **🚚 Entrega Futura** = registra um aviso de que uma frota (já alocada em
  algum CC/frente) será entregue/devolvida no futuro. É só um aviso à parte:
  cadastrar ou dar baixa aqui **não move** a frota no board principal.
  Ao "Marcar entregue", o registro sai da lista de pendentes e vai para o
  histórico.
