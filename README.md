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
```
4. Deploy — a URL gerada pode ser compartilhada com qualquer pessoa

## Como usar
- **Arrastar card** entre colunas = muda o Centro de Custo
- **Clicar em ⚡ frente** no card = abre modal para mudar a Frente de Corte
- **Exportar Excel** = gera relatório com 3 abas
