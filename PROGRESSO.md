# Progresso do Gestor da Nave — 30/06/2026

---

## ✅ O que foi feito hoje

### 1. Frontend (`index.html`)
- **Filtros nas páginas DB1/DB2** — adicionados 3 dropdowns no `section-header` (ao lado do contador):
  - 📅 **Data**: Este Mês / Próximo Mês / Últimos 7 Dias / Próximos 7 Dias
  - 👨‍🏫 **Instrutor/Agente**: "Todos" + lista dinâmica
  - ✔️ **Tipo**: Todos / PALESTRA / TREINAMENTO / REUNIÃO / CONSULTORIA / OUTRA
- **Variáveis de estado** adicionadas: `currentFilterData`, `currentFilterInstrutor`, `currentFilterTipo`
- **CSS inline** nos botões/dropdowns (funciona, pode extrair depois)

### 2. Descoberta da API real do SAN
| Item | Valor |
|------|-------|
| **Endpoint** | `https://api.navedoconhecimento.rio/consolidated_number/search` |
| **Método** | `GET` com query `q=<JSON URL-encoded>` |
| **Auth** | `Authorization: Bearer <JWT>` (OAuth2) |
| **CORS** | `access-control-allow-origin: *` ✅ |
| **Exemplo payload** | Ver abaixo |

**Payload decodificado que funciona:**
```json
{
  "page": 1,
  "max": 100,
  "filters": {
    "conditions": [
      { "field": "unit_id", "value": "-1" },
      { "type": "contains", "field": "type", "value": "indicador_%" },
      { "type": "greater-than-or-equal", "field": "start_expected", "value": "2026-06-01 00:00:00" },
      { "type": "less-than-or-equal", "field": "end_expected", "value": "2026-06-30 23:59:59" }
    ]
  }
}
```

### 3. Como testar AGORA (no console do browser logado no painel)
```js
const token = 'COLE_O_TOKEN_DO_HEADER_AUTHORIZATION_AQUI';

fetch('https://api.navedoconhecimento.rio/consolidated_number/search?language=pt&q=' + encodeURIComponent(JSON.stringify({
  page: 1, max: 50,
  filters: { conditions: [
    { field: 'unit_id', value: '-1' },
    { type: 'contains', field: 'type', value: 'indicador_%' },
    { type: 'greater-than-or-equal', field: 'start_expected', value: '2026-06-01 00:00:00' },
    { type: 'less-than-or-equal', field: 'end_expected', value: '2026-06-30 23:59:59' }
  ]}
})), {
  headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' }
}).then(r => r.json()).then(d => { console.log(d); copy(JSON.stringify(d, null, 2)); });
```

---

## 🔄 Próximos passos (para continuar amanhã/depois)

### A. Integrar a API no `index.html`
1. **Criar Cloudflare Worker / Vercel Edge Function** que:
   - Guarda o JWT em secret/env (não no frontend)
   - Recebe parâmetros do frontend (`page`, `max`, `filters`)
   - Adiciona `Authorization: Bearer <token>`
   - Faz fetch na API real
   - Devolve JSON limpo pro frontend
2. **No `index.html`**: substituir `carregar()` / `fetchJSONP` por chamada ao teu Worker

### B. Worker exemplo (Cloudflare)
```js
// wrangler.toml: vars = { API_BASE = "https://api.navedoconhecimento.rio" }
// Secrets: JWT_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9..."

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    const q = url.searchParams.get('q') || '{"page":1,"max":50,"filters":{"conditions":[]}}';
    
    const resp = await fetch(`${env.API_BASE}/consolidated_number/search?language=pt&q=${encodeURIComponent(q)}`, {
      headers: {
        'Authorization': `Bearer ${env.JWT_TOKEN}`,
        'Accept': 'application/json'
      }
    });
    
    const data = await resp.json();
    return new Response(JSON.stringify(data), {
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
    });
  }
}
```

### C. Ajustar filtros do frontend
- Conectar `applyFilter('data', ...)` / `applyFilter('instrutor', ...)` / `applyFilter('tipo', ...)` à montagem do payload `q`
- Atualizar `renderizar()` para usar `matchesSearch` + filtros server-side (ou híbrido)

### D. Renovar token automaticamente
- O JWT expira (olhar `exp` no payload)
- Worker pode ter rota `/refresh` que usa `refresh_token` se disponível
- Ou rodar script agendado pra pegar token novo via OAuth client credentials

---

## 📁 Arquivos modificados hoje
- `F:\projeto\gestor 2.0\index.html` — filtros no header das páginas DB1/DB2 + variáveis de estado
- `F:\projeto\gestor 2.0\PROGRESSO.md` — este arquivo

---

## 🔑 Credenciais / Tokens (NÃO COMMITAR)
- **Login painel**: `wagnerccorreia` / `mamute65` (apenas para obter token no browser)
- **JWT atual**: expira em ~8h (ver `exp` no payload) — pegar novo no Network tab logado no painel
- **API Base**: `https://api.navedoconhecimento.rio`

---

## 💡 Ideias futuras
- Cache local (IndexedDB) para funcionar offline
- Sync bidirecional: editar status no gestor → PATCH na API
- Dashboard com métricas consolidadas (total indicadores, por unidade, por tipo)
- Exportar CSV/PDF direto do gestor

---

**Continuar daqui:** rodar o teste no console → se der JSON válido → criar Worker → integrar no `index.html`.