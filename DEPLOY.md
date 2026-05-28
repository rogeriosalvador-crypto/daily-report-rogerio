# Deploy Daily Report — Render.com

## Visão Geral
Este app Flask exibe o Daily Report da carteira Rogério Salvador.
A URL pública será acessível de qualquer dispositivo (celular, tablet, computador).

**URL esperada:** `https://daily-report-rogerio.onrender.com`

---

## Opção A: Deploy via GitHub (Recomendado)

### Passo 1 — Criar conta no Render
Acesse: https://render.com  
Clique em **Get Started for Free** e crie a conta (pode usar a conta do GitHub).

### Passo 2 — Enviar código para o GitHub
```bash
# No terminal, dentro da pasta daily-report-public:
git init
git add .
git commit -m "Daily Report app - deploy público"

# Criar repositório no GitHub (https://github.com/new)
# Depois conectar:
git remote add origin https://github.com/SEU_USUARIO/daily-report-rogerio.git
git push -u origin main
```

### Passo 3 — Criar Web Service no Render
1. Acesse o dashboard: https://dashboard.render.com
2. Clique em **New +** → **Web Service**
3. Conecte sua conta do GitHub quando solicitado
4. Selecione o repositório `daily-report-rogerio`
5. Configure:
   - **Name:** `daily-report-rogerio`
   - **Environment:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Instance Type:** `Free`
6. Clique em **Create Web Service**

### Passo 4 — Aguardar o deploy
- O Render fará o build automaticamente (1-3 minutos)
- A URL pública aparecerá no topo do dashboard
- Formato: `https://daily-report-rogerio.onrender.com`

---

## Opção B: Deploy via Upload Direto

### Passo 1 — Criar conta no Render
Acesse: https://render.com e crie a conta.

### Passo 2 — Upload do arquivo ZIP
1. Use o arquivo `daily-report-public.zip` (gerado junto com este projeto)
2. No dashboard Render, clique em **New +** → **Web Service**
3. Selecione a opção de upload de código (caso disponível no seu plano)
4. Faça upload do ZIP

> **Nota:** O Render free tier funciona melhor com GitHub. O upload direto pode requerer configuração adicional.

---

## Configurações do Servidor

| Parâmetro       | Valor                          |
|-----------------|-------------------------------|
| Runtime         | Python 3.11                   |
| Servidor WSGI   | Gunicorn                      |
| Porta           | Definida automaticamente pelo Render via `$PORT` |
| Build Command   | `pip install -r requirements.txt` |
| Start Command   | `gunicorn app:app`            |

---

## Atualizar o Relatório

Após o deploy, você pode atualizar os dados de duas formas:

### Via API (recomendado para automação)
```bash
curl -X POST https://daily-report-rogerio.onrender.com/atualizar \
  -H "Content-Type: application/json" \
  -d @data/relatorio.json
```

### Via código Python
```python
import requests, json

url = "https://daily-report-rogerio.onrender.com/atualizar"
with open("data/relatorio.json") as f:
    dados = json.load(f)

resp = requests.post(url, json=dados)
print(resp.json())
```

---

## Limitações do Free Tier

- O serviço entra em modo de **sleep** após 15 minutos sem uso
- O primeiro acesso após sleep demora ~30 segundos para "acordar"
- Para manter sempre ativo, considere usar um serviço de ping (ex: UptimeRobot)

---

## Estrutura do Projeto

```
daily-report-public/
├── app.py              # Aplicação Flask principal
├── wsgi.py             # Entry point WSGI para Gunicorn
├── requirements.txt    # Dependências Python
├── render.yaml         # Configuração de deploy Render
├── DEPLOY.md           # Este arquivo de instruções
├── data/
│   └── relatorio.json  # Dados do relatório (atualizável via API)
└── templates/
    └── index.html      # Dashboard HTML/CSS/JS
```

---

## Suporte

Em caso de problemas:
- Verifique os logs no dashboard Render → seu serviço → **Logs**
- Teste localmente: `gunicorn app:app` (porta padrão 8000)
- Documentação Render: https://render.com/docs
