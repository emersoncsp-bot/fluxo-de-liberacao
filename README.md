# Fluxo de Liberação de Produtos

Aplicação React para gestão do fluxo de liberação de produtos, com etapas por departamento.

---

## 🚀 Deploy no Vercel (recomendado)

### Opção 1 — Via GitHub (mais simples)

1. Crie um repositório no GitHub e envie esta pasta
2. Acesse [vercel.com](https://vercel.com) → "New Project"
3. Importe o repositório
4. Deixe as configurações padrão (Vite é detectado automaticamente)
5. Clique em **Deploy**

### Opção 2 — Via Vercel CLI

```bash
npm install -g vercel
cd fluxo-liberacao
npm install
vercel
```

---

## 💻 Rodar localmente

```bash
cd fluxo-liberacao
npm install
npm run dev
```

Acesse: http://localhost:5173

---

## 🔐 Contas de acesso (demo)

| E-mail | Senha | Departamento |
|--------|-------|-------------|
| qualidade.tecnica@empresa.com | 123456 | Controle da Qualidade [Área técnica] |
| qualidade.lib@empresa.com | 123456 | Controle da Qualidade [Liberação Intermediária] |
| planejamento.uap@empresa.com | 123456 | Planejamento UAP |
| planejamento.central@empresa.com | 123456 | Planejamento Central |
| admin@empresa.com | admin | Admin |

---

## 📋 Formato do arquivo CSV de importação

O arquivo deve ter as colunas abaixo (separadas por `;`, `,` ou `tab`):

```
ultima_ordem;lote;ippn;qualidade_qts;deposito_sap;motivo_bloqueio;motivo_bloqueio_texto;razao_bloq;num_cassete
```

Exemplo de linha:
```
ORD-001;LOTE-2024-01;IPN123;Aprovado;DEP-01;BLQ-001;Contaminação detectada;RAZAO-A;CASS-001
```

---

## 🔄 Etapas do fluxo

1. **Análise do Bloqueio** — Controle da Qualidade [Área técnica]
2. **Definição do Recurso** — Planejamento UAP
3. **Criação de Ordem** — Planejamento Central
4. **Instrução da Qualidade** — Controle da Qualidade [Área técnica]
5. **Liberação para Vínculo** — Controle da Qualidade [Liberação Intermediária]
6. **Vínculo dos Lotes** — Planejamento UAP
7. **Ativação de Flag** — Controle da Qualidade [Liberação Intermediária]
8. **Concluído**

---

## 🛠 Estrutura do projeto

```
fluxo-liberacao/
├── public/
│   └── favicon.svg
├── src/
│   ├── App.jsx       # Aplicação principal
│   ├── main.jsx      # Entry point
│   └── index.css     # Estilos globais
├── index.html
├── package.json
├── vite.config.js
└── vercel.json
```
