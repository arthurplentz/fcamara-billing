# Controle de Faturamento — Grupo Fcamara

App React (Vite) para acompanhar o pipeline de faturamento dos analistas do Grupo Fcamara.

## Rodar localmente

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

## Deploy

O deploy é automático via GitHub Actions (`.github/workflows/deploy.yml`) a cada push na branch `main`, publicando em GitHub Pages.

## Login (dados de exemplo)

| Usuário       | Senha          | Perfil |
|---------------|----------------|--------|
| Daniela       | `daniela`      | Admin  |
| Fernanda      | `fernanda`     | Analista |
| Layza Arruda  | `layza arruda` | Analista |

> Os dados ficam no `localStorage` do navegador (não há backend).
