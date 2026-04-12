# Open-XL

A professional Excel Add-in built with Office.js and modern JavaScript tooling.

## Stack

| Tool | Purpose |
|------|---------|
| Office.js (CDN) | Excel API |
| Webpack 5 | Bundler |
| Babel | ES2021+ transpilation |
| ESLint + Prettier | Code quality |
| Jest | Unit testing |
| GitHub Actions | CI/CD |

## Project Structure

```
Open-XL/
├── public/
│   └── manifest.xml          # Office Add-in manifest
├── src/
│   ├── taskpane/             # Task pane UI (HTML + JS + CSS)
│   ├── commands/             # Ribbon command handlers
│   ├── functions/            # Custom Excel functions
│   ├── services/             # Excel API service layer
│   ├── store/                # App state management
│   ├── utils/                # Shared utilities
│   ├── dialogs/              # Dialog windows
│   └── assets/               # Icons and static assets
├── tests/
│   ├── unit/                 # Jest unit tests
│   ├── integration/          # Integration tests
│   └── e2e/                  # End-to-end tests
├── .github/workflows/        # CI/CD pipelines
└── docs/                     # Documentation
```

## Getting Started

```bash
npm install
npm start        # Dev server at https://localhost:3000
```

Sideload `public/manifest.xml` in Excel to test the add-in.

## Scripts

| Command | Description |
|---------|-------------|
| `npm start` | Start dev server (HTTPS) |
| `npm run build` | Production build |
| `npm test` | Run tests |
| `npm run lint` | Lint source files |
| `npm run format` | Format with Prettier |

## License

MIT
