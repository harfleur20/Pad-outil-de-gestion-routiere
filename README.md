# Projet PAD - Sprint 0

Squelette initialise: Electron + React + TypeScript + theme PAD.

## Demarrage
1. Installer les dependances
   npm install
2. Lancer en mode dev
   npm run dev

## Structure
- app/electron/main.cjs: processus principal Electron
- app/electron/preload.cjs: pont securise
- app/web: frontend Vite React
- app/web/src/styles/theme.css: couleurs PAD

## Prochaine etape
- Connecter SQLite avec le schema `docs/schema-sqlite-v1.sql`
- Construire l'import des voies depuis Excel
