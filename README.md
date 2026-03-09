# EarlySign · Dashboard DR

Dashboard de suivi chantier **SeveUp / EarlySign** — généré depuis le template Excel `EarlySign_Template_Demo.xlsx`.

## Structure

```
CL_Dashboard_RAF/
├── index.html        ← Dashboard principal (DR)
├── netlify.toml      ← Config Netlify
└── README.md
```

## Déploiement Netlify

1. Push ce dossier sur un repo GitHub
2. Sur app.netlify.com → "Add new site" → "Import an existing project"
3. Sélectionner le repo
4. Build settings :
   - **Base directory** : `CL_Dashboard_RAF` (ou laisser vide si à la racine)
   - **Publish directory** : `.`
   - **Build command** : *(laisser vide)*
5. Deploy site ✅

## Données

Les données sont actuellement hardcodées dans `index.html` (section `DATA`).
Prochaine étape : lecture dynamique depuis un fichier `data.json` généré par le launcher Python SeveUp.
