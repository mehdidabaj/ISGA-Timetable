# ISGA Timetable

Application web pour consulter les emplois du temps des filières ISGA.

## Sans aucune clé (recommandé)

1. Ouvrez le classeur SharePoint ISGA dans le navigateur.
2. Téléchargez une copie (`.xlsx`).
3. Dans l’app, cliquez **Importer le Excel**.

L’app parse les feuilles (filières) et affiche les séances. Re-importez le fichier quand le planning change.

## Synchro automatique (optionnel)

Aucune clé Azure / Client Secret n’est nécessaire.

Demandez au propriétaire du fichier un lien de partage **« Toute personne disposant du lien peut afficher »**, puis renseignez `SHAREPOINT_SHARE_URL`. L’API `/api/timetable` le relira alors toutes les ~30 secondes.

Les variables `MS_TENANT_ID` / `MS_CLIENT_ID` / `MS_CLIENT_SECRET` sont **optionnelles** (Microsoft Graph) et inutiles dans la plupart des cas.

## Développement

```bash
npm install
npm run dev
```

Ouvrir [http://127.0.0.1:3000](http://127.0.0.1:3000).
