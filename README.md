# ISGA Timetable

Application web pour consulter les emplois du temps des filières ISGA, synchronisés depuis le classeur Excel SharePoint.

## Fonctionnement

- L’API `/api/timetable` lit le fichier Excel SharePoint (lien de partage ou Microsoft Graph).
- Le résultat est mis en cache **30 secondes** pour une synchro quasi temps réel.
- L’interface rafraîchit automatiquement toutes les 30 secondes et permet de filtrer par filière / semaine.
- Sans accès SharePoint configuré, l’app affiche des **données d’exemple** pour prévisualiser l’UI.

## Configuration SharePoint

Deux options (recommandé : les deux) :

1. **Lien de partage** `SHAREPOINT_SHARE_URL`  
   Dans SharePoint/OneDrive : *Partager → Toute personne disposant du lien peut afficher*.

2. **Microsoft Graph (app Azure AD)**  
   - `MS_TENANT_ID`
   - `MS_CLIENT_ID`
   - `MS_CLIENT_SECRET`  
   Permissions applicatives : `Files.Read.All` ou `Sites.Read.All` (admin consent).

Le lien source par défaut pointe vers le classeur ISGA fourni.

## Développement

```bash
npm install
npm run dev
```

Ouvrir [http://127.0.0.1:3000](http://127.0.0.1:3000).

## Scripts

- `npm run dev` — serveur de développement
- `npm run build` — build production
- `npm start` — serveur production
