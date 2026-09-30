# Compliance Portal — Eiffel Investment Group

Portail statique (Azure Static Web App) avec des Azure Functions (`api/`) qui font le pont vers
Microsoft Graph / SharePoint et vers des API tierces, en gardant les secrets côté serveur.

## Configuration — variables d'environnement des Azure Functions

Ces variables se définissent dans les **Application Settings** de l'Azure Static Web App / Function
App, depuis le Portail Azure. Elles ne doivent **jamais** être commitées dans le dépôt (ni dans le
front-end, ni dans `api/local.settings.json`).

| Variable | Utilisée par | Rôle |
|---|---|---|
| `OPENAI_API_KEY` | `api/analyze` | Clé de l'API OpenAI (Doc Review). |

### Développement local

Pour tester une Function en local (Azure Functions Core Tools / SWA CLI), créer un fichier
`api/local.settings.json` (déjà couvert par `.gitignore` — **ne pas le committer**) :

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "OPENAI_API_KEY": "..."
  }
}
```
