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
| `COMPLYADVANTAGE_ACCESS_KEY` | `api/complyadvantage-cases` | Access key du flux OAuth2 client credentials de l'API ComplyAdvantage Mesh. |
| `COMPLYADVANTAGE_ACCESS_SECRET` | `api/complyadvantage-cases` | Access secret associé, jamais exposé au front-end. |
| `COMPLYADVANTAGE_BASE_URL` | `api/complyadvantage-cases` | *(Optionnel)* URL de base de l'API Mesh, si elle diffère de la valeur par défaut codée dans la Function. Permet une correction sans redéploiement. |

### Développement local

Pour tester une Function en local (Azure Functions Core Tools / SWA CLI), créer un fichier
`api/local.settings.json` (déjà couvert par `.gitignore` — **ne pas le committer**) :

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "OPENAI_API_KEY": "...",
    "COMPLYADVANTAGE_ACCESS_KEY": "...",
    "COMPLYADVANTAGE_ACCESS_SECRET": "..."
  }
}
```

## Module ComplyAdvantage — point de vigilance avant mise en prod

Le module de suivi des cas ComplyAdvantage (`complyadvantage.html` + `api/complyadvantage-cases`)
a été construit à partir de la documentation publique de l'API Mesh
(https://docs.mesh.complyadvantage.com/reference/createtokenv3), qui ne garantit pas certains
détails. **Avant la mise en production**, faire un premier appel de test avec les vraies clés API
(ou via la collection Postman fournie par ComplyAdvantage, si disponible) pour confirmer :

1. **L'URL de base exacte de l'API Mesh** et le chemin de l'endpoint de token (`/v3/token` supposé).
   Ajustable sans redéploiement via `COMPLYADVANTAGE_BASE_URL`, ou dans le code
   (`MESH_BASE_URL_DEFAULT` / `TOKEN_PATH` en tête de `api/complyadvantage-cases/index.js`) si le
   chemin lui-même diffère.
2. **Le format exact du corps de la requête de token** (noms de champs attendus par `/v3/token`) —
   à ajuster dans `getMeshToken()`.
3. **Le nom exact du champ "assigné à"** dans la réponse réelle de `GET /v2/cases` — non documenté
   publiquement. Le mapping actuel (`extractAssignee()`) teste plusieurs noms plausibles et affiche
   "Non assigné" si aucun ne correspond, sans jamais planter ; à corriger une fois le nom confirmé.

Ces trois points sont aussi signalés par des commentaires `À CONFIRMER` directement dans
`api/complyadvantage-cases/index.js`.
