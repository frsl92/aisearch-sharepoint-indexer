# SharePoint indexing: one Python script

`provision.py` creates or updates a SharePoint indexing pipeline on an **existing
Azure AI Search service**, with four deployment types:

| Deployment type | What gets indexed | External AI services |
| --- | --- | --- |
| `no_embedding` (default) | One searchable document per SharePoint file, using extracted text and metadata; no skillset, OCR, image extraction or vectors | None |
| `image_embeddings` | Searchable OCR-enriched text chunks plus image records with Vision vectors | AI Services/Vision only |
| `text_embeddings` | Searchable text chunks with OpenAI vectors; no OCR or image extraction | OpenAI only |
| `all_embeddings` | OCR-enriched text chunks with OpenAI vectors plus image records with Vision vectors | OpenAI and AI Services/Vision |

Embedding modes use 2,000-character chunks with 500-character overlap. Text
embeddings use `text-embedding-3-large` (3,072 dimensions); image embeddings use
Vision model `2023-04-15` (1,024 dimensions). Only selected vector fields, skills
and query vectorizers are deployed. `no_embedding` requires only Search and
SharePoint settings; it creates a data source, index and indexer. Other types
also create a skillset. Scanned/image-only content needs OCR, so it is not
searchable as text in `no_embedding` or `text_embeddings`.

**No SharePoint ACLs are ingested or enforced.** The permission fields,
permission filters and ACL projections from the references have been removed.
Anyone authorized to query the index can retrieve its content regardless of
their SharePoint permissions. Only ingest content appropriate for that audience.

## Run

Requires Python 3.11+ and [uv](https://docs.astral.sh/uv/).

```powershell
uv sync
Copy-Item .\.env.example .\.env
# Edit .env: Search/SharePoint settings, optional keys and DEPLOYMENT_TYPE.
uv run --env-file .env python provision.py --dry-run
uv run --env-file .env python provision.py
uv run --env-file .env python provision.py --status
```

Start with `DEPLOYMENT_TYPE=no_embedding` in `.env`. Change it when ready, or
override it for a command using `--type`:

```powershell
uv run --env-file .env python provision.py --type no_embedding
uv run --env-file .env python provision.py --type image_embeddings
uv run --env-file .env python provision.py --type text_embeddings
uv run --env-file .env python provision.py --type all_embeddings
```

`--type` also applies to `--dry-run` and `--status`. It takes precedence over
`DEPLOYMENT_TYPE`.

`.env.example` contains your supplied resource URLs and IDs, but **no keys or
secrets**. `.env` is ignored by Git. `--dry-run` prints the four resource
definitions with secrets redacted and makes no Azure calls.

The script uses preview REST API `2026-08-01-preview` and bypasses the Import Data
wizard. It sends a normal data source body **without** an `identity` property
or `dataSource` wrapper, avoiding the wizard's `DataNoneIdentity` payload.

Package indexes are configured outside this repository. On Windows, uv automatically
reads user-level settings from `%APPDATA%\uv\uv.toml`; index entries there use
`[[index]]`, not `[[tool.uv.index]]`. Configure your required feed before running
uv. The generated `uv.lock` records registry and package download URLs but is
ignored by Git. Each machine resolves its own dependency versions.

## Authentication

Each connection chooses independently: **a populated key/secret wins; blank
means system-assigned managed identity**.

| Connection | Key/secret variable | When blank |
| --- | --- | --- |
| Python → Search | `SEARCH_ADMIN_KEY` (admin key, not query key) | System identity of the **Azure host running Python**, using `ManagedIdentityCredential` |
| Search → OpenAI skill and query vectorizer | `AZURE_OPENAI_KEY` | **Search service's** system identity |
| Search → OCR/Vision skills and image query vectorizer | `AZURE_AI_SERVICES_KEY` | **Search service's** system identity |
| Search → SharePoint | `SHAREPOINT_APPLICATION_SECRET` | Search service identity federated to `SHAREPOINT_APPLICATION_ID`, using `SEARCH_IDENTITY_CLIENT_ID` |

**A local laptop has no Azure system-assigned identity.** For local provisioning,
populate `SEARCH_ADMIN_KEY`; the other keys can remain blank so the service-side
calls still use Search's identity. The script does not silently fall back to
Azure CLI, a developer account, or a user-assigned identity.

For managed identity, administrators must configure:

- The Python host identity: **Search Service Contributor** on the Search service.
  Search must accept Entra/RBAC authentication.
- For text embeddings: the Search identity needs **Cognitive Services OpenAI
  User** on the resource hosting the embedding deployment.
- For image embeddings/OCR: the Search identity needs **Cognitive Services User**
  on the multi-service OCR/Vision resource. Neither AI role is needed for
  `no_embedding`.
- SharePoint: application permissions **Microsoft Graph `Files.Read.All` and
  `Sites.Read.All`**, with administrator consent for this non-ACL document-library
  scenario. Configure the ingestion app's federated credential to trust Search's
  system identity when not using an application secret.

`SEARCH_IDENTITY_CLIENT_ID` is Search's managed identity **application/client ID**,
not its principal/object ID. It is not the identity of the Python host.
SharePoint has no Search-style API key: its optional key-like credential is the
ingestion app's client secret. Secret-based and federated fields are never
combined in its connection string.

The script does not create identities, grant roles/consent, deploy models or
change networking.

## Environment settings

All settings are listed in `.env.example`.

- `SEARCH_ENDPOINT`: the Search service HTTPS endpoint.
- `DEPLOYMENT_TYPE`: one of the four types above; defaults to `no_embedding`.
- `SEARCH_INDEX_NAME`: **base name**, defaults to `sharepoint-rag`. The selected
  type is appended with underscores changed to dashes, e.g.
  `sharepoint-rag-no-embedding`. Other resource names default to that index name
  plus `-datasource`, `-skillset` and `-indexer`. Individual resource base-name
  overrides are available; the type suffix is appended to those too.
- `SHAREPOINT_SITE_URL`, `SHAREPOINT_APPLICATION_ID`,
  `SHAREPOINT_TENANT_ID`: the SharePoint site and ingestion app identifiers.
- `SEARCH_IDENTITY_CLIENT_ID`: required only when the SharePoint application
  secret is blank.
- `SHAREPOINT_LIBRARY_URL`: optional specific library URL; blank uses the site's
  default library.
- `AZURE_OPENAI_ENDPOINT`: resource root URL, **not** a project URL or `/openai/v1`
  path. The script normalizes it to exactly one trailing `/` for the embedding
  skill and query vectorizer. `AZURE_OPENAI_DEPLOYMENT` must name an existing
  **text-embedding-3-large** deployment. Both are required only for
  `text_embeddings` and `all_embeddings`.
- `AZURE_AI_SERVICES_ENDPOINT`: multi-service resource root used for skillset
  OCR/Vision billing and the image query vectorizer. Required only for
  `image_embeddings` and `all_embeddings`. Unused AI service settings are not
  validated or included in the deployed JSON.
- `INDEXER_SCHEDULE`: defaults to `PT1H` (hourly); blank disables scheduled runs.
  Supported intervals are 5 minutes to 24 hours, for example `PT15M` or `PT24H`.

## Behavior and prerequisites

The three or four REST PUTs run in dependency order. They **create or replace the
named definitions**; they are not a transactional deployment. Earlier successful
PUTs remain in place if a later request fails.

**You can rerun the script.** A new indexer starts automatically; when the indexer
PUT reports an update, the script requests an incremental `/run`. Stable source
keys prevent duplicate full-document records in `no_embedding`. No automatic
reset, full re-embedding or deletion is performed. A run with unchanged source
content can legitimately process zero files. If an indexer is already running,
the service may return 409: inspect `--status` and retry after it finishes.

**Switching types creates a separate pipeline**, so each type has its own schema,
full initial crawl and incremental checkpoint. This avoids trying to remove
fields or mixing full-document and chunk records in an existing index. For
example, `sharepoint-rag-no-embedding` and `sharepoint-rag-text-embeddings` can
coexist; running the same type again reuses its resources. Queries and `--status`
must target the selected type's index/indexer.

Previously deployed pipelines **remain active**, including their schedules and
any embedding charges. Disable unused indexers separately if you only want one
type running. Older resources created before type suffixes were introduced are
not modified or deleted.

For a fresh full crawl within the same type after changing skills, chunking or
scope, use a new `SEARCH_INDEX_NAME` base or intentionally reset/rebuild
separately. Updating a definition does not backfill unchanged source documents.

`--status` prints the service's execution results, errors and warnings. Successful
resource creation **does not mean indexing succeeded**. API failures exit nonzero;
review the failing response and indexer status. A client timeout can occur after
the service accepted a request, so inspect existing resources before retrying.

Before running, verify:

- Search is Basic or higher, SharePoint preview registration is complete, and
  the connector is available for your service.
- For the selected embedding type, the model/resource exists and its regional
  availability, capacity and quotas meet your needs. Enable semantic ranker if
  you intend to issue semantic queries.
- Firewalls and tenant policies permit the connections. The native SharePoint
  connector documents restrictions on private endpoints and Conditional Access.
- For image types, your OCR/Vision resource supports the skills. For key-based skillset billing,
  review the documented same-region requirement; identity-based billing has
  different regional requirements.
- Image extraction, OCR, embeddings and Search usage can incur charges.

Scheduled runs pick up new/changed content and source deletions. The computer
running this script does not need to stay on. Monitor indexing failures and
coverage in Search; folder renames, extraction limits, encryption and unsupported
files still require operational attention.

## Development

```powershell
uv sync --group dev
uv run pytest -q
uv run ruff check .
uv run mypy provision.py
```

Tests use mocked HTTP and identities; they do not provision Azure resources.

References: [SharePoint connector](https://learn.microsoft.com/en-us/azure/search/search-howto-index-sharepoint-online),
[skillset billing/authentication](https://learn.microsoft.com/en-us/azure/search/cognitive-search-attach-cognitive-services),
[Vision skill](https://learn.microsoft.com/en-us/azure/search/cognitive-search-skill-vision-vectorize),
[Vision vectorizer](https://learn.microsoft.com/en-us/azure/search/vector-search-vectorizer-ai-services-vision).
