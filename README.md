# SharePoint indexing with Azure AI Search

This project provisions Azure AI Search, a Microsoft Foundry resource and a
`text-embedding-3-large` model deployment. It can then connect Search to an
existing SharePoint app registration and create a SharePoint indexing pipeline.

`azd provision` runs the infrastructure and the enabled setup steps. There is
no application to deploy, so `azd deploy` is not needed.

> **Before using this with real data:** the SharePoint indexer and Search
> serverless tier are previews. This project does **not** copy or enforce
> SharePoint document permissions: anyone allowed to query the index can
> retrieve its indexed content. Review access, costs, networking and preview
> limitations before using it, especially in production.

## 1. Prerequisites

Install the following:

- [Azure CLI (`az`)](https://learn.microsoft.com/cli/azure/install-azure-cli)
  and [Azure Developer CLI (`azd`)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd).
  Use a recent `azd` version that supports `azd env set --file` and
  `azd hooks run`.
- [Python 3.11 or later](https://www.python.org/downloads/) and
  [uv](https://docs.astral.sh/uv/getting-started/installation/). The hooks use
  `uv run` to run the Python scripts.
- PowerShell 7 (`pwsh`) on Windows, or a POSIX shell on Linux/macOS. The
  examples below use PowerShell and run from the repository root.

The deploying identity needs permission to create subscription deployments and
resource groups, create resources, and write role assignments. One practical
combination is **Contributor** and **Role Based Access Control Administrator**
at the subscription scope. Azure RBAC and Microsoft Entra permissions are
separate; the SharePoint app and optional federation steps have additional
requirements below.

Register the `Microsoft.Search` and `Microsoft.CognitiveServices` resource
providers if necessary. Check that the selected regions support the Search
SKU, embedding model/version and capacity; image modes also need Vision/OCR
availability. Azure charges apply.

Complete the [SharePoint indexer preview registration](https://aka.ms/azure-cognitive-search/indexer-preview).
It is separate from the Search serverless feature registration.

## 2. Prepare the SharePoint app registration

Reuse an existing app if it already has the required access, or create one in
the SharePoint tenant. The automated federation workflow requires the app and
Search managed identity to be in the **same tenant**.

1. In the [Microsoft Entra admin center](https://entra.microsoft.com/), select
   the tenant containing the SharePoint site. Go to **Entra ID > App
   registrations > New registration**.
2. Enter a name, select **Single tenant only**, and register the app. A redirect
   URI is not needed for this app-only flow.
3. On **Overview**, copy the **Application (client) ID** and **Directory
   (tenant) ID**. These become `SHAREPOINT_APPLICATION_ID` and
   `SHAREPOINT_TENANT_ID`.
4. Open **API permissions > Add a permission > Microsoft Graph > Application
   permissions**. Add **`Files.Read.All`** and **`Sites.Read.All`**.
5. Have an authorized administrator select **Grant admin consent for
   \<tenant\>**, and verify that consent is granted for both permissions.

These are application permissions on the **ingestion app** and grant broad
tenant-wide read access; review them with your administrator. They are not the
Graph permissions used by the person or client running the post-provision
scripts. This project indexes document libraries and does not configure ACL
ingestion. See Microsoft's
[SharePoint indexer setup](https://learn.microsoft.com/azure/search/search-how-to-index-sharepoint-online)
for its current requirements and restrictions.

**Who can do this?** Microsoft's app-registration guide lists **Application
Developer** as a supported role for registration; tenant policy might also
allow ordinary users to register apps. Granting tenant-wide consent to these
Microsoft Graph *application* permissions is a separate, more privileged
operation: Microsoft lists **Privileged Role Administrator** or an appropriately
configured custom role. **Application Administrator** and **Cloud Application
Administrator** alone cannot grant these Graph application permissions.
See [register an app](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app)
and [grant admin consent](https://learn.microsoft.com/entra/identity/enterprise-apps/grant-admin-consent).

The app's federated credential can be created after Search exists, either by
the optional post-provision stage or by an authorized administrator. No client
secret is required for federation.

## 3. Prepare an environment file

Create a separate file for each environment from the example. For example,
create `config_files\.env.dev`, then edit it with your subscription, location,
SharePoint IDs and desired settings:

```powershell
Copy-Item .\config_files\.env.example .\config_files\.env.dev
```

Only run the copy command when the destination does not already contain
settings you want to keep. Repeat with `.env.qa` and `.env.prd` for other
environments.

These names are ignored by Git. Keep secrets out of Git and
use protected local or CI secret handling.

**The file is not loaded automatically by `azd provision`.** The `azd env
set --file` command in the deployment steps imports its values into the
selected azd environment. After editing the file, import it again to apply the
changes; values in it can overwrite values already set in that environment.
Alternatively, change individual settings with `azd env set`.

## 4. Environment variables

The example's active values are starting points, not production sizing
recommendations. Replace placeholders before enabling the relevant steps.

### Azure resources and pipeline

| Variable | Description |
| --- | --- |
| `AZURE_SUBSCRIPTION_ID` | Required subscription GUID where the resources are created. |
| `AZURE_RESOURCE_GROUP` | Resource group name. Blank uses `rg-<AZURE_ENV_NAME>`; use a different resource group for each environment. |
| `AZURE_LOCATION` | Required Azure region for Search, such as `polandcentral`. Foundry uses it unless `FOUNDRY_LOCATION` is set. |
| `DEPLOYMENT_TYPE` | Search pipeline mode: `no_embedding`, `image_embeddings`, `text_embeddings`, or `all_embeddings`. See the table below. |
| `AZURE_OPENAI_DEPLOYMENT` | Name of the `text-embedding-3-large` deployment; example: `text-embedding-3-large`. This is a deployment name, not a model selector. Bicep creates it in every mode and outputs its name. |
| `EMBEDDING_MODEL_VERSION` | Model version; example: `1`. It must be available in the chosen region and SKU. |
| `EMBEDDING_CAPACITY` | Positive integer for model quota capacity units; example: `20`. Check available quota; this is not a raw tokens-per-minute value. |
| `SEARCH_SKU` | `basic`, `standard`, `standard2`, `standard3`, or `serverless`. The example uses `basic`; availability and capacity vary by region and subscription. |
| `SEARCH_SEMANTIC_SEARCH` | `disabled`, `free`, or `standard`. The example uses `standard`; availability and limits depend on the service. |
| `SEARCH_INDEX_NAME` | Base name for Search objects; example: `sharepoint-rag`. The deployment type is appended, e.g. `sharepoint-rag-text-embeddings`. Use lowercase letters, digits and hyphens; avoid consecutive hyphens. |
| `INDEXER_SCHEDULE` | Schedule in `PT<n>H<n>M` form, such as `PT15M`, `PT1H`, or `PT24H`; supported range is 5 minutes to 24 hours. An explicitly blank value disables scheduled runs, not manual runs. |

| `DEPLOYMENT_TYPE` | Pipeline |
| --- | --- |
| `no_embedding` | Extracted text and metadata; no vectors or skillset. |
| `image_embeddings` | OCR-enriched text and image vectors using Vision. |
| `text_embeddings` | Text chunks and OpenAI text vectors; no OCR. |
| `all_embeddings` | OCR-enriched text, text vectors and image vectors. |

For `image_embeddings` and `all_embeddings`, Bicep provides a multi-service
Foundry account (`Microsoft.CognitiveServices/accounts`, kind `AIServices`,
SKU `S0`). This is the same account used for OpenAI, not an OpenAI-only
resource; no additional classic `CognitiveServices` account is needed.
Search's system-assigned identity receives **Cognitive Services User** for
OCR/Vision and **Cognitive Services OpenAI User** for text embeddings on that
account. `AZURE_AI_SERVICES_ENDPOINT` is populated automatically and used by
the skillset and image query vectorizer.

The image modes enable image extraction, OCR and image projections into a
1,024-dimension vector field; `all_embeddings` also includes 3,072-dimension
OpenAI text vectors. The skillset explicitly uses Search's system-assigned
identity for billing. Verify that `FOUNDRY_LOCATION` supports the
[Vision multimodal model](https://learn.microsoft.com/azure/search/vector-search-vectorizer-ai-services-vision).
Keep the API keys unset for this keyless deployment. The existing
`CONFIGURE_SEARCH_PIPELINE` flag still controls creation of the Search objects.

Changing the type creates separately named Search objects; it does not remove
or stop the previous pipeline. Existing schedules and charges can continue.

### SharePoint and post-provision settings

| Variable | Description |
| --- | --- |
| `SHAREPOINT_TENANT_ID` | Tenant GUID of the SharePoint site and ingestion app. Required for federation and Search setup. |
| `SHAREPOINT_SITE_URL` | Site URL, e.g. `https://contoso.sharepoint.com/sites/DocumentHub`. Required for Search setup. |
| `SHAREPOINT_APPLICATION_ID` | Ingestion app's **client/application ID** GUID. Required for federation and Search setup. |
| `SHAREPOINT_LIBRARY_URL` | Optional URL of a library within the site. Blank selects the site's default document library. |
| `SHAREPOINT_APPLICATION_SECRET` | Optional ingestion-app client secret. Blank uses Search's federated identity. If set, disable the federation stage and protect the secret. |
| `CONFIGURE_SHAREPOINT_FEDERATION` | `true` enables the federation stage; `false` skips it. |
| `CONFIGURE_SEARCH_PIPELINE` | `true` enables the Search pipeline stage; `false` skips it. |

The example sets **both flags to `false`** for an infrastructure-only first
run. If the flags are omitted entirely, the scripts default federation to
`false` and Search setup to `true`. Use only `true` or `false`. See the next
section before changing them.

### Optional infrastructure and naming overrides

Uncomment these in the environment file only when needed:

| Variable | Description |
| --- | --- |
| `FOUNDRY_LOCATION` | Foundry/model region; blank uses `AZURE_LOCATION`. Verify model and Vision availability. |
| `AZURE_TENANT_ID` | Azure tenant GUID override for authentication. The deployment also outputs its tenant ID; automated federation requires it to match `SHAREPOINT_TENANT_ID`. |
| `AZURE_PRINCIPAL_TYPE` | `User` (default) or `ServicePrincipal`; use the latter when the provisioning identity is a service principal. |
| `SEARCH_REPLICA_COUNT` | Dedicated Search replicas, default `1`; subject to SKU limits. Ignored for serverless. |
| `SEARCH_PARTITION_COUNT` | Dedicated Search partitions: `1`, `2`, `3`, `4`, `6`, or `12`, subject to SKU/region limits; default `1`. Ignored for serverless. |
| `EMBEDDING_SKU` | Model deployment SKU, default `Standard`; availability and quota depend on the model and region. |
| `SEARCH_DATASOURCE_NAME` | Optional data source base name. |
| `SEARCH_SKILLSET_NAME` | Optional skillset base name. |
| `SEARCH_INDEXER_NAME` | Optional indexer base name. Each override receives the deployment type suffix. When unset, names are derived from the index name. |

The following optional keys are for the standalone script against **existing**
resources with key authentication enabled. Leave them unset for this Bicep
deployment: its Search and Foundry resources disable local/key authentication.

| Variable | Connection authenticated by the key |
| --- | --- |
| `SEARCH_ADMIN_KEY` | Python script to Search; this must be an admin key, not a query key. |
| `AZURE_OPENAI_KEY` | Search to the OpenAI embedding skill and query vectorizer. |
| `AZURE_AI_SERVICES_KEY` | Search to the multi-service OCR/Vision skills and image vectorizer. |

### Values generated by azd and the scripts

Normally you do not need to supply these for a new deployment:

| Variable | Meaning |
| --- | --- |
| `AZURE_ENV_NAME` | Environment selected by `azd env new` or `azd env select`; used in resource tags and the default resource-group name. |
| `AZURE_PRINCIPAL_ID` | Object ID of the user or service principal running azd, used for the Search role assignment. Normally resolved by azd. |
| `AZURE_SEARCH_SERVICE_NAME` | Provisioned Search service name. |
| `SEARCH_ENDPOINT` | Provisioned Search service endpoint. |
| `SEARCH_IDENTITY_PRINCIPAL_ID` | Search system-assigned identity's **object/principal ID**; used as the federation subject. |
| `SEARCH_IDENTITY_CLIENT_ID` | Search identity's **application/client ID**, used in the SharePoint connection. After verifying the trust, either post-provision stage resolves and saves the current value, replacing a stale value if Search was recreated. |
| `AZURE_AI_ACCOUNT_NAME` | Provisioned Foundry account name. |
| `AZURE_OPENAI_ENDPOINT` | Provisioned Foundry endpoint used by the text embedding pipeline. |
| `AZURE_AI_SERVICES_ENDPOINT` | Provisioned multi-service endpoint used for Vision/OCR. |

`AZURE_RESOURCE_GROUP`, `AZURE_TENANT_ID`, `AZURE_OPENAI_DEPLOYMENT` and
`DEPLOYMENT_TYPE` are also written as deployment outputs.

## 5. Understand the provisioning stages

`azd provision` runs these in order. A failed stage stops the workflow; already
created resources or credentials are not rolled back.

| Stage | What it does | Required access / prerequisites |
| --- | --- | --- |
| Pre-provision | If `SEARCH_SKU=serverless`, registers `Microsoft.Search/Serverless`, waits for it, and refreshes the Search provider registration. Other SKUs skip this. | Subscription-level feature and provider registration permissions; subscription **Contributor** or **Owner** is a practical choice. Registration is subscription-wide, can require approval, and does not guarantee regional availability. |
| Infrastructure | Creates the resource group, Search with a system-assigned identity, Foundry resource, model deployment and Azure role assignments. | Resource deployment and role-assignment write permissions described in prerequisites. Search gets **Cognitive Services OpenAI User** and **Cognitive Services User** on Foundry; the azd principal gets **Search Service Contributor** on Search. |
| Federation | When `CONFIGURE_SHAREPOINT_FEDERATION=true`, reads the Search identity, adds, reuses or updates its trust on the existing ingestion app, and saves the Search identity client ID. | Search identity and app must exist in the same tenant; app credential-write and Graph identity-read access are required. No app secret can be configured. See permissions below. |
| Search pipeline | When `CONFIGURE_SEARCH_PIPELINE=true`, verifies the existing federation and creates or updates the data source, index, indexer and, for embedding modes, skillset. | **Search Service Contributor** on Search; configured SharePoint app and admin consent; Graph read access and a matching trust, unless using an app secret. Required site and model settings must be valid. |

### Federation and Graph permissions

For a signed-in user, a least-privileged supported option for updating
federation is a **non-admin member who owns this particular ingestion app**
and has default user permissions. App ownership permits more than just
federated-credential changes. Alternatively, a supported Entra role or a
custom role scoped to the app with the necessary app-credential permissions
can be used. The calling client also needs Microsoft Graph delegated
`Application.ReadWrite.All` consent for credential creation/updates and access to read
the Search service principal. Subscription Contributor alone is not enough.
For app-only automation, Graph `Application.ReadWrite.OwnedBy` with ownership
of the ingestion app is a documented option.

The Search-only stage does **not** create or update a credential. With federation, it
reads the Search service principal and lists the app's credentials before
creating Search objects. For delegated work-account access, Graph
`Application.Read.All` plus the appropriate Entra user permissions or app
ownership is a documented read-only option. Graph permissions and consent for
the **azd calling client** are separate from the SharePoint ingestion app's
`Files.Read.All` and `Sites.Read.All` permissions. Access may depend on tenant
policy. See Microsoft's [create credential](https://learn.microsoft.com/graph/api/federatedidentitycredential-post),
[update credential](https://learn.microsoft.com/graph/api/federatedidentitycredential-update),
[list credentials](https://learn.microsoft.com/graph/api/federatedidentitycredential-list)
and [read service principal](https://learn.microsoft.com/graph/api/serviceprincipal-get)
permission requirements.

The federation stage reuses an existing matching trust; otherwise it creates
one named after the Search service. If Search was recreated with the same name,
the stage updates the existing credential's subject to the new managed identity
instead of deleting and recreating the credential. This is allowed only when
its issuer and audience match and its description is exactly
`Search identity for azd environment <environment>.`, as set by this script.
Credentials with different names are left untouched; a same-name credential
that fails these checks requires app-owner/administrator review.

An in-place update works even at the app's 20-credential limit; creating an
additional credential at the limit still fails. After a successful update,
the stage saves the current Search client ID. To retry just the post-provision
steps after this failure, keep `CONFIGURE_SHAREPOINT_FEDERATION=true` and run:

```powershell
azd --environment qa hooks run postprovision
```

`CONFIGURE_SEARCH_PIPELINE` still independently controls Search object setup.
Graph verification checks the trust's configuration, not whether token exchange
has propagated or indexing will succeed. Microsoft's SharePoint guidance describes system-assigned federation,
while the general Entra managed-identity guide describes user-assigned
identities; verify actual indexing for this preview flow.

### Choose the flags

| Desired result | `CONFIGURE_SHAREPOINT_FEDERATION` | `CONFIGURE_SEARCH_PIPELINE` |
| --- | --- | --- |
| Infrastructure, federation and Search objects | `true` | `true` |
| Infrastructure and federation only | `true` | `false` |
| Infrastructure and Search objects using an existing trust | `false` | `true` |
| Infrastructure only; wait for an administrator | `false` | `false` |

With an app secret, use `false` for federation; Search setup can be `true`
without a federated trust or its Graph preflight. The ingestion app still needs
its SharePoint permissions and admin consent.

## 6. Deploy with azd

The following follows the order in `commands.txt`, using the Git-ignored
`.env.dev` filename described above:

```powershell
az login --use-device-code
azd auth login --use-device-code

azd env new dev
azd --environment dev env set --file .\config_files\.env.dev
azd --environment dev env get-values
azd --environment dev provision
```

Sign in to the correct tenant and choose the correct subscription. `az` and
`azd` have separate sign-ins; the provisioning hooks use the **azd** identity.
`env get-values` prints the environment, including any secrets: do not run it
in public logs or share its output. Check your settings before provisioning.

Repeat the environment creation, import and provision steps for `qa` and
`prd`, using their respective files and settings. Do not copy azd's generated
resource IDs or endpoints between environments. If an environment already
exists, skip `azd env new` and target it by name in the remaining commands.

### Provision now, wait for an administrator, then set up Search

Use this workflow when you can provision Azure resources but cannot add a
federated credential to the SharePoint app. The example uses `dev`; substitute
your environment and file name throughout.

1. **Provision infrastructure only.** In `config_files\.env.dev`, fill in the
   Azure and SharePoint settings and set both flags to `false`:

   ```dotenv
   CONFIGURE_SHAREPOINT_FEDERATION="false"
   CONFIGURE_SEARCH_PIPELINE="false"
   ```

   Sign in as shown above, then create the environment, import the file and
   provision:

   ```powershell
   azd env new dev
   azd --environment dev env set --file .\config_files\.env.dev
   azd --environment dev provision
   ```

   If `dev` already exists, skip `azd env new`. This creates the Azure
   resources and RBAC, but neither a federated credential nor Search pipeline
   objects. Keep `SHAREPOINT_APPLICATION_SECRET` unset for this federation
   workflow.

2. **Ask an authorized app owner or administrator to create the trust.** After
   provisioning completes, obtain the values below and give them to the
   administrator, together with the SharePoint ingestion app's client ID:

   ```powershell
   azd --environment dev env get-value AZURE_SEARCH_SERVICE_NAME
   azd --environment dev env get-value SEARCH_IDENTITY_PRINCIPAL_ID
   azd --environment dev env get-value AZURE_TENANT_ID
   azd --environment dev env get-value SHAREPOINT_APPLICATION_ID
   ```

   On that existing app registration, the administrator should create a
   federated credential with:

   | Field | Value |
   | --- | --- |
   | Name | Search service name is recommended; the verification also accepts a matching trust under another name. |
   | Issuer | `https://login.microsoftonline.com/<azure-tenant-id>/v2.0` |
   | Subject | `SEARCH_IDENTITY_PRINCIPAL_ID`, the Search identity's **object/principal ID** (not its client ID). |
   | Audience | `api://AzureADTokenExchange` |

   The app and Search identity must be in the same tenant. The administrator
   must also ensure the app has the SharePoint application permissions and
   admin consent described in section 2. The Entra portal's managed-identity
   picker may only offer user-assigned identities; an administrator can use
   Microsoft Graph to configure the system-assigned identity using these
   values. See [managed identity federation](https://learn.microsoft.com/entra/workload-id/workload-identity-federation-config-app-trust-managed-identity).

3. **After the administrator confirms the trust, enable Search setup.** The
   user running this step needs **Search Service Contributor** on the Search
   service and Graph read access to the Search identity and app credential, as
   described in [Federation and Graph permissions](#federation-and-graph-permissions).
   Confirm the SharePoint site, app ID, tenant ID, pipeline type and schedule
   are set in the selected azd environment. For example, use `PT24H` for a
   daily schedule, not `PT1D`.

   Change the flags in the azd environment and run the post-provision hook:

   ```powershell
   azd --environment dev env set CONFIGURE_SHAREPOINT_FEDERATION false
   azd --environment dev env set CONFIGURE_SEARCH_PIPELINE true
   azd --environment dev hooks run postprovision
   ```

   The hook checks the existing trust, resolves and saves the current Search
   identity client ID, and creates or updates the Search objects. It does not
   create the trust in this mode. If the trust is missing or Graph read access
   is denied, it stops before saving the ID or creating Search objects. No
   `azd deploy` or reprovisioning is needed.
   Check indexer execution as described below; trust verification alone does
   not prove that indexing succeeds.

The `azd env set` commands change the stored azd environment, not the source
file. If you later re-import `config_files\.env.dev`, update its flags first;
otherwise its old values can overwrite the changes.

For federation only, set the federation flag to `true` and the Search flag to
`false`, then run the same hook. For both stages, set both flags to `true`.

## 7. Check indexing

In the Azure portal, open the Search service and inspect **Indexers** and the
execution history for the indexer named from your index base and deployment
type. Creating Search objects does **not** prove that authentication or
indexing succeeded. A new indexer starts automatically; updating an existing
one requests an incremental run. Federation and role assignments can take
time to propagate. Review errors, fix the cause, and retry when ready.

Re-running setup creates or replaces the named definitions; it does not
automatically reset the indexer, reprocess unchanged files, or delete older
pipelines. Use a new index base name or intentionally rebuild if a fresh crawl
is needed.

## Existing resources and cleanup

For an existing Search service, `scripts\provision.py` can run separately.
Copy `config_files\.env.example` to a root `.env`, provide the existing Search
endpoint and the required SharePoint/AI settings, then use:

```powershell
uv run --env-file .env python .\scripts\provision.py --dry-run
uv run --env-file .env python .\scripts\provision.py --search-auth azd
uv run --env-file .env python .\scripts\provision.py --search-auth azd --status
```

The standalone script does not load an azd environment, create federation or
grant roles. `--search-auth azd` uses the azd sign-in for Python-to-Search;
without it or a Search admin key, the script expects the Python host's managed
identity. A populated key takes precedence. Service-to-service access still
requires the corresponding Search identity roles and SharePoint trust.

`azd --environment dev down --purge` deletes the environment's managed Azure
resources; it does not delete the external SharePoint app or its federated
credentials. This is destructive: check the environment and resource group
before running it. The analogous commands in `commands.txt` also delete QA and
production resources.
