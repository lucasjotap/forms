# vinext-starter

A clean full-stack starter running on [vinext](https://github.com/cloudflare/vinext), with optional Cloudflare D1 and Drizzle support.

## Prerequisites

- Node.js `>=22.13.0`
- Portable: Windows, macOS, or Linux; no Bash required
- Managed Linux: managed Linux runtime with Bash, `flock`, `curl`, `sha256sum`, and GNU `timeout`
- Git and `tar` (including Windows `tar.exe`) are required for publishing; portable publication does not require Bash

## Sites Lifecycle

The Sites initializer copies the shared starter and selects managed-linux only when `SITES_MANAGED_LINUX_CONTAINER=1`; otherwise it selects portable. It saves the selection only in ignored `.sites-runtime/execution-profile.json`. Both profiles copy/configure first, then use the plugin's separate `install-dependencies.mjs` step to measure installation independently. Edit source under `app/` and follow the Sites skill for installation, preview, builds, and publishing.

Run `node <plugin-root>/scripts/configure-execution-profile.mjs` only when the profile is unknown for the current checkout and environment. Profile changes do not alter tracked source or require reinstalling otherwise-valid dependencies; restart an existing preview to use the new selection. Do not commit or upload `.sites-runtime/`.

This starter does not use `wrangler.jsonc`.

Publish archives outside the source checkout or under its ignored `outputs/` directory. Local `.env*`, `.dev.vars*`, TypeScript build state, and tool runtime directories are ignored; keep the hosting manifest, migrations, and authored scripts in source control.

`install:ci` runs `npm ci` once against the shared lockfile, disables parent-workspace discovery, and includes required dev/optional dependencies despite production/omit settings. Sharp defaults to prebuilt binaries unless explicitly configured otherwise. Do not overlap installers.

- **Portable:** Preserve host HOME, npm cache, registry, proxy, temporary paths, retry/concurrency settings, and lifecycle-script policy. Use `--prefer-offline --no-audit --no-fund`.
- **Managed Linux:** Use the existing project-local HOME/cache/tmp setup and Linux install lock, tarball preflight, and timeout. Restore the image-seeded npm cache only when its lockfile hash matches; retain network fallback. Builds keep their existing timeout. These helpers are not invoked by the portable profile.

`scripts/sites-env.mjs` preserves the caller's HOME, npm cache, proxy, XDG, and temporary-directory configuration while defaulting Wrangler and Miniflare state to the checkout. If npm reports an unwritable cache, select a writable path with `npm_config_cache` for that install. The `dev` and `start` scripts also keep Wrangler logs inside the checkout. Generated `.sites-runtime/` and `.wrangler/` directories are disposable and ignored by Git.

On portable, `npm run dev` uses `vinext dev` with HMR, starting at port 5173. Vinext records the running server in ignored `.vinext/` state, rejects an ordinary duplicate launch, and recovers stale state after a stopped process; exactly simultaneous starts can race. Pass `--port <port>` or `--hostname <host>` after `npm run dev --` when needed; keep portable previews on loopback.

For managed Linux browser QA, follow the Sites skill's managed preview reference. The daemon runs the project's Vite dev script with `--host 0.0.0.0 --port 4173 --strictPort`; use the CLI's returned URL. The ignored local profile survives the daemon's cleared environment.

The portable profile simulates ChatGPT sign-in only for loopback development requests. Visit `/signin-with-chatgpt?return_to=/` to sign in as `local_seedy` (`seedy@sites.test`, display name `Seedy`) and `/signout-with-chatgpt?return_to=/` to sign out. The development cookie preserves that identity across server restarts. Mock auth is disabled in the managed-linux profile and is not included in production builds; hosted authentication remains dispatch-owned.

The Worker uses `vinext/server/fetch-handler`, including Vinext's config-aware image handling. After building, `npm start` runs that Worker locally through Wrangler on `127.0.0.1`, sharing `.wrangler/state` with dev preview and local D1 migrations; it does not deploy the site or simulate sign-in. Use the URL printed by the server. Pass `npm start -- --port <port>` to select a different built-preview port.

Local previews use Miniflare's placeholder `Request.cf` metadata without a network lookup. Set `CLOUDFLARE_CF_FETCH_ENABLED=true` to opt into fetching preview metadata; this setting does not change hosted request metadata.

Local tool usage metrics are disabled by default. Set `WRANGLER_SEND_METRICS=true` to opt in.

## Included Shape

- edit site code under `app/`
- `app/chatgpt-auth.ts` provides optional dispatch-owned ChatGPT sign-in helpers
- `.openai/hosting.json` declares optional Sites D1 and R2 bindings
- `vite.config.ts` simulates declared bindings for local development
- `worker/index.ts` is the production Worker entrypoint and preserves request-scoped connector bindings
- `scripts/vite/` contains the maintained Vite integrations and the vendored Sites plugin license
- `scripts/connector-preview/` contains the local connector preview Worker and session helpers
- the remaining `scripts/` files select execution profiles and run installation, development, and builds; the plugin's helpers orchestrate the Sites lifecycle
- `dist/` contains generated deployment output; `.vinext/`, `.wrangler/`, and `.sites-runtime/` contain ignored local tool state
- `db/index.ts` reads the D1 binding from the Cloudflare Worker environment
- `db/schema.ts` starts intentionally empty
- `@cloudflare/workers-types` provides Worker types; `cloudflare-env.d.ts` declares optional `DB`/`BUCKET` bindings—update these declarations if binding names change
- `examples/d1/` contains an optional D1 example surface
- `drizzle.config.ts` supports local migration generation when needed

## Workshop collection flow

The public form at `/avaliacao.html` collects opt-in diagnostic answers and shows
the participant up to three suggestions. Sending answers to the organizer still
requires the consent checkbox and goes to `/api/responses` (D1), exactly as before.

### Downloadable control model (generated in the browser)

After the suggestions appear, the participant can click **Baixar meu modelo de
controle (.xlsx)**. The file is built entirely in the browser; the download makes
no API call and sends no answers anywhere. The only network requests are two
same-origin static GETs, loaded lazily on the first click:

- `public/vendor/exceljs.min.js` — ExcelJS 4.4.0 browser build, MIT license
  (`public/vendor/exceljs.LICENSE.txt`), ~948 KB (~258 KB gzip). It is
  self-hosted, not loaded from a CDN. It is pinned as a devDependency; refresh the
  copy with `npm run vendor:exceljs` after changing the version.
- `public/planilha-modelo.js` — the workbook definition (`PlanilhaModelo.buildWorkbook`)
  and the CSV fallback (`PlanilhaModelo.buildCsv`).

The workbook follows `../PLANILHA_UNIVERSAL_CONTEXT.md` and
`../../08_diagnostico_e_devolutiva/LEIA_PRIMEIRO.md`. It adapts to the form's
"Usar em qual atividade?" answer:

| Activity | Tabs (in order; grey = complementary) |
|---|---|
| Workshop de Educação Financeira e Planejamento | LEIA PRIMEIRO, FLUXO DE CAIXA (passo 1), CONTAS E COMPROMISSOS (passo 2), ORÇAMENTO DO MÊS (passo 3), PLANO DE MELHORIA (passo 4), PAINEL, then grey CLIENTES E PEDIDOS, PRECIFICAÇÃO, ROTINA SEMANAL |
| Projeto 1 · Gestão de pequenos negócios | LEIA PRIMEIRO, PAINEL, FLUXO DE CAIXA, CONTAS E COMPROMISSOS, PRECIFICAÇÃO, PLANO DE MELHORIA, then grey CLIENTES E PEDIDOS, ROTINA SEMANAL |
| Projeto 2 · Organização e desenvolvimento | LEIA PRIMEIRO, PAINEL, ROTINA SEMANAL, CLIENTES E PEDIDOS, PLANO DE MELHORIA, then grey FLUXO DE CAIXA, CONTAS E COMPROMISSOS, PRECIFICAÇÃO |
| anything else | the universal 8 tabs (LEIA PRIMEIRO, PAINEL, FLUXO DE CAIXA, CONTAS E COMPROMISSOS, CLIENTES E PEDIDOS, PRECIFICAÇÃO, ROTINA SEMANAL, PLANO DE MELHORIA) |

The workshop version follows the participant workbook (caderno) page by page: LEIA PRIMEIRO
maps each workshop step to a tab, and the examples reuse Casos A/B/C. The tool answer also
changes the file: "papel" adds print guidance and taller rows for handwriting, and
"celular" adds tips for the mobile apps.

Every version separates input cells (light blue) from formula cells (grey-green). It has list,
date and value validation, conditional formatting, frozen headers, filters, tab colors, A4 print
areas (header + 25 rows, one page per tab, PAINEL two pages), pt-BR currency/date formats and
example rows labelled "EXEMPLO — APAGUE ANTES DE USAR". PAINEL leaves those example rows out
of its totals. The only form data written into the file is the business name, the activity,
the suggested focus, and each diagnosis suggestion (title, action and how to check). The
suggestions are pre-filled as the person's own rows in PLANO DE MELHORIA.
Formulas use only functions available in both Excel and Google Sheets (`SUMIFS`,
`COUNTIFS`, `IFERROR`, `INDEX`/`MATCH`, `REPT`, …). Empty rows return blanks or "—".
ExcelJS cannot create native charts, so PAINEL uses text bars (`REPT("█", …)`).

A link below the button offers a simple fallback: a UTF-8 (with BOM) CSV that uses `;`
as the separator and quotes every field. It is also generated client-side.

To change the spreadsheet, edit `public/planilha-modelo.js`; no build step is needed for it.

### Testing the spreadsheet

From the project root (`09_site_coleta/`), with `npm run dev -- --port 5185` running in `source/`:

```sh
python qa_xlsx.py                    # Playwright: fill form, download .xlsx + .csv, validate both
python qa_xlsx.py --file model.xlsx --csv model.csv   # validate existing files only
python qa_local.py                   # consent, D1 write, admin dashboard
```

`qa_xlsx.py` needs `playwright`, `openpyxl` and, optionally, `formulas`. Set `QA_BASE_URL` to test
another server, e.g. the built Worker started by `npm start -- --port 8799`. The script records every
network request made during the downloads and fails if any request is not a GET, goes to `/api/`, or
contains the business name or answers. It checks the sheets, headers, formulas, validations, freeze
panes, filters and CSV column counts. If LibreOffice (`soffice`) is installed, it also recalculates
the file and looks for error values. With the `formulas` package, it recalculates the file and a filled
test scenario. On Windows with Excel installed, it also opens each version in a private Excel instance
(never the user's open window). There it checks for a repair prompt, error cells and clipped text, and
counts printed pages; set `QA_EXCEL=0` to skip this. Screenshots go to `../qa_output/screens/`.
Google Sheets and phones are not tested automatically, so open the file there before a workshop.

The protected `/admin` page shows the individual records and aggregate, name-free
bars by question so the facilitator can adapt examples and teaching order. The JSON
export remains protected by the same administrator check. No participant data from
the database is ever placed in a downloaded spreadsheet.

## Workspace Auth Headers

Signed-in visitors receive both `oai-authenticated-user-id` and `oai-authenticated-user-email`. Private Sites require every visitor to sign in; public Sites may also have anonymous visitors, for whom neither header is present.

The user ID is stable for the same user on the same Site and different across Sites. Use it as the durable user key; use email and name for display or contact purposes.

SIWC-authenticated workspace sites may also receive `oai-authenticated-user-full-name` when the user's SIWC profile has a non-empty `name` claim. The full-name value is percent-encoded UTF-8 and is accompanied by `oai-authenticated-user-full-name-encoding: percent-encoded-utf-8`.

Treat the full name as optional and fall back to email when it is absent:

```tsx
import { headers } from "next/headers";

export default async function Home() {
  const requestHeaders = await headers();
  const userId = requestHeaders.get("oai-authenticated-user-id");
  const email = requestHeaders.get("oai-authenticated-user-email");
  const encodedFullName = requestHeaders.get("oai-authenticated-user-full-name");
  const fullName =
    encodedFullName &&
    requestHeaders.get("oai-authenticated-user-full-name-encoding") ===
      "percent-encoded-utf-8"
      ? decodeURIComponent(encodedFullName)
      : null;

  const displayName = fullName ?? email;
  // ...
}
```

## Optional Dispatch-Owned ChatGPT Sign-In

Import the ready-to-use helpers from `app/chatgpt-auth.ts` when the site needs optional or required ChatGPT sign-in:

- Use `getChatGPTUser()` for optional signed-in UI.
- Use the returned `userId` as the stable user key for user-owned records; do not use email as a durable identifier.
- Use `requireChatGPTUser(returnTo)` for server-rendered pages that should send anonymous visitors through Sign in with ChatGPT.
- In a Server Component, start sign-in with `<a href={chatGPTSignInPath(returnTo)} target="_top">`. The auth helper module is server-only; do not import it into a Client Component.
- Do not use `fetch`, XHR, a client-side router, or a framework link that can prefetch the sign-in route. SIWC must start as a top-level navigation.
- Never request the AuthAPI authorization endpoint directly. The dispatch-owned `/signin-with-chatgpt` route must start the SIWC flow.
- Use `chatGPTSignOutPath(returnTo)` for browser sign-out links or actions.
- Pass a same-origin relative `returnTo` path for the destination after sign-in or sign-out. The helper validates and safely encodes it.
- Mark protected pages with `export const dynamic = "force-dynamic"` because they depend on per-request identity headers.

Dispatch owns `/signin-with-chatgpt`, `/signout-with-chatgpt`, `/callback`, the OAuth cookies, and identity header injection. Do not implement app routes for those reserved paths. Routes that do not import and call the helper remain anonymous-compatible.

SIWC establishes identity only; it does not prove workspace membership. Use the Sites hosting platform's access policy controls for workspace-wide restrictions, or enforce explicit server-side membership or allowlist checks.

Use SIWC for account pages, user-specific dashboards, saved records, and write actions tied to the current ChatGPT user. Leave public content anonymous.

## Local D1 migrations

For a D1-backed local preview, generate SQL with `npm run db:generate`. Build once through the Sites skill's build entrypoint (or `npm run build` for standalone use) to generate `dist/server/wrangler.json`, rebuilding if bindings change. From the project root, apply each pending migration in order:

```sh
node --import ./scripts/sites-env.mjs ./node_modules/wrangler/bin/wrangler.js d1 execute DB --local --config dist/server/wrangler.json --persist-to .wrangler/state --file drizzle/0000_example.sql
```

Replace the filename with the pending migration and `DB` with your D1 binding name if different. Use `.wrangler/state`, not `.wrangler/state/v3`; Wrangler adds the versioned directories. Do not replay migrations already applied locally. This updates only the preview database; publishing applies production migrations separately.

## Diagnostic Commands

- `npm run install:ci`: perform the one locked dependency install
- `npm run dev`: start the Vite/Vinext development server
- `npm run build`: build the deployable Sites artifact
- `npm run start`: preview the built Worker locally with D1/R2 support
- `npm run db:generate`: generate Drizzle migrations after schema changes

When using the Sites plugin, follow its skill instructions for installation, builds, and publishing. These npm commands remain available for standalone use.

The portable build runs Vinext directly without a host `timeout` command. The managed-linux build uses `scripts/build-verified.sh` and its existing `SITES_BUILD_TIMEOUT` setting.

## Learn More

- [vinext Documentation](https://github.com/cloudflare/vinext)
- [Drizzle D1 Guide](https://orm.drizzle.team/docs/get-started/d1-new)
