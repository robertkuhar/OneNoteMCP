# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Safety

The server talks to a real OneNote account, and its tools can overwrite and delete notes. `updatePageContent` replaces the entire page body. Don't start the server or call its tools against live notebooks unless asked; use the mocked Graph client in tests instead.

## Commands

- Build: `npm run build`
- Test: `npm test`
- Single test: `npx jest -t "getNotebook tool"`
- Lint: `npm run lint`
- Run the server: `npm run run` (needs a `.env`; see the README)

These commands differ from what you might expect:

- `npm run clean` is Windows-only.
- Lint already reports 99 errors, so check for new ones rather than expecting a clean run.

## Architecture

`src/index.ts` sets up auth and the Graph client, then passes them to `registerNotebookTools`, `registerPageTools` and `registerSectionTools` in `src/functions/`. The server uses stdio.

- **Stdout is reserved for the MCP protocol.** Log with `console.error`. `utils/logger.ts` is unused, and its `info` and `debug` levels write to stdout.
- **The compiled output is CommonJS.** `index.ts` and `jsonCache.ts` call `require("dotenv")`, so adding `"type": "module"` would break them.
- **Auth uses a device code login.** Tokens are cached as JSON in `AUTH_CACHE_DIR`, which must be set or `jsonCache.ts` throws on import. `AZURE_CLIENT_SECRET` must be set too, but it is only checked, never passed to the credential. The first login has to be run manually, before Claude Desktop launches the server.
- **Keep `utils/tokenCachePlugin.ts`.** It replaces `@azure/identity-cache-persistence`, whose plugin fails with a `node:path` resolution error when Claude launches the server.

### Adding or changing tools

- Follow the existing handlers: a zod schema, a JSON `resource` result with the Graph path as its `uri`, and errors rethrown as `Failed to ...: ${getErrorMessage(error)}`.
- Fetch Graph collections with `getAll` from `utils/graph.ts`, because Graph returns only 20 items per page by default. `listNotebooks` doesn't use it yet.
- Limit concurrent Graph calls with `mapWithConcurrency` to avoid throttling.
- `searchPages` deliberately walks each section and matches client-side. `/me/onenote/pages` fails on accounts with many sections, and client-side matching avoids OData filter injection.

## Fork

`origin` is robertkuhar/OneNoteMCP and `upstream` is swax/OneNoteMCP. Open PRs against `origin` unless asked otherwise; `gh` may default to the parent repo on a fork, so pass `--repo robertkuhar/OneNoteMCP`.
