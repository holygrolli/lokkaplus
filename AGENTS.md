## AGENTS — Guidance for AI coding agents

This file gives concise, actionable guidance for AI coding agents working in this repository. Link to fuller docs rather than duplicating them.

### Quick summary
- **Project**: Lokka — an MCP server for Microsoft Graph and Azure RM.
- **Server entrypoint**: built JS at [src/mcp/build/main.js](src/mcp/build/main.js#L1) (built from TypeScript in [src/mcp/src/main.ts](src/mcp/src/main.ts#L1)).

### Common commands
- **Build server**: `npm run build` (run in `src/mcp`). See [src/mcp/package.json](src/mcp/package.json#L1).
- **Start server**: `node build/main.js` (after building) or `npm run start` in `src/mcp`.
- **Website**: docs are a Docusaurus site. From `website`: `npm run start` (dev) and `npm run build` (static site). See [website/package.json](website/package.json#L1).

Example quick-start (from repo root):

```
cd src/mcp
npm install
npm run build
npm run start
```

### Key files & docs (link, don't copy)
- README: [README.md](README.md#L1)
- MCP server package: [src/mcp/package.json](src/mcp/package.json#L1)
- Server docs and testing notes: [src/mcp/README.md](src/mcp/README.md#L1), [src/mcp/TESTING.md](src/mcp/TESTING.md#L1)
- Server source: [src/mcp/src/main.ts](src/mcp/src/main.ts#L1) and other `.ts` files in `src/mcp/src`.
- Website docs: [website/docs](website/docs)

### What agents should know
- TypeScript sources for the MCP server live under `src/mcp/src`. Build with `tsc` (script provided).
- The published package uses `build/` as the distributable; modifications to runtime behavior usually require rebuilding.
- Environment configuration is via environment variables documented in [src/mcp/README.md](src/mcp/README.md#L1). Follow the README for auth modes (`USE_CLIENT_TOKEN`, `USE_INTERACTIVE`, `USE_CERTIFICATE`, `USE_GRAPH_BETA`).

### Conventions & constraints
- Prefer editing TypeScript files in `src/mcp/src` and run `npm run build` before testing runtime changes.
- Keep docs in `website/docs`; edits to them affect the static site — run `npm run start` from `website` to preview.
- Node engine for the website is `>=18` (see `website/package.json`). Server code targets modern Node versions; use Node 18+ when running locally.

### When to create/update `AGENTS.md` vs `.github/copilot-instructions.md`
- Prefer `AGENTS.md` for repository-scoped, developer-facing guidance for AI agents. Use `.github/copilot-instructions.md` only for organization-level or GitHub-specific instructions.

### Suggested next customizations
- Add a minimal `.github/copilot-instructions.md` that points to this file and lists local dev commands (optional).
- Create small skills for `run-local-server`, `generate-docs-preview`, and `token-management` to automate common tasks for agents.

---
If you'd like, I can also add a minimal `.github/copilot-instructions.md` pointing here and containing a small checklist for pull requests.
