# Excel MCP Server — Agent Rules

## Project Overview

This is a **Model Context Protocol (MCP) server** for Excel operations, built with TypeScript. It provides tools for reading, writing, formatting, and analyzing Excel files (.xlsx, .xls) with enterprise-grade security. The server integrates with GitHub Copilot, Cline, and other MCP-compatible clients in VS Code.

## Assigned Expert Roles & Multi-Perspective Evaluation

When handling any request, the AI assistant operates across **six specialized expert personas**. Every proposed plan, code change, and output MUST be evaluated through each role's lens:

1. **Expert Software Solution Architect**: Enforces Clean Architecture, Single Responsibility Principle (SRP), Facade pattern, modular decoupling, and long-term maintainability.
2. **Senior Software Developer**: Enforces clean code, idiomatic TypeScript (ES2022, ESM, strict mode, zero unapproved `any`), optimal performance, and defensive error handling.
3. **Senior QA & Test Engineer**: Scrutinizes edge cases, validates data boundaries, writes/runs tests, and enforces green builds (`npm run build` with zero warnings/errors).
4. **Financial & Accounting Data Analyst**: Guarantees mathematical correctness for financial formulas (NPV, IRR, VAT, amortizations, rounding, accounting formats, auditability).
5. **System Security Specialist**: Enforces non-bypassable `PermissionChecker` controls, prevents path traversal, validates sizes/extensions, and eliminates secret or path leakage.
6. **Senior Excel & Technical Documentation Specialist**: Ensures deep Excel domain correctness (ExcelJS models, styling, number formats, cell/range semantics) and writes precise, professional technical documentation.

## Architecture

### Clean Architecture Layers

```
src/
├── index.ts                          # Entry point — bootstraps server
├── types/                            # Shared type definitions (no logic)
├── security/                         # Permission checking & access control
│   └── permission-checker.ts
├── services/                         # Business logic (Excel operations)
│   ├── excel-service.ts              # Facade — delegates to specialized services
│   ├── excel-workbook-manager.ts     # Workbook open/close/save lifecycle
│   ├── excel-cell-operations.ts      # Cell read/write/copy/find-replace
│   ├── excel-formatting.ts           # Font, alignment, border, color, number formats
│   ├── excel-structure-operations.ts # Worksheet add/delete/rename/copy
│   ├── excel-accounting.ts           # Financial calculations & formats
│   ├── excel-advanced-accounting.ts  # NPV, IRR, amortization, tax, currency
│   └── excel-formula-analyzer.ts     # Formula parsing & analysis
├── tools/                            # MCP tool layer
│   ├── tool-definitions.ts           # Tool schemas (Zod + JSON Schema)
│   ├── tool-handler.ts               # Tool dispatch & execution
│   ├── definitions/                  # Per-category tool definition files
│   └── handlers/                     # Per-category handler files
├── server/                           # MCP server wiring
│   └── excel-mcp-server.ts
└── utils/
    └── logger.ts                     # Logging utility
```

### Design Principles

1. **Single Responsibility** — Each service file handles one domain (cells, formatting, accounting, etc.).
2. **Facade Pattern** — `excel-service.ts` is the central facade that delegates to specialized services.
3. **Security First** — All file operations MUST pass through `PermissionChecker` before execution.
4. **Type Safety** — Strict TypeScript (`strict: true`) with `noUnusedLocals`, `noUnusedParameters`, `noImplicitReturns`.

## Coding Conventions

### TypeScript

- Target: **ES2022**, Module: **ESNext**, Module Resolution: **bundler**.
- Use **ESM** (`"type": "module"` in package.json). Do NOT use CommonJS `require()`.
- All types go in `src/types/index.ts`. Do NOT scatter type definitions across service files.
- Use `zod` schemas for input validation of tool parameters.
- Use `zod-to-json-schema` to generate JSON schemas for MCP tool definitions.

### Naming

- Files: `kebab-case.ts` (e.g., `excel-cell-operations.ts`).
- Classes: `PascalCase` (e.g., `ExcelService`, `PermissionChecker`).
- Methods/functions: `camelCase`.
- Constants: `UPPER_SNAKE_CASE`.
- Tool names: `snake_case` prefixed with `excel_` (e.g., `excel_read_cell`, `excel_set_font_style`).

### Error Handling

- Wrap tool handler logic in try/catch. Return MCP-compatible error responses (never throw unhandled).
- Include descriptive error messages with context (file path, cell reference, etc.).
- Security violations must return clear permission-denied messages without leaking internal paths.

## Adding New Tools

When adding a new MCP tool, follow this checklist:

1. **Define the Zod schema** in `src/tools/definitions/` (or `tool-definitions.ts`).
2. **Implement the handler** in `src/tools/handlers/` (or `tool-handler.ts`).
3. **Add the business logic** in the appropriate service file under `src/services/`.
4. **Register the tool** in `tool-definitions.ts` so the MCP server exposes it.
5. **Add permission checks** — use `PermissionChecker` for any file system operation.
6. **Document the tool** in `docs/tools/` with parameters, examples, and category.
7. **Update `README.md`** tool tables.

## Security Rules

- **NEVER** bypass `PermissionChecker`. All file paths must be validated.
- **NEVER** access paths outside `MCP_ALLOWED_PATHS` or inside `MCP_DENIED_PATHS`.
- **NEVER** process files exceeding `MCP_MAX_FILE_SIZE`.
- Only allow whitelisted extensions: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`.
- Environment-sensitive config (`.env`) must NOT be committed. Use `.env.example` as template.

## Dependencies

| Package | Purpose |
|---------|---------|
| `@modelcontextprotocol/sdk` | MCP protocol implementation |
| `exceljs` | Excel file read/write engine |
| `zod` | Runtime input validation |
| `zod-to-json-schema` | Schema generation for MCP tools |

Do NOT add new dependencies without justification. Prefer `exceljs` APIs over alternative Excel libraries.

## Development Workflow

```bash
npm run dev        # Development with hot reload (tsx watch)
npm run build      # Compile TypeScript → dist/
npm run start      # Run compiled server
npm test           # Run Jest tests
npm run lint       # ESLint check
npm run lint:fix   # ESLint auto-fix
npm run format     # Prettier formatting
```

- Node.js **>= 18.0.0** is required.
- Always run `npm run build` and verify no TypeScript errors before committing.
- Keep `dist/` out of version control (already in `.gitignore`).

## Standard Operating Procedure (SOP) for Tasks

All AI agents must follow the 5-stage loop defined in [docs/AI_WORKFLOW.md](docs/AI_WORKFLOW.md) and evaluate every task against the **6 Assigned Expert Roles**:

1. **Context Discovery**: Read existing types (`src/types/index.ts`) and relevant service files before modifying anything. Never edit blindly.
2. **Layered Planning**: Plan changes along the architectural hierarchy (`Types → Domain Services → Security → Tool Layer → Docs`) and conduct an architectural/security pre-evaluation.
3. **Safe Execution**: Follow TypeScript strict conventions, handle errors gracefully with try/catch, preserve financial numerical precision, and never crash the MCP server process.
4. **Automated Verification**: Always run `npm run build` and tests to verify TypeScript compiles with zero errors before reporting completion.
5. **Multi-Role Handoff**: Review and report changes through the lenses of Architect, Developer, Tester, Security Specialist, Accounting Analyst, and Excel/Doc Specialist with direct clickable file links.

See [docs/AI_WORKFLOW.md](docs/AI_WORKFLOW.md) for full guidelines, C-O-C-V prompting templates, and Definition of Done.

## Documentation

- `docs/AI_WORKFLOW.md` — AI assistant SOP, collaboration rules & Definition of Done.
- `docs/PROJECT_AUDIT_REPORT.md` — Comprehensive 6-role project audit report & roadmap.
- `docs/ARCHITECTURE.md` — Detailed architecture documentation.
- `docs/API.md` — Full API reference.
- `docs/SECURITY.md` — Security model and configuration.
- `docs/TROUBLESHOOTING.md` — Common issues and solutions.
- `docs/tools/` — Per-category tool documentation with examples.
- `docs/SETUP_GUIDE.md` — Installation and Cline integration guide.
- `docs/ANTIGRAVITY_SETUP.md` — Google Antigravity configuration.
- `docs/GITHUB_COPILOT_SETUP.md` — GitHub Copilot configuration.
- `docs/CLINE_SETUP.md` — Cline extension configuration.
