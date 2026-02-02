# SPFx React Toolkit - Project Overview

## Purpose

A React hooks library for SharePoint Framework (SPFx) development. Provides type-safe, state-managed hooks for accessing SPFx context, services, and APIs.

## Tech Stack

- **Language**: TypeScript 5.3+
- **React**: >= 17 (peer dependency)
- **SPFx**: >= 1.15.0 (peer dependency)
- **Build**: Gulp (SPFx toolchain)
- **Output**: CommonJS (lib/) + ESM support planned

## Project Structure

```
src/
├── core/           # Provider components and context
├── hooks/          # React hooks (main exports)
├── utils/          # Utility functions
├── extensions/     # SPFx extension examples
└── webparts/       # Demo web part for testing
docs/
└── api/hooks/      # API documentation (Markdown)
lib/                # Compiled output
```

## Commands

| Command | Purpose |
|---------|---------|
| `npm run build` | Build (gulp bundle) |
| `npm run serve` | Dev server with hot reload |
| `npm run lint` | ESLint check |
| `npm run test` | Run tests |
| `gulp bundle --ship` | Production build |
| `gulp package-solution --ship` | Create .sppkg |

## Naming Conventions

- Hooks: `useSPFx*` (e.g., `useSPFxContext`, `useSPFxMSGraphClient`)
- Interfaces: `SPFx*Info` or `SPFx*Result` for return types
- Options: `SPFx*Options` for configuration objects
- Files: kebab-case for utilities, PascalCase for components

## Code Style

- Functional components with hooks
- JSDoc for all public APIs
- `readonly` on interface properties
- Type-safe generics where appropriate
- Memory leak prevention with `isMountedRef` pattern

## Key Patterns

1. **Provider Pattern**: Wrap app in `SPFxReactToolkitProvider`
2. **Invoke Pattern**: `invoke(client => ...)` for state-managed API calls
3. **isReady Pattern**: Check `isReady` before using async-initialized resources
4. **Effect Separato**: Break circular dependencies with separate useEffects
