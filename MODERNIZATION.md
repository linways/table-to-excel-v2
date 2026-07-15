# Modernization: v1.0.4 → v2.0.0

Branch: `modernize/esm-build-v2` · Commit: `5c1f273`

This document explains **what broke**, **why**, and **exactly what changed** to fix
it. The goal of v2 was to make `@linways/table-to-excel` load in a modern
bundler (Vite / Rollup / esbuild) **without any consumer-side workaround**.

---

## 1. The bug we fixed

### Symptom

Importing the library inside a Vite-built app threw at **module load time** —
before any table was ever exported:

```
TypeError: Cannot read properties of undefined (reading 'TYPED_ARRAY_SUPPORT')
```

### Root cause

v1.0.4 shipped `dist/tableToExcel.js` as a **Parcel v1 UMD bundle**. Two facts
about that bundle combined to crash in the browser:

1. **Parcel passes a hidden 4th argument to every module factory.** Parcel's
   runtime invokes each wrapped module like this:

   ```js
   moduleFn.call(exports, require, module, exports, this)
   //                                              ^^^^  → readable as arguments[3]
   ```

   In **Node**, top-level `this` is the Node `global` object, so `arguments[3]`
   hands each module a reference to `global`.

2. **The inlined `buffer@4.x` polyfill relies on that back door.** It does:

   ```js
   var t = arguments[3];          // expects Node global
   ... t.TYPED_ARRAY_SUPPORT ...  // reads a property off it
   ```

When the same bundle is served by Vite as a **native ES module**, it runs in
**strict mode**, where top-level `this` is `undefined`. Parcel therefore
forwards `undefined` as the 4th argument, `arguments[3]` is `undefined`, and
`undefined.TYPED_ARRAY_SUPPORT` throws the instant the module is imported.

Where the old `buffer@4` polyfill came from: v1 imported ExcelJS through a
**deep dist path** —
`../node_modules/exceljs/dist/es5/exceljs.browser` — which is exceljs 1.x's
pre-bundled ES5 build that inlines `buffer@4`. That is the exact code Parcel
wrapped and shipped.

### The old consumer-side workaround (now removed)

The FAPI app previously carried a Vite build-time plugin
(`vite-plugins/fix-table-to-excel-globals.ts`) that regex-rewrote every
`arguments[3]` in the bundle to `(arguments[3] || globalThis)`, in **two**
plugin shapes (esbuild for the dev prebundle, Vite `transform` for the Rollup
production build). v2 makes that plugin unnecessary.

---

## 2. What changed

### 2.1 `src/tableToExcel.js` — import ExcelJS by package name

```diff
-import ExcelJS from "../node_modules/exceljs/dist/es5/exceljs.browser";
+import ExcelJS from "exceljs";
```

Removes the deep dist path that pulled in the `buffer@4` polyfill. The bundler
now resolves `exceljs` normally and picks its modern browser build.

### 2.2 `exceljs` 1.x → 4.4.0

`package.json` dependency bumped `^1.10.0` → `^4.4.0`. ExcelJS 4 uses native
`Uint8Array` and does **not** ship the crashing `buffer@4` read-of-`global`
form. ExcelJS is kept **external** (not inlined) so consumers bundle a single
clean copy of `exceljs@4`.

### 2.3 Build tool: Parcel → tsup (esbuild)

```diff
- "build": "parcel build src/tableToExcel.js"
+ "build": "tsup src/tableToExcel.js --format esm,cjs --clean --sourcemap"
```

tsup emits standard, spec-compliant ESM and CJS with **no** Parcel runtime
wrapper and **no** hidden `arguments[3]` convention — so the whole bug class is
gone at the source, not patched over.

### 2.4 `package.json` — dual ESM/CJS package metadata

| Field         | v1.0.4                    | v2.0.0                                            |
| ------------- | ------------------------- | ------------------------------------------------- |
| `version`     | `1.0.4`                   | `2.0.0`                                            |
| `type`        | *(none, CJS)*             | `module`                                          |
| `main`        | `./dist/tableToExcel.js`  | `./dist/tableToExcel.cjs`                          |
| `module`      | —                         | `./dist/tableToExcel.js`                           |
| `browser`     | —                         | `./dist/tableToExcel.js`                           |
| `exports`     | —                         | import→`.js`, require→`.cjs`, default→`.js`        |
| `sideEffects` | —                         | `["./dist/tableToExcel.js", "./dist/tableToExcel.cjs"]` |

`sideEffects` lists the dist files because the module has a real side effect —
it assigns `window.TableToExcel = TableToExcel` — so tree-shakers must not drop
it.

### 2.5 Test runner: Mocha/Babel → Vitest

```diff
- "test": "mocha -r @babel/register -r @babel/polyfill -r jsdom-global/register"
+ "test": "vitest run"
```

New `vitest.config.ts` runs the existing `test/parser.test.js` (29 chai
assertions) under **jsdom**. All 29 pass against exceljs 4.4 — confirming the
DOM-parsing / styling API surface is unchanged.

```ts
// vitest.config.ts
import { defineConfig } from 'vitest/config'
export default defineConfig({
  test: { environment: 'jsdom', globals: true, include: ['test/**/*.test.js'] },
})
```

Removed devDeps: `@babel/*`, `parcel-bundler`, `mocha`, `jsdom-global`, plus the
`babel` config block. Added: `tsup`, `vitest`, `typescript` (tsup needs it
present even for a JS entry), `jsdom@25`.

### 2.6 `dist/` rebuilt

Removed stale Parcel artifacts (`index.html`, `index.js`,
`tableToExcel.517b1af9.js`, `.map` files). New outputs:

- `dist/tableToExcel.js` (ESM) + `.js.map`
- `dist/tableToExcel.cjs` (CJS) + `.cjs.map`

---

## 3. Verification

| Layer                         | Check                                                                 | Result |
| ----------------------------- | --------------------------------------------------------------------- | ------ |
| **1. Build static analysis**  | `grep 'arguments\[3\]'` and `grep 'TYPED_ARRAY_SUPPORT'` in new dist   | 0 hits |
| **2. Unit tests**             | 29 parser tests under Vitest/jsdom on exceljs 4.4                      | Pass   |
| **3. Consumer build**         | FAPI `pnpm build` with the workaround plugin **removed**              | Pass   |

The v1 crash was a read of a property off `undefined` (`arguments[3].X`); a
grep of the v2 dist confirms zero such reads remain — the bug is structurally
absent, not merely masked.

---

## 4. Consumer migration

1. Point the dependency at v2 (git branch or, once published, `^2.0.0`).
2. Delete any `arguments[3]` / `TYPED_ARRAY_SUPPORT` Vite/esbuild workaround
   plugin and its registrations.
3. Clear the bundler cache (`rm -rf node_modules/.vite`) and re-export a sheet.

The runtime API is unchanged — `TableToExcel.convert(table, opts)` and the
`data-*` cell attributes work exactly as before.

### Follow-ups (not yet done)

- Ship `.d.ts` types so consumers can drop their local module shim.
- Add a UMD build if any `<script>`-tag consumers still need one.
