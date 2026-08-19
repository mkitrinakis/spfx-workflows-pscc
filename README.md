# pscc — Form Customizer (consumer)

Thin Form Customizer that imports shared form logic from **workflows-core** (SPFx Library).

## Projects

| Project | Role |
|---------|------|
| `../workflows-core` | Shared library (fields, renderer, services, styles) — component ID `c286440f-d2ae-43f8-90c9-392e23d89c94` |
| `pscc` (this repo) | List-specific entry `FormFormCustomizer.ts` — component ID `f79cca91-406d-4369-8e63-db4b71692545` |

## Local development

```powershell
# 1. Build library first (required after core changes)
cd ..\workflows-core
heft build

# 2. Install + start consumer
cd ..\pscc
npm install
heft start --clean
```

Edit existing item:

```powershell
heft start --clean --serve-config form_EditForm
```

## Deploy order

1. Deploy `workflows-core.sppkg` to tenant App Catalog
2. Deploy `pscc.sppkg`
3. Associate form customizer on the list; set `configFileUrl` in component properties
4. Upload `sample-configs/sample-document-form.xml` to SharePoint Site Assets

## Library dependency

`package.json` uses `"workflows-core": "file:../workflows-core"`.

After `npm install`, verify `node_modules/workflows-core/lib/index.js` exists (run `heft build` in workflows-core first).

At runtime, SharePoint loads the library component from the deployed library package; the extension manifest references library ID `c286440f-d2ae-43f8-90c9-392e23d89c94`.
