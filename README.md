# Liquid4All host

Published static files for the **Liquid4All** Excel Office.js add-in (Blazor WebAssembly).

This repository is the free HTTPS host for GitHub Pages. The product source stays in the private repo.

- Add-in: https://wolfgangmenabruhn.github.io/liquid4all-host/
- Manifest: https://wolfgangmenabruhn.github.io/liquid4all-host/manifest.xml
- Support: https://wolfgangmenabruhn.github.io/liquid4all-host/support.html

## Use in Excel

1. Wait until GitHub Pages serves the URL above (HTTPS).
2. Sideload `manifest.xml` (this file, or the URL) into desktop Excel.
3. Open **Home → Liquid4All**.

Workbook data stays in the Excel file. This site only serves the task pane.

## Update the host

From the private source repo:

```powershell
dotnet publish Liquid4All.OfficeAddIn\Liquid4All.OfficeAddIn.csproj -c Release -o .\publish\addin
```

Copy `publish\addin\wwwroot\*` into this repository, keep `.nojekyll`, restore `<base href="/liquid4all-host/" />` in `index.html`, then push `main`.
