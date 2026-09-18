# Liquid4All host

Published static files for the **Liquid4All** Excel Office.js add-in (Blazor WebAssembly).

This repository is the free HTTPS host for GitHub Pages. Application source stays in the private repo `Liquid4All-AppSource` (`C:\Daten\Projects\WMB\Mitosoft.Liquid4All`). This folder is **not** a project in `Mitosoft.Liquid4All.slnx`: it is its own git repo with no `.csproj`, only published static files.

| What | URL |
|------|-----|
| Task pane | https://wolfgangmenabruhn.github.io/liquid4all-host/ |
| Manifest | https://wolfgangmenabruhn.github.io/liquid4all-host/manifest.xml |
| Support | https://wolfgangmenabruhn.github.io/liquid4all-host/support.html |

Workbook data stays in the Excel file. This site only serves the task pane.

## Use in Excel

1. Confirm GitHub Pages serves the task-pane URL over HTTPS.
2. Sideload `manifest.xml` (this file, or the URL above) into desktop Excel.
3. Open **Home → Liquid4All**.

## Update the host after an application change

Do this whenever you change UI, business logic, assets, or `manifest.xml` in the **private source** repo and you want Excel (and later AppSource) to load that build.

Folders on this machine:

| Role | Path |
|------|------|
| Source (you edit here) | `C:\Daten\Projects\WMB\Mitosoft.Liquid4All` |
| Host clone (this repo) | `C:\Daten\Projects\WMB\liquid4all-host` |

Host origin: `https://wolfgangmenabruhn.github.io/liquid4all-host/`

### Preferred: `Liquid4All.HostPublish`

In the source solution, run the **HostPublish** project (Visual Studio: set startup to `Liquid4All.HostPublish` and Start — not the Excel multi-startup profile), or:

```powershell
cd C:\Daten\Projects\WMB\Mitosoft.Liquid4All
dotnet run --project Liquid4All.HostPublish\Liquid4All.HostPublish.csproj
```

That publishes Release, replaces generated files in this clone, keeps `.nojekyll` / `404.html` / `support.html` / `README.md`, patches `index.html`, and rewrites `manifest.xml`. It does **not** `git commit` or `git push`. Then continue at [Commit and push](#6-commit-and-push-this-repo).

If this clone is not at `..\liquid4all-host`, pass `--host-dir <path>` or set `LIQUID4ALL_HOST_DIR`.

The numbered steps below are the same work, done by hand.

### 1. Finish the change in source

1. Implement and test locally (Visual Studio **Excel** profile, or `npm run start` in `Liquid4All.ExcelLaunch`).
2. If this is a release, bump versions together:
   - `manifest.xml` → `<Version>`
   - `Liquid4All.OfficeAddIn`, `Liquid4All.Business`, `Liquid4All.ViewModels` `.csproj` → `<Version>`
3. Keep `excelInterop.js` navigation based on `<base href>` (needed for this Pages subpath). Keep icons under `Liquid4All.OfficeAddIn/wwwroot/assets/`.

### 2. Publish the WASM files

From the **source** repo:

```powershell
cd C:\Daten\Projects\WMB\Mitosoft.Liquid4All
dotnet publish Liquid4All.OfficeAddIn\Liquid4All.OfficeAddIn.csproj -c Release -o .\publish\addin
```

The website to copy is `publish\addin\wwwroot\` (not the folder above it).

### 3. Copy published files into this host repo

Overwrite the previous Blazor output, but **do not delete** host-only files.

**Copy (replace):** everything under `wwwroot` (`_framework\`, `css\`, `assets\`, `index.html`, `excelInterop.js`, and the `.br` / `.gz` siblings).

**Keep as they are in this repo:**

| File | Why |
|------|-----|
| `.nojekyll` | GitHub Pages must serve `_framework` (underscore folders). |
| `404.html` | SPA fallback for client routes such as `/accounts`. |
| `support.html` | Public support URL for AppSource. |
| `README.md` | These instructions. |
| `.gitignore` | OS junk only. |

**Do not copy** the source `manifest.xml` as-is (it still points at `https://localhost:7177`). Rebuild it in the next step.

Example copy (PowerShell, from the source repo):

```powershell
$src = "C:\Daten\Projects\WMB\Mitosoft.Liquid4All\publish\addin\wwwroot"
$dest = "C:\Daten\Projects\WMB\liquid4all-host"
Copy-Item -Path "$src\*" -Destination $dest -Recurse -Force
if (-not (Test-Path "$dest\.nojekyll")) { New-Item -ItemType File -Path "$dest\.nojekyll" | Out-Null }
```

### 4. Point `index.html` at the Pages subpath

`dotnet publish` writes `<base href="/" />`. GitHub Pages lives under `/liquid4all-host/`, so after the copy:

1. Set `<base href="/liquid4all-host/" />`.
2. Keep the small SPA restore `<script>` immediately after `<body>` (see the current `index.html`). If the copy wiped it, paste it back from git (`git checkout -- index.html` will restore the *old* file; better to re-apply the two edits on the new `index.html`).

```powershell
$indexPath = "C:\Daten\Projects\WMB\liquid4all-host\index.html"
$html = Get-Content -Raw $indexPath
$html = $html.Replace('<base href="/" />', '<base href="/liquid4all-host/" />')
$spa = @'

    <script type="text/javascript">
      (function (l) {
        if (l.search[1] === '/') {
          var decoded = l.search.slice(1).split('&').map(function (s) {
            return s.replace(/~and~/g, '&');
          }).join('?');
          window.history.replaceState(null, null, l.pathname.slice(0, -1) + decoded + l.hash);
        }
      }(window.location));
    </script>
'@
if ($html -notmatch 'l.search\[1\] === ''/''') {
    $html = $html.Replace('<body>', '<body>' + $spa)
}
Set-Content -Path $indexPath -Value $html.TrimEnd() -Encoding utf8NoBOM
```

### 5. Rebuild `manifest.xml` for the host URL

Take the **source** `manifest.xml` and replace every `https://localhost:7177` with the Pages origin:

```powershell
$hostRoot = "https://wolfgangmenabruhn.github.io/liquid4all-host"
$manifest = Get-Content -Raw "C:\Daten\Projects\WMB\Mitosoft.Liquid4All\manifest.xml"
$manifest = $manifest.Replace("https://localhost:7177/", "$hostRoot/")
$manifest = $manifest.Replace("https://localhost:7177", "https://wolfgangmenabruhn.github.io")
$manifest = $manifest.Replace(
    '<SupportUrl DefaultValue="' + $hostRoot + '/" />',
    '<SupportUrl DefaultValue="' + $hostRoot + '/support.html" />')
Set-Content "C:\Daten\Projects\WMB\liquid4all-host\manifest.xml" -Value $manifest.TrimEnd() -Encoding utf8NoBOM
```

Check that `SupportUrl` is `.../support.html`, `AppDomain` is `https://wolfgangmenabruhn.github.io`, and `SourceLocation` / icons / `Taskpane.Url` use `.../liquid4all-host/`.

Leave the source `manifest.xml` on localhost so F5 sideload still works.

### 6. Commit and push this repo

```powershell
cd C:\Daten\Projects\WMB\liquid4all-host
git add -A
git status
git commit -m "Publish updated Liquid4All WASM host."
git push origin main
```

GitHub Pages rebuilds from `main` `/`. Wait until the site shows **built** (repo **Settings → Pages**, or a minute or two).

### 7. Confirm Excel is on the new build

1. Open https://wolfgangmenabruhn.github.io/liquid4all-host/ — the pane shell should load (full features need Excel).
2. In Excel, close and reopen the Liquid4All task pane (or re-sideload `manifest.xml` if you changed `<Version>` or URLs).
3. Smoke-test the change you just shipped.

The private source repo does **not** need a commit for the host to update. Only this public repo must be pushed.
