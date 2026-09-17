# Upload to konor123/xls2xlsx

The repository already exists:

https://github.com/konor123/xls2xlsx

## Option A — Git command line

Extract this package, open a terminal in the extracted folder, and run:

```powershell
git init
git branch -M main
git add .
git commit -m "Initial XLS2XLSX release"
git remote add origin https://github.com/konor123/xls2xlsx.git
git push -u origin main
```

If `origin` already exists:

```powershell
git remote set-url origin https://github.com/konor123/xls2xlsx.git
git push -u origin main
```

Then create the first release:

```powershell
git tag v0.1.0
git push origin v0.1.0
```

The GitHub Actions workflow will build `XLS2XLSX.exe` and attach it to the release automatically.

## Option B — GitHub CLI

If `gh` is installed and authenticated:

```powershell
gh auth status
git init
git branch -M main
git add .
git commit -m "Initial XLS2XLSX release"
git remote add origin https://github.com/konor123/xls2xlsx.git
git push -u origin main
git tag v0.1.0
git push origin v0.1.0
```

## Option C — GitHub website

1. Open `https://github.com/konor123/xls2xlsx`.
2. Choose **Add file → Upload files**.
3. Upload the extracted repository contents while preserving the folder structure.
4. Commit to `main`.
5. To create a release build, create a `v0.1.0` tag from Git or GitHub Desktop. The workflow reacts to `v*` tags.
