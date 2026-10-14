# Contributing

## Branches

| Branch | Purpose |
|---|---|
| `main` | Releases only – one merge commit per release, tagged `vX.Y.Z` |
| `dev` | Integration branch for the next release |
| `feature/<name>`, `bugfix/<name>` | Work branches, created from `dev` |

1. Create a branch from `dev`: `git switch dev && git pull && git switch -c feature/<name>`.
2. Open a pull request into `dev`.
3. Merge it with **Squash and merge** – one commit per feature or fix on `dev`.
4. Do not commit directly to `main` or `dev`.

## Release

1. Update `APPLICATION_VERSION` in `source/modules/_config_Application.bas` on `dev`
   (via a `release/X.Y.Z` branch and pull request).
2. Pull request `dev` -> `main`, merged with **Create a merge commit** (title `Release X.Y.Z`).
   No squash, so `main` and `dev` keep a shared history. The workflow `sync-dev-with-main.yml`
   then fast-forwards `dev` to `main`, so no merge back is needed.
3. Tag the merge commit on `main` and push only this tag:
   ```
   git switch main && git pull
   git tag -a vX.Y.Z -m "Release X.Y.Z"
   git push origin vX.Y.Z
   ```
4. Publish a GitHub release for the tag with the release notes. The workflow
   `build-release-zip-file.yml` builds the add-in and attaches `ACLibImportWizard_vX.Y.Z.zip`.

`git log --first-parent main` lists the releases only.

## Development

- The add-in is built from the source files in `source/`, exported with the
  [Version Control add-in](https://github.com/joyfullservice/msaccess-vcs-integration).
  Do not commit Access binaries (`.accda`, `.accdb`).
- Edit and export the add-in in a trusted location; import or rebuild from `source/`
  after editing the source files.
- Code, comments and commit messages in English; commit subject in imperative mood, max. ~72 characters.
- Text files use CRLF line endings; keep the encoding of exported files (UTF-8 with BOM in `source/`).
- Use copies of databases for tests (trusted location), never your real databases.
