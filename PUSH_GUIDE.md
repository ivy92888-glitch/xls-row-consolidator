# Push guide

This repository was published to GitHub using the REST API (no `git` client
required and no large push tooling). The same method works for any small repo.

## What was pushed

Eight files were committed to the `main` branch:

```
src/RowRollup.bas
README.md
LICENSE
docs/Configuration.md
examples/sample-data.xlsx
PUSH_GUIDE.md
.gitignore
.gitattributes
```

## How it was done (summary)

1. Created the empty repo via `POST /user/repos`.
2. Activated it (a placeholder file was created via `PUT /repos/.../contents/<file>`).
3. For each real file, created a blob (`PUT /repos/.../contents/<path>`) with
   base64 content and a commit message.
4. Verified with `GET /repos/.../git/trees/main?recursive=1`.

All calls authenticated with a classic personal access token (scope `repo`).
The token was supplied at runtime and was **never written to disk**.

## Security notes

- A PAT is equivalent to a password. After finishing all publishes, **delete or
  rotate the token** at GitHub ▸ Settings ▸ Developer settings ▸ Personal access
  tokens.
- This repo contains no credentials, no private data and no secrets.
- If you re-run a publish, prefer a fine-grained token scoped to this single
  repository.
