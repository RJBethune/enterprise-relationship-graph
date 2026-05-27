# CLAUDE.md — Executive Office workspace

Project-specific rules for Claude. These supplement the global project instructions.

## 5. Editing Safety

**Back up before any Edit or Write.** No exceptions.

Before touching a file with the `Edit` or `Write` tool, snapshot it. The cheapest backup that works depends on git state:

- If the file is tracked and `git status <file>` is clean: no copy needed. `git show HEAD:<file>` is your restore path.
- If the file is tracked but dirty, or untracked: copy to `/tmp/cowork-backups/<filename>` via bash before the first edit. Mkdir the folder if missing.
- If you'll make several edits, one backup at the start is enough — it's the pre-edit state you need, not a snapshot per change.

Record the backup path in your reply so the user knows where to recover from if something goes wrong.

## 6. Verify after every edit

The `Edit` and `Write` tools on this workspace have silently corrupted a large file twice (see "Known issue" below). After each edit on any file ≥ 50 KB, verify:

1. **Size delta is sane.** Run `wc -c <file>` via bash and confirm the size grew (or shrank) by approximately the bytes you added (or removed). If a pure-addition edit left the file the same size, the tail was truncated.
2. **Tail is intact.** Run `tail -5 <file>` and confirm the closing tokens are still there — `</html>`, `}`, the final statement, whatever the file's natural ending is. If the tail is mid-statement or mid-tag, the file is corrupted.
3. **Syntax parses, if applicable.** For HTML with embedded JS, extract the `<script>` body and run `node --check` on it. For Python, `python -m py_compile`. For JSON, `python -m json.tool`.

If verification fails, do not continue editing. Restore from the backup and use the workaround below.

## 7. Known issue: large-file truncation on this workspace

**Symptom.** When `Edit` or `Write` adds N bytes to a file on the `C:\Repos\ExecutiveOffice` mount, the file's tail can get silently truncated by approximately the same N bytes. The file's reported size stays close to its original size; the addition is real (visible in `grep`) but matching bytes disappear from the end.

**Observed on.** `enterprise-relationship-graph.html` (~240 KB) — twice, on May 25 and May 26, 2026. Both times the truncation broke `boot();` and the surrounding toolbar wiring, leaving the file ending mid-statement.

**Root cause.** Not confirmed. Strongly correlated with file size near or above ~240 KB. Bash `>` redirect writes the same content without truncating, so the issue is in the file tools or their mount adapter, not the underlying filesystem.

**Workaround — bulk-edit via Python + bash redirect.** Use this when `Edit`/`Write` is unsafe (any file > 50 KB on this workspace, or any file where verification has failed once):

1. **Read the source of truth into `/tmp`.** Either `cat <file> > /tmp/work.<ext>` or `git show HEAD:<file> > /tmp/work.<ext>` if you want the last clean version.
2. **Apply all edits in one Python pass.** Use exact-string `str.replace()` with these guards:
   - The anchor must appear exactly once in the source — assert this before replacing.
   - The anchor must be present — fail loudly if missing (don't silently skip).
   - Do every edit you need in this one pass; don't loop back to `Edit` later.
3. **Write back via bash redirect.** `cat /tmp/work.<ext> > <file>`. Bash redirect overwrites in place without unlinking — important because the workspace mount blocks `rm` on existing files.
4. **Verify per Section 6.** Size, tail, syntax.

A reference implementation of the Python pass lives in the session transcript from May 26, 2026 (the `enterprise-relationship-graph.html` stamp-feature edits).

## 8. Workspace mount quirks (good to know)

These are properties of this mount, not bugs to work around — just things that surprise you:

- **No `rm` on existing files.** Bash cannot unlink files in the workspace. It *can* overwrite them in place via `>` redirect or truncate them to empty via `: > file`. If you create a stray test file you can't delete, truncate it to 0 bytes and tell the user.
- **`git checkout -- <file>` fails with "Operation not permitted"** for the same reason — checkout tries to unlink first. Use `git show HEAD:<file> > <file>` instead.
- **File mtimes don't update reliably** via `Edit`/`Write`. Don't use mtime to decide whether a file was modified; use content hashes or `git diff`.
- **`Read` can only access mounted folders** (the workspace and the session's outputs/uploads). It can't read `/tmp` paths. If you build content in `/tmp`, route it back through bash to view it.
