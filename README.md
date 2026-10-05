# IMAP Migrator — Email Migration, Backup & Restore

**Unshackle yourself from your email provider. Take your mailbox with you.**

IMAP Migrator is a free, open-source Python command-line tool for copying email between IMAP accounts, creating incremental local EML backups, and restoring archived messages to an IMAP server. Migrate your email history, keep a local archive, or switch providers without deleting your source messages. Resumable transfers, Thunderbird OAuth for Gmail and Outlook, and independent byte-exact verification help you check what actually arrived.

## Your email history should survive your next provider change

- **Transfer email between accounts:** copy messages and map the source folder hierarchy into a dedicated destination root. Folder names and nesting must be compatible with the destination server.
- **Back up your mailbox to EML files:** export individual messages with their original fetched bytes, plus JSON `.metadata` files recording folders, timestamps, flags, and SHA-256 hashes.
- **Refresh an incremental IMAP backup:** rerun export to add new messages, check existing files, and refresh metadata. Previously exported messages removed from the source are retained; import uses the latest completed snapshot.
- **Restore email from a local archive:** upload exported EML messages to another IMAP account without needing the original provider online.
- **Migrate using different internet connections:** export on the connection that works well for the source, then import on the connection that works well for the destination.
- **Resume an interrupted email migration:** keep the SQLite journal and rerun. The script tracks accepted uploads and checks uncertain outcomes before attempting recovery.
- **Check whether your provider changed your mail:** independently fetch and compare source and destination content, distinguish exact preservation from recognized formatting changes, and report missing or damaged copies.

The source account remains read-only. Normal copying does not delete destination mail either; explicit repair can replace tracked damaged copies after verifying their replacements.

## Gmail, Outlook, Hotmail, Yandex, GMX, and other IMAP accounts

Looking to migrate Gmail to Outlook, transfer Outlook or Hotmail email to another provider, copy Yandex Mail to GMX, or move your mailbox to a self-hosted IMAP server? These are the kinds of account-to-account migrations this tool is designed for, provided both accounts expose compatible IMAP access and authentication.

Password and supplied OAuth access-token authentication are available. For Gmail and Microsoft Outlook/Hotmail, the script can also read an existing Thunderbird OAuth refresh token and request a fresh IMAP access token. Thunderbird must already have the account configured with OAuth2.

Provider rules still apply: folder nesting, quotas, supported flags, and message rewriting differ between servers. The built-in provider notices record observations rather than promising universal compatibility or identical storage behavior. See the verification and provider-notice sections below.

## Common email migration and backup questions

### How do I copy all my email to another account?

Define named accounts in one INI, then run `python3 imap-migrator.py migrate --from yandex --to hotmail --root Yandex`. Keep the journal to resume. The `verify` command independently compares both sides and writes a TSV report. Commands and configuration details follow below.

### Can I download my Gmail or Outlook mailbox as EML files?

Yes, when the account permits IMAP access. Use `export --from gmail --to /path/to/archive` to create a local email backup. Each message has its own EML file and companion metadata; filenames use UIDs rather than potentially unsafe subject text. Export does not require a destination account.

### Can I import an email backup into a different provider?

Use `import --from /path/to/archive --to hotmail --root Backup` with an archive produced by this script. Import restores the latest completed snapshot, verifies local message hashes before connecting, and performs full destination verification. Arbitrary EML collections without the archive metadata are not supported.

### Will my emails, attachments, dates, and read/unread status be preserved?

Export preserves the fetched message bytes, including attachments. Migration and import transfer the original IMAP INTERNALDATE and the portable read, answered, flagged, and draft flags. A destination provider may rewrite message bytes or reject metadata; verification reports those differences. Checked-equivalent formatting is reported separately and still does not pass as byte-identical. The script cannot force a provider to store messages verbatim.

### Is this a mailbox sync tool or a one-way migration tool?

It is a one-way email migration, backup, and restore tool. It does not mirror source deletions or provide continuous two-way synchronization. An ordinary rerun resumes copying and restores missing tracked messages; the `repair` command additionally checks and repairs tracked damaged destination copies.

## Requirements and setup

Python **3.10 or newer**, compatible IMAP access, and implicit TLS on port 993. The script uses Python's standard library. Thunderbird OAuth additionally requires the system NSS library (`libnss3`). No pip packages are needed.

```bash
cp imap-migrator.ini.sample imap-migrator.ini
chmod 600 imap-migrator.ini
```

Edit the account sections with your own server, username and authentication settings. Account names are arbitrary, case-sensitive aliases; `[gmail]`, `[work]` or `[account@gmail.com]` all work. Configure as many accounts as needed and choose their roles per command. Only the selected accounts' connection settings are required.

Choose exactly one authentication method per account:

| Setting | Meaning |
|---|---|
| `password` | IMAP password or provider-specific app password |
| `token` | Current OAuth **access token** accepted by IMAP; not a refresh token |
| `thunderbird_profile` | Reuse an existing Thunderbird Google/Microsoft OAuth credential |

For Thunderbird OAuth, `user` must match the saved account address. Thunderbird must already use OAuth2 for it. Set `thunderbird_primary_password` only if its profile requires one. The profile is read without modification; refresh tokens and credentials are not saved in archives or journals. Reauthorize in Thunderbird if a refresh token is revoked.

The optional `[defaults]` section sets `state_dir`, `batch_messages`, `batch_mib` and `retries`. Relative `state_dir` and Thunderbird profile paths are resolved beside the INI. Command-line paths are relative to the current directory. `--config PATH` selects a different INI.

## Commands

### Copy directly between accounts

```bash
python3 imap-migrator.py migrate --from yandex --to hotmail --root Yandex
```

All selectable source folders are mapped below the dedicated destination root, such as `Yandex/INBOX`. Rerun the same command to resume or restore missing tracked messages. Normal migration uploads mail and creates/subscribes folders; it does not delete destination messages.

Reverse the account roles or choose another destination without editing the INI. Use a dedicated root containing only this migration. Existing messages are adopted by whole-message content after line-ending normalization and duplicate multiplicity, never by Message-ID alone. Unmatched existing destination messages stop copying into that folder.

### Export a local backup

```bash
python3 imap-migrator.py export --from yandex --to ./yandex-backup
```

Only the selected source account is connected. Export creates exact fetched EML bytes and metadata, with no SQLite journal. Rerun with the same archive directory to refresh or resume. Add `--full-verify` to re-fetch existing source bodies and compare them with the archive.

### Import an archive

```bash
python3 imap-migrator.py import --from ./yandex-backup --to hotmail --root Yandex
```

Only the destination account is connected. Import validates the completed local snapshot and all active message hashes before connecting, then copies/resumes and performs full destination verification. It can restore to the same account originally exported. Arbitrary EML collections without this script's metadata are not supported.

Export and import can run on different networks, computers or days. Transfer the complete archive directory, including metadata.

### Verify without uploading or deleting

```bash
# Compare live accounts.
python3 imap-migrator.py verify --from yandex --to hotmail --root Yandex --report verification.tsv

# Compare an archive with its destination.
python3 imap-migrator.py verify --archive ./yandex-backup --to hotmail --root Yandex --report verification.tsv
```

`verify` always re-fetches complete destination bodies; a live source is independently fetched too. It creates/subscribes no destination folders and uploads/deletes no messages. It refreshes local journal inventories and writes the report. `--from ACCOUNT` and `--archive DIRECTORY` are mutually exclusive.

### Repair tracked destination copies

```bash
python3 imap-migrator.py repair --archive ./yandex-backup --to hotmail --root Yandex --report repair.tsv
# Or use --from yandex instead of --archive ./yandex-backup.
```

Repair works for unfinished migrations too. It restores missing mail and attempts replacements for tracked damaged copies. Old copies are deleted only after replacements pass the existing content/portable-flags/INTERNALDATE repair policy. Selective deletion requires UIDPLUS. Untracked destination messages are not automatically deleted.

Failed candidates and old copies remain journaled. Reruns recheck a surviving candidate instead of appending endless duplicates. A confirmed missing, uncommitted candidate can be retired and uploaded again; committed cleanup records require more conservative recovery. A provider that consistently changes uploaded messages may prevent repair from succeeding.

The repair policy can accept checked-equivalent content; final verification still requires byte equality. A completed repair therefore does not necessarily produce a byte-identical final result.

### Show provider observations

```bash
python3 imap-migrator.py providers
```

This command needs no INI or network connection. Destination notices also appear during migration/import/verification/repair.

## Useful options

| Option | Commands | Purpose |
|---|---|---|
| `--config PATH` | All | Select the account INI; may precede or follow the command |
| `--progress` | All | Enable elapsed timestamps, waiting indicators and live statistics |
| `--log PATH` | All | Append plain status/results; TSV reports hold message details when requested |
| `--report PATH` | Migration/import/verify/repair | Write a TSV report; **required** for verify and repair |
| `--journal PATH` | Migration/import/verify/repair | Select an explicit existing/new journal |
| `--full-verify` | Migrate/export | Re-fetch existing bodies rather than rely on cached content/archive checks |
| `--skip-tree` | Migrate/import/repair | Skip upfront tree preparation; check/create folders as reached |
| `--retry-pending` | Migrate/import/repair | Explicit recovery retry when an uncertain upload has no new destination UIDs |
| `--resolve-pending-uid UID` | Migrate/import/repair | Explicitly identify an uncertain uploaded destination UID after inspecting it |

Use `python3 imap-migrator.py --help` or `python3 imap-migrator.py COMMAND --help` for command-specific help. Progress is off by default; ordinary status output is plain. Copy counters count uploads during the current run, not previously copied messages or verification successes. Import's source-read byte counter measures local archive reads.

## Verification and TSV reports

The final summary separates byte-identical messages, line-ending-only changes, header-only changes and other/uncertain content differences. Categories are mutually exclusive, with percentages based on matched messages. Missing/unmatched source messages, extra/unmatched destination copies, portable flag differences and INTERNALDATE differences are separate counts. Repair runs also report uploaded replacements, completed repairs and failed candidates.

Byte differences remain verification failures even when decoded payloads match. Header-only means different headers with matching decoded MIME structure/payload; it does not validate signatures or prove complete equivalence. Only `\Seen`, `\Answered`, `\Flagged` and `\Draft` are applied to the destination. `\Recent`, `\Deleted` and custom keywords are excluded. INTERNALDATE is the server timestamp, not the message's `Date:` header.

Reports are UTF-8 TSV files readable in a spreadsheet or with tab-delimited tools. Message rows include:

- Phase, status and content category.
- Source/destination folders, UIDVALIDITY and UIDs; Subjects, Dates and local source EML paths.
- Original and line-ending-normalized SHA-256 hashes, sizes and line-ending counts.
- Header fields added/removed/changed, decoded MIME comparison, flags and server dates.
- Repair state, retained old UID, candidate identifiers and explanatory details.

For header changes, filter **`phase=verification`** and **`content_category=headers_only`**. Repair events use separate rows; count final verification rows when assessing the destination after repair. Unique unmatched content correspondences are explicitly labeled and do not alter journal bindings; ambiguous duplicates are not paired by guesswork.

No per-message display limit applies to reports. Message diagnostics go to TSV rather than flooding the terminal/text log. Reusing a report path overwrites an existing report with the same schema; unrelated existing files are refused. Rows are flushed as written. The final `phase=run` row distinguishes a complete verification pass from an interrupted/failed run. Tabs/newlines/backslashes inside cells are escaped; spreadsheet formula-like strings are stored as text.

Exit codes: **0** verified byte-identical, **1** operation error, **2** verification differences, **130** interrupted. Successful export also returns 0. Migration's normal final pass refreshes metadata but may use cached hashes; use `verify` or `migrate --full-verify` for independent body downloads. Import, verify and repair always use full verification.

## Journals and resuming

Journals are automatically selected under `[defaults] state_dir` (default: `imap-migrator-state` beside the INI). Their identity includes the source/destination server and username, destination root and, for archive operations, the archive ID. Changing direction, destination or root selects a different journal. Aliases, passwords and command names do not change the identity: verify/repair reuse the corresponding migration/import journal.

Keep the state directory across runs and keep it with the INI when moving installations. A journal stores hashes, UIDs, flags, dates and bindings, not credentials or message bodies. Journals and newly created archive files use private Unix permissions. Locks prevent simultaneous use of a journal/archive; read commands retry with bounded backoff.

Use `--journal PATH` if you want to choose a journal location explicitly; use the same path for subsequent related commands. A journal rejects different accounts, roots or archive identities. Keep it when recovering pending uploads or failed candidates.

APPEND intent is recorded before uploading, and acknowledged UIDs are saved durably. An uncertain upload is never blindly retried. If recovery cannot identify its result uniquely, it stops with the pending record retained. `--retry-pending` allows one explicit retry only when repeated inventories find no new destination UIDs; a delayed server commit still carries duplicate risk. If candidate UIDs exist, inspect raw messages before using `--resolve-pending-uid UID`. These two options are mutually exclusive and unavailable in `verify`.

## Archive format

Each message uses `uid-000000000123.eml` and a companion `.metadata` file, independent of potentially unsafe subject text. `folder.metadata` records folder identity; `archive.metadata` records source identity and filesystem mappings; `snapshot.metadata` lists the latest completed snapshot. Preserve all metadata for import.

Message metadata includes original UID/UIDVALIDITY, folder, flags, INTERNALDATE with timezone, byte counts and SHA-256. EML/metadata modification times use INTERNALDATE; folder metadata and directories use the newest message timestamp. File managers display timestamps in the local timezone.

Portable folder names remain readable. Reserved names, unsafe characters, excessive path depth and case/Unicode collisions use stable aliases; exact source names stay in metadata. Aliases do not bypass destination IMAP nesting limits.

Export refreshes add new messages, check existing files and update metadata. Files removed from the source are retained locally, but import restores only the latest completed snapshot. This is not selectable historical snapshots or a version history of flag changes. An interrupted refresh must finish before import; source UIDVALIDITY changes require a new archive directory.

Writes use synced temporary files and atomic replacements. Archives are not automatically deleted after upload and contain no credentials. SHA-256 detects accidental corruption; the archive is not signed or encrypted. Keep an independent backup.

## Provider notices and limits

Provider notices record observations, not universal guarantees. Current examples include Outlook/Hotmail header normalization, Gmail line-ending rewriting and slow uploads, and GMX's observed three-level folder hierarchy limit, counting the destination root. Unknown hosts are explicitly marked unknown. Notices never relax verification or authorize duplicate matching.

Upfront destination-tree preparation creates missing parents, checks writable selection and subscribes folders for client visibility before uploading/deleting messages. Subscription failures warn; quota and message acceptance are still checked during transfer. `--skip-tree` skips this preflight, so later folder errors can occur after earlier folders have copied.

Keep the source and dedicated destination root stable during migration/verification. IMAP does not provide an atomic account snapshot. Unsupported hierarchy delimiters, namespace ambiguity and mapping collisions are rejected; source-folder aliases are not silently invented for the destination.

Transfers overlap bounded source prefetch with serial destination uploads. Defaults are 25 messages/16 MiB per batch; a larger single message is handled alone. Multiple batches and MIME parsing can coexist in memory. No MULTIAPPEND, parallel destination sessions or guaranteed provider-independent throughput is promised.
