# IMAP migrator

A TLS-only, standard-library Python IMAP migration script. Python 3.10 or newer is required. Thunderbird Microsoft/Google OAuth additionally requires the system NSS library (`libnss3`). Normal migration never deletes mail from either endpoint. Explicit `--repair` can delete specific journaled destination UIDs after verifying replacements; the source is always read-only. The source is opened read-only, and bodies are fetched with `BODY.PEEK[]`.

## Run

Keep `imap-migrator.py` and your existing `imap-migrator.ini` together:

```bash
chmod 600 imap-migrator.ini
python3 imap-migrator.py
```

The previous configuration remains compatible. Additional settings in the revised sample are optional. Copy authentication details into your existing INI; do not replace them with the sample placeholders. Thunderbird profiles are read without modification.

To independently re-download and verify both sides:

```bash
python3 imap-migrator.py --verify-only --full-verify
```

Other options:

```bash
python3 imap-migrator.py --config /path/to/imap-migrator.ini
python3 imap-migrator.py --full-verify
python3 imap-migrator.py --help
```

Exit codes: `0` byte-identical and verified, `1` operation stopped/failed, `2` verification failed, `130` interrupted.

## Existing migration attempts

The dedicated destination root must contain only this migration. Before uploading into a populated folder, the script compares existing destination content against the source, including the number of occurrences of identical messages. It stops that folder before uploading if destination messages cannot be matched. It does not infer identity from Message-ID alone, silently ignore changed content, or delete excess messages.

Your reported `Arc` counts (249 source versus at least 292 destination) will trigger this check if those counts remain unchanged. To perform a clean migration without touching earlier attempts, choose both a new root and a new journal:

```ini
[migration]
destination_root=Yandex-new
journal=imap-migrator-new.sqlite3
```

Do not delete the earlier root until the new migration has passed `--verify-only --full-verify` and you have reviewed it. Changing accounts or roots while keeping the same journal is rejected deliberately.

## Resume and interruption

Keep the SQLite journal (`imap-migrator.sqlite3` by default) across runs. It contains folder names, UIDs, hashes, dates, flags, and source-to-destination bindings, but no credentials or message bodies. A process lock prevents two runs from using the same journal simultaneously. The journal is created with private Unix permissions.

Read operations reconnect and retry with bounded backoff. Uploads use durable write-ahead records: the pending record is committed before sending APPEND. An accepted APPENDUID is saved before the pending record is cleared. If the server accepts an upload but its reply is lost, recovery examines new destination UIDs and requires one unique matching message. If recovery cannot establish the outcome, the script stops and retains the pending record. An unaccepted upload and an accepted upload whose content was rewritten can be indistinguishable after a disconnect; safely resolving every such case automatically is impossible without additional server guarantees.

An explicit tagged APPEND rejection clears the pending record and stops with the server error. Interrupted reads can be retried; an uncertain APPEND is never blindly retried. Do not remove the journal to bypass an unresolved upload: that discards the evidence needed to prevent duplicates. Inspect the destination and pending record before making any manual repair.

Destination UID assignments are retained even when later content verification finds differences. Accepted messages whose returned content differs produce warnings, and copying continues through all folders. Rerunning reports the differences without appending the same journaled source message again. Exact verification still returns exit code 2 when content differs; it does not silently certify altered messages. If a mapped destination UID disappears, an ordinary rerun now restores that source message without deleting any destination mail (patch 007). Accepted copies that still exist but differ require explicit repair for replacement.

## Performance and memory

A source worker downloads ahead while the destination uploads. Each connection has a single owner. The queue is bounded, and the source worker sends keepalive NOOPs while waiting for uploads. The consumer can reconnect safely if the source closes during the final batch.

Both message count and estimated byte size limit batches. Defaults are 25 messages and 16 MiB. A single larger message is transferred alone; this is not an absolute per-message memory limit. Several batches can coexist briefly in memory, and parsing/hashing makes temporary byte copies. Complete folder bodies are not retained. Metadata and hashes use memory proportional to message count.

Uploads remain serial on one destination connection; downloaded bodies overlap those uploads. Accepted destination bodies are read back in batches, avoiding a verification request per message when APPENDUID is available. The script does not implement MULTIAPPEND or multiple concurrent upload sessions. Provider latency, throttling, and upload limits still apply; there is no promised speedup multiplier.

First-time adoption of an existing populated folder requires downloading both sides to establish content identity. Normal resumes download only bodies missing from the journal, plus new messages being transferred. Metadata is refreshed each run. Without APPENDUID, each successful upload requires additional reconciliation, so that fallback is slower.

## Verification

Every accepted upload is read back and checked, in batches. By default the final pass checks current UID inventories, flags, dates, sizes, and cached content hashes. This is explicitly reported as cached verification, not an independent fresh body comparison. UIDVALIDITY identifies the lifetime of UID assignments; a changed validity triggers reindexing, and a change during reconnect stops the operation.

Use `--full-verify` for independent body downloads from both endpoints. Exact preservation compares original bytes. A separate conservative fingerprint classifies recognized formatting and transfer representation changes as “equivalent under checked rules.” Text payload line-ending differences can qualify for checked equivalence; binary payload bytes must match. Equivalence does not produce an exact-preservation success: byte differences still return exit code 2. Flags and INTERNALDATE are checked independently. Unparseable dates are errors, never silently treated as matching. Content rewritten by the destination is reported as a verification difference even if the server accepted the upload. With complete journal bindings, verification distinguishes accepted-but-changed messages from missing messages and compares flags/dates separately from content hashes. Copying continues despite these differences. For the first changed message in a folder, a diagnostic lists added, removed, and changed header field names and compares decoded MIME payloads/structure. It prints no header values or message bodies. This diagnostic does not relax verification or authorize adoption of preexisting messages.

Only `\Seen`, `\Answered`, `\Flagged`, and `\Draft` are transferred. `\Deleted`, server-owned `\Recent`, and custom keywords are excluded. Mailboxes may not support every portable flag; any resulting flag difference is reported.

Keep both accounts stable during migration and verification. Concurrent folder/message/flag changes can cause a stop or failed verification. Checks detect changes during individual folder scans, but IMAP does not provide an atomic snapshot of an entire account. Leave the dedicated destination root untouched until the migration is complete.

Source hierarchy components that contain the destination delimiter are rejected before copying because they cannot be mapped losslessly. Ambiguous destination namespaces and case-insensitive mapping collisions are rejected too.

## Offline regression tests

```bash
python3 -m unittest -v test_imap_migrator.py
```

The tests use a real `imaplib` client against an in-process local IMAP test server, with synthetic mail only. They cover normal transfer/resume, literal-tail metadata, case-insensitive flags, read reconnects, lost APPEND replies, missing APPENDUID, duplicate multiplicity, destination rewriting, pending recovery, changed UIDVALIDITY, quota rejection, disappeared destination mail, journal locking/identity, folder mapping, dates, and byte-limited batches. They do not exercise live Yandex/Outlook, TLS handshakes, or Thunderbird OAuth.


## Provider notices and checked equivalence (patch 004)

Before either connection opens, the script warns about recorded behavior for the selected destination host. Untested hosts are labeled UNKNOWN. Use `--provider-notes` to show the built-in table without configuration or a connection. See `IMAP-PROVIDER-NOTES.md` for the current evidence and how to add another tested provider.

Message comparisons report BYTE IDENTICAL, EQUIVALENT UNDER CHECKED RULES; ORIGINAL BYTES CHANGED, or CHANGED OR UNCERTAIN / INCOMPLETE. Provider entries never bypass checks. Changes to recipients, display names, unknown headers, body/attachment payloads, filenames, or MIME structure are not automatically accepted as harmless. Signed messages and malformed/unsupported MIME remain unclassified when bytes differ. Unknown header values and repeated field order are preserved. The checker does not authenticate mail or validate signatures.

The journal stays compatible. On the first run after patch 004, cached bodies without the new versioned equivalence fingerprint are downloaded once and rehashed. Source-to-destination UID bindings are preserved, so accepted messages are not uploaded again. Subsequent runs reuse the enriched cache. For an independent comparison, use `--verify-only --full-verify`.

The semantic fingerprint is used for reporting and, from patch 006, uniquely checked-equivalent pending-upload reconciliation. Legacy destination adoption still uses the existing exact-content/multiplicity rule.


## Repair missing or damaged destination messages (patch 005)

After the current copying process has stopped, apply patch 005 after patches 001–004, keep the existing journal, and run:

```bash
python3 imap-migrator.py --repair --full-verify
```

`--repair` forces full body downloads even without `--full-verify`. It is mutually exclusive with `--verify-only`. It first resumes any pending uploads and repair cleanup, then checks the source and destination folders and copies missing messages. For journaled messages with content/metadata differences, it appends a replacement, reads the replacement back, and requires exact bytes or checked equivalence **plus matching portable flags and INTERNALDATE** before committing the replacement mapping. Recognized equivalent formatting alone never triggers replacement.

Only after the replacement passes verification does repair mark the specific old destination UID as Deleted and issue `UID EXPUNGE` for that UID. It never issues ordinary EXPUNGE or CLOSE. If an existing message must be replaced and the server cannot selectively expunge a UID, repair refuses before appending its replacement. Restoring a missing message needs no deletion capability.

The journal records candidate UIDs, original source/old-copy hashes, and pending cleanup. An interruption can resume the same candidate or cleanup without appending another replacement. If the replacement fails verification, the old copy and candidate are retained, the original mapping remains, and the final check reports differences. A subsequent repair pass rechecks that same candidate rather than creating repeated duplicates. A failed candidate that still fails needs inspection; it is not automatically discarded or replaced with endless new candidates.

Repair leaves untracked mail alone. Existing untracked messages may be adopted only under the original exact-content/multiplicity rule; unidentified extras cause a stop. It does not propagate source deletions, delete obsolete folders, or clean unrelated duplicates. Source/destination UIDVALIDITY changes, source-content changes during an outstanding repair, or changes to the old target cause a stop before deleting it.

A final independent full verification follows repairs. Exit code 2 can still mean all mail is present but a provider reformatted it into checked-equivalent messages. Outstanding failed candidates and incomplete cleanup also prevent success. Keep the source originals and journal until you have reviewed the final results.

Use the same command to resume an interrupted repair. An ordinary migration refuses to proceed while repair candidates/cleanup remain; `--verify-only --full-verify` can inspect them without modifying messages.


## Recover an unresolved APPEND (patch 006)

Pending recovery now reports the destination UIDVALIDITY, last UID before the upload, new UID count, exact matches, checked-equivalent matches, and acknowledged APPENDUID when available. If no exact match exists but precisely one new candidate passes conservative checked equivalence against the original journaled source fingerprint, the pending upload is recovered automatically. This does not mark it byte-identical; full verification still reports original-byte differences.

If recovery reports no new destination UIDs, you can explicitly allow one retry:

```bash
python3 imap-migrator.py --repair --full-verify --retry-pending
```

The script reconnects to the destination, checks UIDVALIDITY and inventories again, validates the original source UID/content, and updates the durable pending record in place before uploading. It refuses if new destination UIDs exist, and never retries automatically after a second lost reply. A delayed original server-side commit remains possible: this option explicitly accepts a residual duplicate risk. Keep the journal/source and inspect the final full verification.

If candidate UIDs exist but neither exact nor checked equivalence identifies the message, inspect the raw source and destination candidate messages. Only after identifying the correct uploaded message, use:

```bash
python3 imap-migrator.py --repair --full-verify --resolve-pending-uid 123
```

Replace 123 with the actual destination UID, not the source UID or a sequence number. This explicitly asserts upload identity; it does not certify content. The UID must exist, be new or acknowledged for the pending upload, and not belong to another mapped message/repair. Repair still verifies a replacement before deleting an old copy. The retry and explicit-UID options are mutually exclusive and cannot be used with `--verify-only`.

Patch 006 was produced without running tests at the user's request.


## Copy/resume versus repair (patch 007)

The script does not rely on the user deciding that migration is finished. Each run derives remaining work from source/destination inventories and the journal. Repair is also valid for an unfinished migration.

| Command | Operation |
|---|---|
| No flags | Copy/resume, including restoring absent journaled messages; no destination deletion |
| `--repair` | Copy/resume plus verified replacement/deletion; full verification is implied |
| `--verify-only --full-verify` | Independent comparison; no upload/deletion |
| Add `--retry-pending` when recovery reports zero new UIDs | Explicitly allow one ambiguous-upload retry with residual duplicate risk; not a general resume requirement |

`--repair --full-verify` is accepted for compatibility but prints that full verification is already implied. `--repair --retry-pending` is sufficient when both operations are intended. An unresolved outcome is not treated as a confirmed missing message: ordinary missing-message restoration does not silently retry an ambiguous APPEND. If an older repair has unfinished candidate/cleanup records, default copying stops with directions to use `--repair`; it does not silently gain permission to delete mail.

The startup mode line describes the selected behavior. Known destination folder-depth limits are checked for the entire mapping before copying or deleting mail, rather than discovering the limit halfway through the migration. GMX's recorded limit is three total levels, counting the destination root. See `IMAP-PROVIDER-NOTES.md`. Deeper mappings need a compatible destination or a deliberate remapping plan; this patch does not silently change existing folder names.

Patch 007 and the updated regression expectations were not tested, at the user's request.

Gmail OAuth: set `server=imap.gmail.com`, `user=` to the exact account address saved in Thunderbird, and `thunderbird_profile=` to its profile directory. Thunderbird must already access that account using OAuth2. The script selects the saved Google credential with the `https://mail.google.com/` scope and refreshes it using Thunderbird’s public installed-app client details. Calendar/contacts-only tokens are excluded. No profile changes or token persistence are performed. If Google revokes the refresh token, reauthorize the account in Thunderbird and rerun. Use a separate journal when changing destination accounts/providers.

Before normal copying, repair, or pending-upload recovery, the script prepares the entire mapped destination tree: it creates missing parent folders first and checks writable selection of every message folder. Preparation failure aborts before message uploads/deletions; successfully created folders remain for resuming. It then subscribes folders, including parent containers, for mail-client visibility; subscription failures warn and do not prevent copying. Existing folders are reused. `--verify-only` creates/subscribes nothing. Writable selection cannot guarantee APPEND permissions, available quota, or acceptance of every message; copying and verification still check those operations.

Status output uses monotonic elapsed timestamps `[HH:MM:SS.hh]` since startup, with hundredths of a second and no initial zero-time banner. Connections, source/destination enumeration, counting, copying, hashing, recovery, repair, and verification notices are timestamped. Normal copying displays separate folder/overall processed, remaining, copied-this-run, and already-mapped counters; processed counts are not verification results. The folder summary also reports destination occupancy at the time of its scan. Copy speed counts uploads rather than already-mapped messages; elapsed copy time uses hours/minutes/seconds. Progress counters have no timestamp prefixes; the speed line retains one elapsed-copy-time field. Operation messages remain timestamped. Terminals redraw the progress block; redirected output records periodic snapshots. Repair retains its operation-specific results rather than using normal-copy counters.

Destination preparation logs each folder, CREATE, writable SELECT, return to EXAMINE, and subscription before sending the operation; per-folder and total preparation durations are shown. It reuses the destination listing already obtained for mapping and sends source NOOP keepalives between preparation operations when the source has been idle for over 30 seconds. A single blocking network operation can still exceed that interval; normal read reconnect handling remains in place. OAuth token-refresh completion and IMAP-authentication completion are logged separately to distinguish authentication latency from folder preparation.

When both source and destination folders are empty, normal copying emits one timestamped “Nothing to copy” line instead of a progress block. Overall counters remain unchanged. Nonempty destination folders are explicitly reported and still undergo reconciliation; final source-change checks, folder preparation, and verification remain in place for empty folders.

Use `./imap-migrator.py --skip-destination-tree-deploy` to resume without repeating the upfront destination-tree creation, writable-selection checks, and subscriptions. It also works with `--repair`. Mapping/collision checks, known provider-depth checks, source counting, journal recovery, and content verification still run. Each folder is checked/created as reached during copying or repair, so a missing/new folder can still be restored; later folder errors may therefore occur after earlier folders have copied. The option does not change the journal or assume that messages are complete, and it has no effect with `--verify-only`.


## Offline migration and incremental EML backup

```bash
./imap-migrator.py --export --path /path/to/archive
# Change connection/network, then:
./imap-migrator.py --import --path /path/to/archive
```

Export connects only to the source; import connects only to the destination. Export needs only a configured `[source]`; `[migration]` is optional for batching/retry defaults. Import needs `[destination]` and `[migration]` with `destination_root`; it reads the original source identity from the archive, not source credentials. Import works with `--repair`, `--verify-only`, pending-recovery options, and `--skip-destination-tree-deploy`. Normal migration without archive flags remains available. `--path` is required for both archive modes and cannot be used alone.

Example layout for source folders `INBOX` and `Arc|Family` (source delimiter `|`):

```text
archive/
    archive.metadata
    snapshot.metadata
    archive.lock
    INBOX/
        folder.metadata
        uid-000000000123.eml
        uid-000000000123.metadata
    Arc/
        Family/
            folder.metadata
            uid-000000000456.eml
            uid-000000000456.metadata
```

EMLs are exact fetched bytes. Each UTF-8 JSON `.metadata` records the source folder, UIDVALIDITY, UID, original INTERNALDATE string and Unix epoch, portable flags, original flags (excluding the session-only `\Recent`), reported RFC822.SIZE, actual byte count, and SHA-256. Folder metadata contains folder identity only, not a collection of message records. The archive header maps original folders to filesystem directories. The snapshot records active folder identities and UIDs only. Empty selectable folders are retained.

Both message files get modification times from the source IMAP INTERNALDATE, including its timezone offset. The timestamp represents one instant; file managers display it in the computer's timezone. The original date/offset remain in metadata. This is not the message's `Date:` header or the filesystem creation time. File access times may change when files are read.

Readable folder components are preserved when portable. Windows-reserved names/characters, trailing spaces/dots, overlong or deep paths, case/Unicode-normalization collisions, and archive-reserved filenames use stable generated aliases. Hierarchy delimiters become filesystem separators. Exact original names and delimiters remain in metadata and control import; directory aliases do not rename destination IMAP folders or bypass destination nesting limits. Very deep paths may be represented by a flat alias directory. Keep the archive base path reasonably short on Windows.

Rerun export with the same path to refresh/resume. Existing complete pairs are hash checked, their flags metadata is refreshed if needed, and their EML bodies are reused. `--export --full-verify` also fetches existing source bodies and checks that their bytes still match the archive; identical EMLs are not rewritten. Source UIDVALIDITY changes stop export without reassigning old files; use a new archive directory. Export retains old messages/folders removed from the source, but import includes only the latest completed snapshot, not all retained history. This is retained-message backup, not a versioned history of flag changes or multiple independently selectable snapshots.

Each EML is committed before its companion metadata through `.partial` files and atomic replacements. Files are flushed and synced; newly created archive files/directories have private permissions on Unix. An exclusive archive lock prevents simultaneous export/import by this script. Refresh marks the archive incomplete before changing message metadata; interrupted refreshes must finish before import. The previous snapshot file and historical EMLs are retained, but an incomplete archive is deliberately not importable. Files are never automatically removed after upload. SHA-256 detects corruption, not malicious replacement of both a message and its metadata; the archive is not signed or encrypted.

Import validates every active EML/metadata pair and hash before destination connection, then reuses normal pending-upload recovery, content comparison, copying, and repair. Full verification is always enabled against the local snapshot. Import uses a separate journal filename with `-import-<archive-id>` inserted before its extension; preserve that journal for resuming. The original live-migration journal is left separate. Restoring to the same provider/account is allowed because the source is an immutable local snapshot; choose the destination root deliberately. Import restores the script's existing portable flags (`\Seen`, `\Answered`, `\Flagged`, `\Draft`) and INTERNALDATE; arbitrary source keywords and `\Deleted` are retained in backup metadata but not applied to destination messages.

Verification proves preservation of exported bytes and portable metadata, not that the live source remains unchanged. Export rescans folder/message metadata before marking completion, but IMAP cannot provide an atomic snapshot of an actively changing account. Repeated source changes can require another export attempt. Credentials are never written into the archive. No tests were run for patch 015.
