#!/usr/bin/env python3
r"""TLS-only IMAP migration. Python 3.10+, standard library; NSS for Thunderbird OAuth.

Source mailboxes are EXAMINEd and bodies read with BODY.PEEK[]. Source mail is never deleted.
Normal migration never deletes mail. Explicit --repair may remove only journaled
destination UIDs after a replacement passes verification. Source stays read-only.
The SQLite journal contains hashes/UIDs/metadata, never passwords or mail bodies.
Keep it across runs. Only one process may use a journal at a time.

UIDVALIDITY and live UID inventories validate cached identities. Normal resumes
reuse content hashes; --full-verify re-downloads both endpoints independently.
An uncertain APPEND is NEVER blindly retried. If content cannot identify its
outcome uniquely, migration stops with its durable pending record intact.
Existing destination mail is adopted only by exact content (ignoring line endings)
and multiplicity. Unmatched destination messages stop that folder before copying.

Only portable flags Seen/Answered/Flagged/Draft are copied; Deleted is excluded.
The dedicated destination root must not be modified during migration. Keep the
source stable too: concurrent changes cause verification failure or a safe stop.
"""

# The INI path is resolved relative to this script, not the current directory.
INI_FILE = "imap-migrator.ini"

# Fixed by design: implicit TLS (IMAPS) only.
IMAP_PORT = 993
SOCKET_TIMEOUT = 120

# Thunderbird's current public Microsoft desktop OAuth application.
TB_MS_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
TB_MS_TOKEN_ENDPOINT = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
TB_MS_OAUTH_ORIGIN = "oauth://login.microsoftonline.com"
TB_MS_IMAP_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All"

# Thunderbird's public installed-app credentials: required to refresh tokens
# already issued to Thunderbird (not credentials belonging to a user).
# Source: mailnews/base/src/OAuth2Providers.sys.mjs in comm-central.
TB_GOOGLE_CLIENT_ID = "406964657835-aq8lmia8j95dhl1a2bvharmfk3t1hgqj.apps.googleusercontent.com"
TB_GOOGLE_CLIENT_SECRET = "kSmqreRr0qwBWJgbf5Y-PjSU"
TB_GOOGLE_TOKEN_ENDPOINT = "https://www.googleapis.com/oauth2/v3/token"
TB_GOOGLE_OAUTH_ORIGIN = "oauth://accounts.google.com"
TB_GOOGLE_IMAP_SCOPE = "https://mail.google.com/"

import base64
import builtins
import shutil
import unicodedata
import uuid
import collections
import datetime as _dt
import hashlib
import imaplib
import configparser
import ctypes
import ctypes.util
import json
import re
from pathlib import Path
import signal
import ssl
import urllib.error
import urllib.parse
import urllib.request
import sys
import os
import time
import sqlite3
import queue
import threading
import contextlib
import argparse
import email.parser
import email.policy
from dataclasses import replace
from dataclasses import dataclass


# All status times are monotonic elapsed time, independent of wall-clock changes.
RUN_STARTED = time.monotonic()
OUTPUT_LOCK = threading.RLock()
_progress_rows = 0
_progress_columns = None
LOG_FILE = None


def elapsed_stamp():
    ticks = max(0, int((time.monotonic()-RUN_STARTED)*100))
    seconds, hundredths = divmod(ticks, 100)
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return f'[{hours:02d}:{minutes:02d}:{seconds:02d}.{hundredths:02d}]'


def log(*args, sep=' ', end='\n', file=None, flush=True, file_only=False):
    global _progress_rows, LOG_FILE
    stream = sys.stdout if file is None else file
    text = sep.join(str(arg) for arg in args)
    with OUTPUT_LOCK:
        # A normal message freezes the last progress block, preventing subsequent
        # terminal redraws from erasing diagnostics (including worker messages).
        if not file_only:
            _progress_rows = 0
        stamp = elapsed_stamp()
        text = '\n'.join(stamp+' '+line if line else '' for line in text.split('\n'))
        if not file_only:
            builtins.print(text, end=end, file=stream, flush=flush)
        if LOG_FILE is not None:
            try:
                builtins.print(text, end=end, file=LOG_FILE, flush=True)
            except OSError as exc:
                failed_log, LOG_FILE = LOG_FILE, None
                with contextlib.suppress(OSError):
                    failed_log.close()
                builtins.print(stamp+' WARNING: log file write failed; file logging disabled: '+str(exc),
                               file=sys.stderr, flush=True)


def human_duration(seconds):
    seconds = max(0, int(seconds))
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return (f'{hours}h {minutes:02d}m {seconds:02d}s' if hours else
            f'{minutes}m {seconds:02d}s' if minutes else f'{seconds}s')


NSS_LOCK = threading.RLock()

STANDARD_FLAGS = {
    rb"\Seen": r"\Seen",
    rb"\Answered": r"\Answered",
    rb"\Flagged": r"\Flagged",
    rb"\Draft": r"\Draft",
}

_stop_requested = False


def _request_stop(signum, frame):
    global _stop_requested
    _stop_requested = True
    log("\nTermination requested; stopping at the next safe point...", flush=True)





def check_stop():
    if _stop_requested:
        raise KeyboardInterrupt


# ------------------------- IMAP modified UTF-7 -------------------------------

def _b64_mod_encode(raw: bytes) -> str:
    return base64.b64encode(raw).decode("ascii").rstrip("=").replace("/", ",")


def _b64_mod_decode(text: str) -> bytes:
    text = text.replace(",", "/")
    text += "=" * ((4 - len(text) % 4) % 4)
    return base64.b64decode(text)


def imap_utf7_encode(text: str) -> str:
    """Encode Unicode mailbox name as IMAP modified UTF-7 (RFC 3501)."""
    out = []
    non_ascii = []

    def flush():
        if non_ascii:
            s = "".join(non_ascii)
            out.append("&" + _b64_mod_encode(s.encode("utf-16-be")) + "-")
            non_ascii.clear()

    for ch in text:
        o = ord(ch)
        if 0x20 <= o <= 0x7E:
            flush()
            out.append("&-" if ch == "&" else ch)
        else:
            non_ascii.append(ch)
    flush()
    return "".join(out)


def imap_utf7_decode(text: str) -> str:
    """Decode IMAP modified UTF-7 mailbox name."""
    out = []
    i = 0
    while i < len(text):
        if text[i] != "&":
            out.append(text[i])
            i += 1
            continue
        j = text.find("-", i)
        if j < 0:
            raise ValueError(f"Malformed modified UTF-7 mailbox name: {text!r}")
        payload = text[i + 1:j]
        if payload == "":
            out.append("&")
        else:
            out.append(_b64_mod_decode(payload).decode("utf-16-be"))
        i = j + 1
    return "".join(out)


def quote_mailbox(name: str) -> str:
    """Return a safely quoted modified-UTF-7 mailbox argument."""
    enc = imap_utf7_encode(name)
    return '"' + enc.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _unquote_imap_token(token: str) -> str:
    token = token.strip()
    if token.startswith('"') and token.endswith('"'):
        body = token[1:-1]
        out = []
        escaped = False
        for ch in body:
            if escaped:
                out.append(ch)
                escaped = False
            elif ch == "\\":
                escaped = True
            else:
                out.append(ch)
        if escaped:
            out.append("\\")
        return "".join(out)
    return token


_LIST_RE = re.compile(
    r'^\((?P<flags>.*?)\)\s+(?P<delim>NIL|"(?:\\.|[^"])*")\s+(?P<name>.+)$'
)


@dataclass(frozen=True)
class Folder:
    name: str
    delimiter: str | None
    flags: frozenset[str]


def parse_list_line(line: bytes) -> Folder:
    s = line.decode("ascii", "strict")
    m = _LIST_RE.match(s)
    if not m:
        raise RuntimeError(f"Cannot parse IMAP LIST response: {line!r}")

    flags = frozenset(x.lower() for x in m.group("flags").split())
    delim_token = m.group("delim")
    delimiter = None if delim_token == "NIL" else _unquote_imap_token(delim_token)

    encoded_name = _unquote_imap_token(m.group("name"))
    name = imap_utf7_decode(encoded_name)
    return Folder(name=name, delimiter=delimiter, flags=flags)


# ------------------------------ configuration -------------------------------

@dataclass(frozen=True)
class EndpointConfig:
    server: str
    user: str
    password: str
    token: str
    thunderbird_profile: str
    thunderbird_primary_password: str


CONFIG_PATH = None

def ini_path() -> Path:
    return CONFIG_PATH or Path(__file__).resolve().parent / INI_FILE


def load_config(mode='migration'):
    path = ini_path()
    parser = configparser.RawConfigParser(interpolation=None)
    try:
        with path.open("r", encoding="utf-8") as f:
            parser.read_file(f)
    except FileNotFoundError:
        raise RuntimeError(f"Configuration file not found: {path}")
    except (OSError, configparser.Error) as exc:
        raise RuntimeError(f"Cannot read configuration file {path}: {exc}")

    def endpoint(section: str) -> EndpointConfig:
        if not parser.has_section(section):
            raise RuntimeError(f"{path}: missing [{section}] section")

        server = parser.get(section, "server", fallback="").strip()
        user = parser.get(section, "user", fallback="").strip()
        password = parser.get(section, "password", fallback="")
        token = parser.get(section, "token", fallback="")
        tb_profile = parser.get(
            section, "thunderbird_profile", fallback=""
        ).strip()
        tb_primary = parser.get(
            section, "thunderbird_primary_password", fallback=""
        )

        if not server:
            raise RuntimeError(f"{path}: [{section}] server is empty")
        if not user:
            raise RuntimeError(f"{path}: [{section}] user is empty")

        configured = sum(bool(x) for x in (password, token, tb_profile))
        if configured != 1:
            raise RuntimeError(
                f"{path}: [{section}] must set exactly ONE authentication "
                "source: password=, token=, or thunderbird_profile="
            )

        if tb_profile:
            expanded = Path(tb_profile).expanduser()
            if not expanded.is_absolute():
                expanded = (path.parent / expanded).resolve()
            tb_profile = str(expanded)

        return EndpointConfig(
            server=server,
            user=user,
            password=password,
            token=token,
            thunderbird_profile=tb_profile,
            thunderbird_primary_password=tb_primary,
        )

    src = endpoint("source") if mode != "import" else None
    dst = endpoint("destination") if mode != "export" else None

    if mode != "export" and not parser.has_section("migration"):
        raise RuntimeError(f"{path}: missing [migration] section")
    root = parser.get("migration", "destination_root", fallback="").strip() if parser.has_section("migration") else ""
    if mode != "export" and not root:
        raise RuntimeError(f"{path}: [migration] destination_root is empty")

    try:
        mode = path.stat().st_mode & 0o777
        if mode & 0o077:
            log(
                f"WARNING: {path} permissions are {mode:03o}; consider chmod 600.",
                file=sys.stderr,
            )
    except OSError:
        pass

    return src, dst, root


def auth_description(ep: EndpointConfig) -> str:
    if ep.thunderbird_profile:
        return "Thunderbird OAuth"
    if ep.token:
        return "XOAUTH2 access token"
    return "password"


# --------------------------- Thunderbird OAuth -------------------------------

class SECItem(ctypes.Structure):
    _fields_ = [
        ("type", ctypes.c_uint),
        ("data", ctypes.POINTER(ctypes.c_ubyte)),
        ("len", ctypes.c_uint),
    ]


class NSSLoginReader:
    """Read/decrypt Thunderbird Login Manager entries through system NSS."""

    def __init__(self, profile: str, primary_password: str = ""):
        self.profile = Path(profile)
        self.primary_password = primary_password
        self.nss = None
        self.slot = None

    def __enter__(self):
        if not self.profile.is_dir():
            raise RuntimeError(
                f"Thunderbird profile does not exist: {self.profile}"
            )
        for required in ("logins.json", "key4.db"):
            if not (self.profile / required).is_file():
                raise RuntimeError(
                    f"Thunderbird profile lacks {required}: {self.profile}"
                )

        lib = ctypes.util.find_library("nss3")
        if not lib:
            raise RuntimeError(
                "Cannot locate libnss3. Install NSS/Thunderbird libraries."
            )
        self.nss = ctypes.CDLL(lib)

        self.nss.NSS_Init.argtypes = [ctypes.c_char_p]
        self.nss.NSS_Init.restype = ctypes.c_int
        self.nss.NSS_Shutdown.argtypes = []
        self.nss.NSS_Shutdown.restype = ctypes.c_int

        self.nss.PK11_GetInternalKeySlot.argtypes = []
        self.nss.PK11_GetInternalKeySlot.restype = ctypes.c_void_p
        self.nss.PK11_FreeSlot.argtypes = [ctypes.c_void_p]
        self.nss.PK11_FreeSlot.restype = None
        self.nss.PK11_CheckUserPassword.argtypes = [
            ctypes.c_void_p, ctypes.c_char_p
        ]
        self.nss.PK11_CheckUserPassword.restype = ctypes.c_int

        self.nss.PK11SDR_Decrypt.argtypes = [
            ctypes.POINTER(SECItem),
            ctypes.POINTER(SECItem),
            ctypes.c_void_p,
        ]
        self.nss.PK11SDR_Decrypt.restype = ctypes.c_int

        self.nss.SECITEM_ZfreeItem.argtypes = [
            ctypes.POINTER(SECItem), ctypes.c_int
        ]
        self.nss.SECITEM_ZfreeItem.restype = None

        dbspec = ("sql:" + str(self.profile)).encode("utf-8")
        if self.nss.NSS_Init(dbspec) != 0:
            raise RuntimeError(
                f"NSS_Init failed for Thunderbird profile {self.profile}"
            )

        self.slot = self.nss.PK11_GetInternalKeySlot()
        if not self.slot:
            self.nss.NSS_Shutdown()
            raise RuntimeError("NSS could not get the internal key slot")

        pw = self.primary_password.encode("utf-8")
        if self.nss.PK11_CheckUserPassword(self.slot, pw) != 0:
            self.nss.PK11_FreeSlot(self.slot)
            self.slot = None
            self.nss.NSS_Shutdown()
            raise RuntimeError(
                "Thunderbird Primary Password is required or incorrect. "
                "Set thunderbird_primary_password= in the INI "
                "(leave it empty if Thunderbird has no Primary Password)."
            )

        return self

    def __exit__(self, exc_type, exc, tb):
        if self.slot:
            self.nss.PK11_FreeSlot(self.slot)
            self.slot = None
        if self.nss:
            self.nss.NSS_Shutdown()
            self.nss = None

    def decrypt(self, encoded: str) -> str:
        raw = base64.b64decode(encoded)
        inbuf = (ctypes.c_ubyte * len(raw)).from_buffer_copy(raw)
        inp = SECItem(
            type=0,
            data=ctypes.cast(inbuf, ctypes.POINTER(ctypes.c_ubyte)),
            len=len(raw),
        )
        out = SECItem()

        if self.nss.PK11SDR_Decrypt(
            ctypes.byref(inp), ctypes.byref(out), None
        ) != 0:
            raise RuntimeError("NSS failed to decrypt a Thunderbird login")

        try:
            plain = ctypes.string_at(out.data, out.len)
            return plain.decode("utf-8")
        finally:
            self.nss.SECITEM_ZfreeItem(ctypes.byref(out), 0)


def thunderbird_microsoft_refresh_token(ep: EndpointConfig) -> str:
    return thunderbird_refresh_token(ep, "Microsoft", TB_MS_OAUTH_ORIGIN, TB_MS_IMAP_SCOPE)


def thunderbird_refresh_token(ep: EndpointConfig, provider: str, origin: str, scope: str) -> str:
    profile = Path(ep.thunderbird_profile)
    logins_path = profile / "logins.json"

    try:
        obj = json.loads(logins_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Cannot read {logins_path}: {exc}")

    logins = obj.get("logins", [])
    if not isinstance(logins, list):
        raise RuntimeError(f"{logins_path}: malformed logins array")

    candidates = [
        x for x in logins
        if isinstance(x, dict)
        and x.get("hostname") == origin
        and scope in (x.get("httpRealm") or "").split()
    ]

    if not candidates:
        raise RuntimeError(
            f"No Thunderbird {provider} OAuth login with IMAP scope found in "
            f"{logins_path}. Make sure Thunderbird currently accesses this "
            "account using OAuth2."
        )

    with NSS_LOCK, NSSLoginReader(
        ep.thunderbird_profile, ep.thunderbird_primary_password
    ) as reader:
        matches = []
        for item in candidates:
            try:
                username = reader.decrypt(item["encryptedUsername"])
                refresh_token = reader.decrypt(item["encryptedPassword"])
            except KeyError:
                continue
            if username.casefold() == ep.user.casefold():
                matches.append((item, refresh_token))

    if not matches:
        raise RuntimeError(
            f"Thunderbird has {provider} OAuth entries, but none for "
            f"{ep.user!r} with the IMAP scope."
        )

    matches.sort(
        key=lambda pair: pair[0].get("timePasswordChanged", 0),
        reverse=True,
    )
    if len(matches) > 1:
        log(
            f"WARNING: Thunderbird has {len(matches)} matching {provider} OAuth "
            "entries; using the newest.",
            file=sys.stderr,
        )
    return matches[0][1]


def microsoft_access_token_from_refresh(refresh_token: str) -> str:
    return oauth_access_token_from_refresh(
        refresh_token, "Microsoft", TB_MS_TOKEN_ENDPOINT, TB_MS_CLIENT_ID)


def oauth_access_token_from_refresh(refresh_token: str, provider: str,
                                   endpoint: str, client_id: str,
                                   client_secret: str = "") -> str:
    # Mirror Thunderbird's refresh request: do not add a scope parameter.
    fields = {"client_id": client_id, "grant_type": "refresh_token",
              "refresh_token": refresh_token}
    if client_secret:
        fields["client_secret"] = client_secret
    body = urllib.parse.urlencode(fields).encode("ascii")

    req = urllib.request.Request(
        endpoint,
        data=body,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )

    refresh_started = time.monotonic()
    try:
        with urllib.request.urlopen(req, timeout=SOCKET_TIMEOUT) as resp:
            payload = resp.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")
        try:
            err = json.loads(detail)
            detail = (
                f"{err.get('error', 'OAuth error')}: "
                f"{err.get('error_description', detail)}"
            )
        except json.JSONDecodeError:
            pass
        raise RuntimeError(
            f"{provider} token refresh failed (HTTP {exc.code}): {detail}"
        )
    except urllib.error.URLError as exc:
        raise RuntimeError(f"{provider} token refresh failed: {exc}")

    try:
        result = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"{provider} token endpoint returned invalid JSON: {exc}"
        )

    if not isinstance(result, dict):
        raise RuntimeError(f"{provider} token endpoint returned a non-object response")
    token = result.get("access_token")
    if not isinstance(token, str) or not token:
        raise RuntimeError(
            f"{provider} token response contained no access_token: "
            + str(
                result.get("error_description")
                or result.get("error")
                or "Unspecified OAuth error"
            )
        )
    log(f'  {provider} access token received ({time.monotonic()-refresh_started:.2f}s)')
    return token


def thunderbird_access_token(ep: EndpointConfig) -> str:
    server = ep.server.lower()
    if server in ("imap.gmail.com", "imap.googlemail.com"):
        log(f"  reading Google refresh token from Thunderbird profile {ep.thunderbird_profile}")
        refresh = thunderbird_refresh_token(
            ep, "Google", TB_GOOGLE_OAUTH_ORIGIN, TB_GOOGLE_IMAP_SCOPE)
        log("  requesting a fresh Google IMAP access token")
        return oauth_access_token_from_refresh(
            refresh, "Google", TB_GOOGLE_TOKEN_ENDPOINT,
            TB_GOOGLE_CLIENT_ID, TB_GOOGLE_CLIENT_SECRET)
    if not (
        server == "outlook.office365.com"
        or server.endswith(".outlook.office365.com")
        or server.endswith(".outlook.com")
        or server.endswith(".hotmail.com")
    ):
        raise RuntimeError(
            "thunderbird_profile= supports Microsoft/Outlook and Gmail IMAP. "
            "Use password= or token= for other providers."
        )

    log(
        f"  reading Microsoft refresh token from Thunderbird profile "
        f"{ep.thunderbird_profile}"
    )
    refresh = thunderbird_microsoft_refresh_token(ep)
    log("  requesting a fresh Microsoft IMAP access token")
    return microsoft_access_token_from_refresh(refresh)


# ------------------------------ connection ----------------------------------

def connect_tls(server: str) -> imaplib.IMAP4_SSL:
    context = ssl.create_default_context()
    return imaplib.IMAP4_SSL(
        server, IMAP_PORT, ssl_context=context, timeout=SOCKET_TIMEOUT
    )


def connect_password(server: str, user: str, password: str) -> imaplib.IMAP4_SSL:
    c = connect_tls(server)
    try:
        typ, data = c.login(user, password)
        if typ != "OK":
            raise RuntimeError(f"Login failed for {server}: {data!r}")
        return c
    except Exception:
        try:
            c.logout()
        except Exception:
            pass
        raise


def connect_xoauth2(server: str, user: str, access_token: str) -> imaplib.IMAP4_SSL:
    c = connect_tls(server)
    auth_blob = (
        f"user={user}\x01"
        f"auth=Bearer {access_token}\x01\x01"
    ).encode("utf-8")
    try:
        typ, data = c.authenticate("XOAUTH2", lambda challenge: b"" if challenge else auth_blob)
        if typ != "OK":
            raise RuntimeError(f"XOAUTH2 authentication failed for {server}: {data!r}")
        return c
    except Exception:
        try:
            c.logout()
        except Exception:
            pass
        raise


def connect_endpoint(ep: EndpointConfig) -> imaplib.IMAP4_SSL:
    if ep.thunderbird_profile:
        access_token = thunderbird_access_token(ep)
        log(f'  Opening TLS connection and authenticating to {ep.server} with XOAUTH2')
        return connect_xoauth2(ep.server, ep.user, access_token)
    if ep.token:
        return connect_xoauth2(ep.server, ep.user, ep.token)
    return connect_password(ep.server, ep.user, ep.password)


# ------------------------- protocol and snapshots ----------------------------

TRANSIENT = (imaplib.IMAP4.abort, OSError, EOFError)
UID_RE = re.compile(rb'\bUID\s+(\d+)\b', re.I)
DATE_RE = re.compile(rb'\bINTERNALDATE\s+"([^"]+)"', re.I)
SIZE_RE = re.compile(rb'\bRFC822.SIZE\s+(\d+)\b', re.I)
START_RE = re.compile(rb'^\d+\s+\(')
FLAG_MAP = {k.lower(): v for k, v in STANDARD_FLAGS.items()}


def require_ok(result, operation):
    typ, data = result
    if typ != 'OK':
        raise RuntimeError(f'{operation}: {typ} {data!r}')
    return data or []


def date_epoch(value):
    # Month names are protocol English, independent of the host locale.
    m = re.fullmatch(r'\s*(\d{1,2})-([A-Za-z]{3})-(\d{4}) (\d\d):(\d\d):(\d\d) ([+-])(\d\d)(\d\d)', value)
    if not m:
        raise RuntimeError(f'Invalid INTERNALDATE: {value!r}')
    day, month, year, hour, minute, second, sign, zh, zm = m.groups()
    months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    try:
        offset = (int(zh)*60 + int(zm)) * (1 if sign == '+' else -1)
        if int(zh) > 23 or int(zm) > 59:
            raise ValueError('invalid zone')
        return int(_dt.datetime(int(year), months.index(month.lower())+1,
            int(day), int(hour), int(minute), int(second),
            tzinfo=_dt.timezone(_dt.timedelta(minutes=offset))).timestamp())
    except ValueError as exc:
        raise RuntimeError(f'Invalid INTERNALDATE: {value!r}') from exc


def hashes(raw):
    return (hashlib.sha256(raw).hexdigest(),
            hashlib.sha256(raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')).hexdigest())


def fetch_records(data, bodies=False, all_flags=False):
    """Join metadata before/after a literal; never search message bytes for UID."""
    groups = []
    meta, raw = b'', None
    for item in data or []:
        if item is None:
            continue
        head = item[0] if isinstance(item, tuple) else item
        if not isinstance(head, bytes):
            raise RuntimeError('Unsupported FETCH response')
        if START_RE.match(head) and meta:
            groups.append((meta, raw))
            meta, raw = b'', None
        meta += b' ' + head
        if isinstance(item, tuple):
            if raw is not None or len(item) != 2 or not isinstance(item[1], bytes):
                raise RuntimeError('Unexpected multiple FETCH literals')
            raw = item[1]
            # Validate wire framing against the literal's own declaration.
            # Exchange can report an approximate RFC822.SIZE that differs
            # from the MIME bytes returned by BODY.PEEK[].
            literal_size = re.search(rb'\{(\d+)\}$', head)
            if literal_size is None or len(raw) != int(literal_size.group(1)):
                raise RuntimeError('FETCH literal length differs from its declared size')
    if meta:
        groups.append((meta, raw))
    found = {}
    for meta, raw in groups:
        uid = UID_RE.search(meta)
        if not uid:
            # An unsolicited FLAGS update need not include UID.
            continue
        date, size = DATE_RE.search(meta), SIZE_RE.search(meta)
        if not date or not size or not re.search(rb'\bFLAGS\s+\(', meta, re.I):
            # Ignore unsolicited incomplete records; requested UIDs are checked below.
            continue
        if bodies and raw is None:
            continue
        d = date.group(1).decode('ascii')
        flag_match = re.search(rb'\bFLAGS\s+\(([^)]*)\)', meta, re.I)
        flags = sorted({FLAG_MAP[x.lower()] for x in flag_match.group(1).split()
                        if x.lower() in FLAG_MAP})
        rec = dict(uid=int(uid.group(1)), size=int(size.group(1)),
                   date=d, epoch=date_epoch(d), flags=flags)
        if all_flags:
            rec['original_flags'] = sorted({x.decode('ascii') for x in flag_match.group(1).split()
                                            if x.lower() != rb'\Recent'})
        if raw is not None:
            # Keep RFC822.SIZE as server metadata for snapshots/cache checks.
            # Hash and transfer the actual literal bytes, without rewriting.
            rec['strict'], rec['canon'] = hashes(raw)
            rec['semantic_version'] = SEMANTIC_VERSION
            rec['semantic'] = semantic_fingerprint(raw)
            rec['raw'] = raw
        found[rec['uid']] = rec
    return found


# Increment when comparison rules change; cached fingerprints are versioned.
SEMANTIC_VERSION = 1
ADDRESS_HEADERS = {'from', 'to', 'cc', 'bcc', 'sender', 'reply-to',
                   'resent-from', 'resent-to', 'resent-cc', 'resent-bcc', 'resent-sender'}
SIGNED_TYPES = {'multipart/signed', 'multipart/encrypted', 'application/pkcs7-mime',
                'application/x-pkcs7-mime', 'application/pkcs7-signature',
                'application/pgp-signature', 'application/pgp-encrypted'}


def semantic_fingerprint(raw):
    """Conservative equivalence, not authenticity or general MIME correctness.

    Unknown headers remain exact. Repeated field order and MIME child order are
    preserved. No transport headers are ignored. Signed mail is unclassified.
    """
    try:
        message = email.parser.BytesParser(policy=email.policy.default).parsebytes(raw)
        def walk(part):
            if part.defects or part.get_content_type() in SIGNED_TYPES:
                raise ValueError('malformed/signed MIME')
            names = [name.lower() for name, _ in part.raw_items()]
            if any(name in {'dkim-signature', 'domainkey-signature'} or name.startswith('arc-') for name in names):
                raise ValueError('signature-bearing message')
            headers = collections.defaultdict(list)
            raw_headers = list(part.raw_items())
            for name, value in raw_headers:
                name = name.lower()
                # These MIME fields are unique. Duplicate parsing would lose data.
                if name in {'content-type', 'content-disposition', 'content-transfer-encoding', 'mime-version'} and names.count(name) != 1:
                    raise ValueError('duplicate MIME field')
                normalized = value.replace('\r\n', '\n')
                if name in ADDRESS_HEADERS:
                    # Comments can contain meaningful user text; do not discard them.
                    if '(' in value or ')' in value:
                        raise ValueError('address comments not classified')
                    header = email.policy.default.header_factory(name, value)
                    if header.defects:
                        raise ValueError('malformed address')
                    normalized = [(group.display_name, [(address.display_name,
                        address.username, address.domain) for address in group.addresses])
                        for group in header.groups]
                elif name in {'content-type', 'content-disposition'}:
                    header = email.policy.default.header_factory(name, value)
                    if header.defects:
                        raise ValueError('malformed MIME parameters')
                    params = []
                    for key, val in header.params.items():
                        key = key.lower()
                        if name == 'content-type' and part.get_content_maintype() == 'multipart' and key == 'boundary':
                            continue  # Parsed tree/preamble/epilogue cover boundary framing.
                        if key == 'charset':
                            val = val.lower()
                        params.append((key, val))
                    kind = header.content_type if name == 'content-type' else header.content_disposition
                    normalized = (kind, sorted(params))
                elif name == 'content-transfer-encoding':
                    encoding = value.strip().lower()
                    if encoding not in {'7bit', '8bit', 'binary', 'base64', 'quoted-printable'}:
                        raise ValueError('unsupported transfer encoding')
                    continue  # Equivalent only if decoded payloads match below.
                elif name in {'date', 'resent-date', 'message-id', 'mime-version'}:
                    normalized = value.strip()
                headers[name].append(normalized)
            encoding = str(part.get('Content-Transfer-Encoding', '7bit')).strip().lower()
            if encoding not in {'7bit', '8bit', 'binary', 'base64', 'quoted-printable'}:
                raise ValueError('unsupported transfer encoding')
            if part.is_multipart():
                payload = part.get_payload()
                if not isinstance(payload, list):
                    raise ValueError('malformed MIME tree')
                children = [walk(child) for child in payload]
                body = ('children', children, part.preamble, part.epilogue)
            else:
                payload = part.get_payload(decode=True)
                if payload is None or part.defects:
                    raise ValueError('undecodable payload')
                if part.get_content_maintype() == 'text':
                    payload = payload.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
                body = ('payload', hashlib.sha256(payload).hexdigest())
            return (sorted(headers.items()), body)
        tree = walk(message)
        return hashlib.sha256(json.dumps(tree, ensure_ascii=True).encode('ascii')).hexdigest()
    except Exception:
        return None


def equivalent(a, b):
    return (a.get('semantic_version') == b.get('semantic_version') == SEMANTIC_VERSION
            and a.get('semantic') is not None and a['semantic'] == b.get('semantic'))


# Observations, not guarantees or allowlists. Add entries only with evidence.
PROVIDER_NOTES = [
    dict(provider='Google Gmail', hosts=('imap.gmail.com', 'imap.googlemail.com'),
         observed='2026-10-05',
         behavior='Very slow IMAP uploads observed: about 0.25-0.3 copied messages/second '
                  '(roughly one message every 3-4 seconds). Changing networks did not '
                  'resolve the reported slowness.',
         scope='User reports across two ISPs and ten VPN locations. Supplied import logs '
               'show 0.3 copied messages/second and approximately 25 KiB/second. '
               'Rates measure end-to-end migration work, including verification and other '
               'processing; they do not isolate APPEND latency or establish a fixed Gmail '
               'throttling rule. Results may differ by account, message sizes, workload '
               'and server conditions. This performance observation makes no byte-preservation claim.'),
    dict(provider='Microsoft Outlook/Hotmail', hosts=('outlook.office365.com',),
         observed='2026-10-04',
         behavior='Observed: trims address/date header whitespace; adds recipient angle brackets; '
                  'quotes charset parameters; changes Message-ID field capitalization; '
                  'reorders MIME-Version. RFC822.SIZE differed from the fetched literal.',
         scope='One supplied source/Outlook message diff plus migration logs. '
               'No guarantee for other messages/accounts. Exact byte preservation failed; '
               'the shown header changes are eligible for equivalence checks.'),
    dict(provider='GMX', hosts=('imap.gmx.com',), observed='2026-10-04',
         max_folder_levels=3,
         behavior='Does not allow more than 3 total folder levels (top level plus two subfolder levels). '
                  'The migration root counts as one level. A fourth-level CREATE was rejected.',
         scope='Observed at imap.gmx.com; consistent with GMX folder documentation at '
               'https://hilfe.gmx.net/premium/postfach/ordner-verwalten.html . '
               'Message byte-preservation behavior has not been established.'),
]


def show_provider_notes(server=None):
    if server is None:
        log('Provider | IMAP host | Observed | Behavior')
        for note in PROVIDER_NOTES:
            log(f"{note['provider']} | {', '.join(note['hosts'])} | {note['observed']} | {note['behavior']}")
            log('  Evidence/scope: ' + note['scope'])
        log('Unlisted providers: UNKNOWN; run a representative test and full verification.')
        return
    matches = [note for note in PROVIDER_NOTES if server.casefold() in note['hosts']]
    log(f'\nDestination preservation notice: {server}')
    if not matches:
        log('  No recorded observations for this host. Preservation behavior is UNKNOWN.')
        log('  Test representative mail and inspect full verification before deleting originals.')
    for note in matches:
        log('  ' + note['behavior'])
        log('  Evidence/scope: ' + note['scope'])
    log('  Provider notes never relax comparison rules or authorize duplicate matching.\n')


def check_provider_layout(server, names, delimiter):
    for note in PROVIDER_NOTES:
        limit = note.get('max_folder_levels')
        if server.casefold() not in note['hosts'] or limit is None:
            continue
        excessive = [(name, name.count(delimiter)+1) for name in names.values()
                     if name.count(delimiter)+1 > limit]
        if excessive:
            details = '; '.join(f'{name!r} ({levels} levels)' for name,levels in excessive)
            raise RuntimeError(f'{note["provider"]} allows at most {limit} total folder levels, '
                f'including the migration root. Unsupported mapped folders: {details}. '
                'No messages were uploaded/deleted in this run. Choose a destination supporting '
                'this hierarchy, or plan an explicit folder remapping; keep the existing journal/mail.')


def content_summary(raw):
    """Privacy-preserving diagnostics only; never used to accept/adopt mail."""
    try:
        parsed = email.parser.BytesParser(policy=email.policy.default).parsebytes(raw)
        headers = collections.defaultdict(list)
        for name, value in parsed.raw_items():
            # Hash unfolded values; report field names, never private values.
            unfolded = re.sub(r'\r?\n[ \t]+', ' ', value).strip()
            headers[name.lower()].append(hashlib.sha256(
                unfolded.encode('utf-8', 'surrogateescape')).hexdigest())
        parts = []
        defects = False
        for part in parsed.walk():
            defects = defects or bool(part.defects)
            if part.is_multipart():
                parts.append(('container', part.get_content_type()))
                continue
            payload = part.get_payload(decode=True)
            if payload is None:
                return None
            if part.get_content_maintype() == 'text':
                payload = payload.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
            parts.append((part.get_content_type(), part.get_content_charset(),
                          part.get_content_disposition(), part.get_filename(),
                          str(part.get('Content-ID', '')),
                          hashlib.sha256(payload).hexdigest()))
            defects = defects or bool(part.defects)
        if defects:
            return None
        return dict(headers=dict(headers),
                    mime=hashlib.sha256(json.dumps(parts, sort_keys=True).encode()).hexdigest())
    except Exception:
        # Diagnostics must never interrupt a confirmed transfer.
        return None


def show_content_difference(source_summary, raw, *, file_only=False, label='First changed message'):
    target = content_summary(raw)
    if source_summary is None or target is None:
        log('  Diagnostic: MIME comparison unavailable; inspect an original/returned message pair.',
            file_only=file_only)
        return
    sh, dh = source_summary['headers'], target['headers']
    added = sorted(set(dh)-set(sh))
    removed = sorted(set(sh)-set(dh))
    changed = sorted(name for name in sh.keys() & dh.keys() if sh[name] != dh[name])
    log('  '+label+': header fields added=' + repr(added) +
          ' removed=' + repr(removed) + ' changed=' + repr(changed), file_only=file_only)
    equal = source_summary['mime'] == target['mime']
    log('  Decoded MIME payload/structure: ' + ('MATCH' if equal else 'DIFFERENT') +
          ' (diagnostic only; exact content verification still reports the difference).',
          file_only=file_only)


class AppendRejected(RuntimeError):
    pass


class Session:
    """One owner at a time. Only read operations are automatically retried."""
    def __init__(self, ep, label, retries=4):
        self.ep, self.label, self.retries = ep, label, retries
        self.c = None
        self.mailbox = None
        self.validity = None
        self.last_activity = 0

    def close(self):
        if self.c:
            # Do not use CLOSE: even accidental read/write selection must not expunge.
            with contextlib.suppress(Exception):
                self.c.shutdown()
            self.c = None

    def connect(self):
        self.close()
        log(f'\nConnecting {self.label}: {self.ep.server} ({auth_description(self.ep)})', flush=True)
        self.c = connect_endpoint(self.ep)
        log(f'  {self.label}: IMAP authentication complete', flush=True)
        if self.mailbox is not None:
            old = self.validity
            self._select(self.mailbox)
            if old is not None and old != self.validity:
                raise RuntimeError(f'{self.label}: UIDVALIDITY changed after reconnect; restart to rescan')

    def read(self, operation, func):
        mailbox = f' in {self.mailbox!r}' if self.mailbox is not None else ''
        with OperationStatus(f'{self.label}: {operation}{mailbox}'):
            return self._read(operation, func)

    def _read(self, operation, func):
        for attempt in range(self.retries + 1):
            check_stop()
            try:
                if self.c is None:
                    self.connect()
                result = func(self.c)
                self.last_activity = time.monotonic()
                return result
            except TRANSIENT as exc:
                self.close()
                if attempt == self.retries:
                    raise RuntimeError(f'{self.label}: {operation} failed after retries: {exc}') from exc
                delay = min(2**attempt, 8)
                log(f'\n{self.label}: {operation}: {exc}; reconnecting in {delay}s', file=sys.stderr)
                for _ in range(delay*10):
                    check_stop()
                    time.sleep(.1)

    def _select(self, mailbox, readonly=True):
        require_ok(self.c.select(quote_mailbox(mailbox), readonly=readonly), f'{self.label}: EXAMINE {mailbox!r}')
        vals = self.c.response('UIDVALIDITY')[1]
        if not vals or not vals[0] or not vals[0].isdigit():
            raise RuntimeError(f'{self.label}: no usable UIDVALIDITY')
        self.validity = int(vals[0])
        sticky = self.c.response('UIDNOTSTICKY')[1]
        if sticky and sticky[0] is not None:
            raise RuntimeError(f'{self.label}: mailbox does not support persistent UIDs')
        self.mailbox = mailbox

    def select(self, mailbox):
        # Forget the previous mailbox before reconnecting to a different one.
        self.mailbox, self.validity = None, None
        self.read(f'EXAMINE {mailbox!r}', lambda c: self._select(mailbox))
        return self.validity

    def folders(self):
        def run(c):
            data = require_ok(c.list('', '*'), f'{self.label}: LIST')
            # LIST may return mailbox names as literals.
            lines, pos = [], 0
            while pos < len(data):
                item = data[pos]
                if isinstance(item, tuple):
                    head, literal = item
                    head = re.sub(rb'\{\d+\}$', b'', head).rstrip()
                    item = head + b' "' + literal.replace(b'\\', b'\\\\').replace(b'"', b'\\"') + b'"'
                    if pos+1 < len(data) and data[pos+1] == b'':
                        pos += 1
                if item:
                    lines.append(parse_list_line(item))
                pos += 1
            return lines
        return self.read('LIST', run)

    def snapshot(self):
        uids = self.read('UID SEARCH', lambda c: require_ok(c.uid('SEARCH', None, 'ALL'), 'SEARCH'))
        ids = [int(x) for x in (uids[0] or b'').split()] if uids else []
        records = {}
        for start in range(0, len(ids), 500):
            chunk = ids[start:start+500]
            with OperationStatus(f'{self.label}: reading metadata '
                    f'{start+1}-{start+len(chunk)}/{len(ids)} in {self.mailbox!r}'):
                records.update(self.fetch(chunk, False))
        if set(records) != set(ids):
            raise RuntimeError(f'{self.label}: mailbox changed during snapshot')
        return records

    def fetch(self, ids, bodies=True):
        fields = '(UID FLAGS INTERNALDATE RFC822.SIZE' + (' BODY.PEEK[]' if bodies else '') + ')'
        def run(c):
            data = require_ok(c.uid('FETCH', ','.join(map(str, ids)), fields), f'{self.label}: FETCH')
            recs = fetch_records(data, bodies)
            if any(uid not in recs for uid in ids):
                raise RuntimeError(f'{self.label}: FETCH omitted requested UID; mailbox may have changed')
            return {uid: recs[uid] for uid in ids}
        return self.read('UID FETCH', run)

    def append(self, mailbox, rec):
        with OperationStatus(f'{self.label}: APPEND to {mailbox!r} (source UID {rec["uid"]})',
                             report_completion=False):
            return self._append(mailbox, rec)

    def _append(self, mailbox, rec):
        # Consume any stale APPENDUID before this command.
        self.c.response('APPENDUID')
        self.c.literal = rec['raw']
        flags = '(' + ' '.join(rec['flags']) + ')'
        typ, data = self.c._simple_command('APPEND', quote_mailbox(mailbox), flags, '"'+rec['date']+'"')
        if typ != 'OK':
            raise AppendRejected(f'destination: APPEND rejected: {typ} {data!r}')
        self.last_activity = time.monotonic()
        vals = self.c.response('APPENDUID')[1]
        if vals and vals[0]:
            m = re.fullmatch(rb'(\d+) (\d+)', vals[0])
            if not m:
                raise RuntimeError('Invalid APPENDUID; pending upload must be reconciled')
            validity, uid = map(int, m.groups())
            if validity != self.validity:
                raise RuntimeError('APPENDUID UIDVALIDITY changed; pending upload must be reconciled')
            return uid
        return None

    def fetch_headers(self, ids):
        fields = '(UID FLAGS INTERNALDATE RFC822.SIZE BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])'
        def run(c):
            data = require_ok(c.uid('FETCH', ','.join(map(str, ids)), fields),
                              f'{self.label}: diagnostic header FETCH')
            records = fetch_records(data, bodies=False)
            if any(uid not in records or 'raw' not in records[uid] for uid in ids):
                raise RuntimeError(f'{self.label}: diagnostic headers omitted requested UID')
            return {uid: records[uid]['raw'] for uid in ids}
        return self.read('diagnostic header FETCH', run)


    def require_uid_expunge(self):
        if self.label != 'destination':
            raise RuntimeError('Selective deletion is restricted to the destination')
        data = self.read('CAPABILITY for repair', lambda c: require_ok(c.capability(), 'CAPABILITY'))
        caps = {token.upper() for line in data for token in line.split()}
        if not ({b'UIDPLUS', b'IMAP4REV2'} & caps):
            raise RuntimeError('Destination lacks selective UID EXPUNGE; no replacement/deletion attempted')

    def delete_uid(self, mailbox, validity, uid):
        self.require_uid_expunge()
        def command(c, operation):
            # Each retry reopens read/write and validates identity BEFORE writing.
            self._select(mailbox, readonly=False)
            if self.validity != validity:
                raise RuntimeError('Destination UIDVALIDITY changed before deletion; no UID deleted')
            return operation(c)
        self.read('mark exact repair UID deleted', lambda c: command(c, lambda conn:
            require_ok(conn.uid('STORE', str(uid), '+FLAGS.SILENT', r'(\Deleted)'), 'UID STORE')))
        self.read('expunge exact repair UID', lambda c: command(c, lambda conn:
            require_ok(conn.uid('EXPUNGE', str(uid)), 'UID EXPUNGE')))
        data = self.read('confirm repair deletion', lambda c: command(c, lambda conn:
            require_ok(conn.uid('SEARCH', None, 'UID', str(uid)), 'UID SEARCH')))
        if data and data[0] and str(uid).encode() in data[0].split():
            raise RuntimeError('Repair UID still exists after selective expunge; cleanup remains journaled')
        self.select(mailbox)  # Return to read-only mode; never use CLOSE/ordinary EXPUNGE.


def batches(records, ids, max_count, max_bytes):
    chunk, size = [], 0
    for uid in ids:
        n = records[uid]['size']
        if chunk and (len(chunk) >= max_count or size+n > max_bytes):
            yield chunk
            chunk, size = [], 0
        chunk.append(uid)
        size += n
    if chunk:
        yield chunk


class Prefetch:
    """Bounded producer; source connection remains owned by its worker."""
    def __init__(self, session, records, ids, opts, on_wait=None):
        self.s, self.records, self.ids, self.opts = session, records, ids, opts
        self.on_wait = on_wait
        self.q = queue.Queue(maxsize=1)
        self.cancel = threading.Event()
        self.finished = threading.Event()
        self.error = None
        self.thread = threading.Thread(target=self.run, daemon=True)

    def put(self, item):
        last_ping = time.monotonic()
        while not self.cancel.is_set() and not _stop_requested:
            try:
                self.q.put(item, timeout=.2)
                return True
            except queue.Full:
                if self.cancel.is_set() or _stop_requested:
                    return False
                if time.monotonic()-last_ping > 30:
                    self.s.read('NOOP', lambda c: require_ok(c.noop(), 'NOOP'))
                    last_ping = time.monotonic()
        return False

    def run(self):
        try:
            for ids in batches(self.records, self.ids, self.opts.batch_messages, self.opts.batch_bytes):
                if self.cancel.is_set() or _stop_requested:
                    return
                recs = self.s.fetch(ids)
                if not self.put([recs[uid] for uid in ids]):
                    return
        except BaseException as exc:
            # Publish failure outside the bounded data queue. Cancellation must
            # not try to queue an exception through the same failing path.
            self.error = exc
        finally:
            self.finished.set()

    def __enter__(self):
        self.thread.start()
        return self

    def __iter__(self):
        while True:
            check_stop()
            try:
                item = self.q.get(timeout=.5)
            except queue.Empty:
                if self.finished.is_set():
                    if self.error is not None:
                        raise self.error
                    return
                if self.on_wait is not None:
                    # Run only in the consumer, never in the connection worker.
                    self.on_wait()
                continue
            yield from item

    def __exit__(self, *args):
        self.cancel.set()
        # A FETCH is bounded by SOCKET_TIMEOUT; never reuse/close concurrently.
        self.thread.join(SOCKET_TIMEOUT + 10)
        if self.thread.is_alive():
            raise RuntimeError('Source worker did not stop within its timeout')


# ------------------------- durable journal -----------------------------------

class Journal:
    def __init__(self, path, identity):
        self.path = path
        self.lock = None
        self.db = None
        path.parent.mkdir(parents=True, exist_ok=True)
        self.lock = open(str(path)+'.lock', 'a+b')
        os.chmod(str(path)+'.lock', 0o600)
        try:
            if os.name == 'nt':
                import msvcrt
                self.lock.seek(0)
                self.lock.write(b'0')
                self.lock.flush()
                self.lock.seek(0)
                msvcrt.locking(self.lock.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            self.lock.close()
            raise RuntimeError('Another migration is using this journal') from exc
        try:
            self.db = sqlite3.connect(path)
            os.chmod(path, 0o600)
            self.db.execute('PRAGMA synchronous=FULL')
            self.db.executescript('''
                CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
                CREATE TABLE IF NOT EXISTS cache (
                    side TEXT, folder TEXT, validity INTEGER, uid INTEGER, record TEXT,
                    PRIMARY KEY(side, folder, validity, uid));
                CREATE TABLE IF NOT EXISTS bindings (
                    sf TEXT, sv INTEGER, su INTEGER, df TEXT, dv INTEGER, du INTEGER,
                    PRIMARY KEY(sf, sv, su), UNIQUE(df, dv, du));
                CREATE TABLE IF NOT EXISTS pending (id INTEGER PRIMARY KEY CHECK(id=1), record TEXT);
                CREATE TABLE IF NOT EXISTS repairs (sf TEXT, sv INTEGER, su INTEGER, record TEXT,
                    PRIMARY KEY(sf,sv,su));
            ''')
            encoded = json.dumps(identity, sort_keys=True)
            row = self.db.execute("SELECT value FROM settings WHERE key='identity'").fetchone()
            if row and row[0] != encoded:
                raise RuntimeError('Journal belongs to different accounts/root; choose another journal path')
            with self.db:
                self.db.execute("INSERT OR IGNORE INTO settings VALUES ('identity', ?)", (encoded,))
        except BaseException:
            self.close()
            raise

    def close(self):
        if self.db:
            self.db.close()
        if self.lock:
            self.lock.close()

    def cached(self, side, folder, validity):
        return {uid: json.loads(rec) for uid, rec in self.db.execute(
            'SELECT uid, record FROM cache WHERE side=? AND folder=? AND validity=?',
            (side, folder, validity))}

    def save(self, side, folder, validity, rec):
        record = {k:v for k,v in rec.items() if k != 'raw'}
        with self.db:
            self.db.execute('INSERT OR REPLACE INTO cache VALUES (?,?,?,?,?)',
                (side, folder, validity, rec['uid'], json.dumps(record)))

    def pending(self):
        row = self.db.execute('SELECT record FROM pending WHERE id=1').fetchone()
        return json.loads(row[0]) if row else None

    def begin(self, rec):
        if self.pending():
            raise RuntimeError('Unresolved upload already in journal')
        with self.db:
            self.db.execute('INSERT INTO pending VALUES (1,?)', (json.dumps(rec),))

    def reject(self):
        with self.db:
            self.db.execute('DELETE FROM pending')

    def accepted(self, uid):
        rec = self.pending()
        rec['accepted_uid'] = uid
        with self.db:
            self.db.execute('UPDATE pending SET record=? WHERE id=1', (json.dumps(rec),))

    def bind(self, sf, sv, su, df, dv, du, clear=False):
        with self.db:
            self.db.execute('INSERT INTO bindings VALUES (?,?,?,?,?,?) ON CONFLICT(sf,sv,su) DO UPDATE SET df=excluded.df,dv=excluded.dv,du=excluded.du',
                            (sf, sv, su, df, dv, du))
            if clear:
                self.db.execute('DELETE FROM pending')

    def record(self, side, folder, validity, uid):
        row = self.db.execute('SELECT record FROM cache WHERE side=? AND folder=? AND validity=? AND uid=?',
                              (side,folder,validity,uid)).fetchone()
        return json.loads(row[0]) if row else None

    def bindings(self, sf, sv, df, dv):
        return dict(self.db.execute('SELECT su,du FROM bindings WHERE sf=? AND sv=? AND df=? AND dv=?',
                                  (sf,sv,df,dv)))


    def repair_rows(self):
        return [json.loads(row[0]) for row in self.db.execute('SELECT record FROM repairs ORDER BY sf,sv,su')]

    def stage_repair(self, pending, uid):
        if uid <= pending['before'] or uid == pending.get('old_du'):
            raise RuntimeError('Repair APPENDUID is not a new destination UID; pending record retained')
        task = dict(pending, new_du=uid, state='candidate')
        with self.db:
            self.db.execute('INSERT INTO repairs VALUES (?,?,?,?)',
                (task['sf'],task['sv'],task['su'],json.dumps(task)))
            self.db.execute('DELETE FROM pending')

    def commit_repair(self, task):
        task = dict(task, state='cleanup')
        with self.db:
            self.db.execute('INSERT INTO bindings VALUES (?,?,?,?,?,?) '
                'ON CONFLICT(sf,sv,su) DO UPDATE SET df=excluded.df,dv=excluded.dv,du=excluded.du',
                tuple(task[key] for key in ('sf','sv','su','df','dv','new_du')))
            self.db.execute('UPDATE repairs SET record=? WHERE sf=? AND sv=? AND su=?',
                (json.dumps(task),task['sf'],task['sv'],task['su']))

    def finish_repair(self, task):
        with self.db:
            self.db.execute('DELETE FROM repairs WHERE sf=? AND sv=? AND su=?',
                (task['sf'],task['sv'],task['su']))

    def retire_missing_repair(self, task, surviving_old):
        # Update the binding and retire the vanished candidate atomically. No
        # destination STORE/EXPUNGE is authorized by this journal transition.
        with self.db:
            if surviving_old is not None:
                self.db.execute('INSERT INTO bindings VALUES (?,?,?,?,?,?) '
                    'ON CONFLICT(sf,sv,su) DO UPDATE SET df=excluded.df,dv=excluded.dv,du=excluded.du',
                    tuple(task[key] for key in ('sf','sv','su','df','dv'))+(surviving_old,))
            else:
                self.db.execute('DELETE FROM bindings WHERE sf=? AND sv=? AND su=?',
                    tuple(task[key] for key in ('sf','sv','su')))
            self.db.execute('DELETE FROM repairs WHERE sf=? AND sv=? AND su=?',
                tuple(task[key] for key in ('sf','sv','su')))
            self.db.execute('DELETE FROM cache WHERE side=? AND folder=? AND validity=? AND uid=?',
                ('destination',task['df'],task['dv'],task['new_du']))


@dataclass
class Options:
    journal: Path
    batch_messages: int = 25
    batch_bytes: int = 16*1024*1024
    retries: int = 4
    full_verify: bool = False


def options(args):
    p = configparser.RawConfigParser()
    p.read(ini_path(), encoding='utf-8')
    if not p.has_section('migration'):
        p.add_section('migration')
    section = p['migration']
    path = Path(section.get('journal', 'imap-migrator.sqlite3')).expanduser()
    if not path.is_absolute():
        path = ini_path().parent / path
    opts = Options(path.resolve(), section.getint('batch_messages', 25),
                   section.getint('batch_mib', 16)*1024*1024,
                   section.getint('retries', 4),
                   args.full_verify or section.getboolean('full_verify', False))
    if opts.batch_messages < 1 or opts.batch_messages > 500 or opts.batch_bytes < 1024*1024 or not 0 <= opts.retries <= 10:
        raise RuntimeError('Invalid batch/retry settings')
    return opts


def inventory(s, journal, side, folder, opts, full=False):
    validity = s.select(folder)
    records = s.snapshot()
    cached = journal.cached(side, folder, validity)
    missing = []
    for uid, rec in records.items():
        old = cached.get(uid)
        if not full and old and old['size'] == rec['size'] and old.get('semantic_version') == SEMANTIC_VERSION:
            rec.update(strict=old['strict'], canon=old['canon'],
                       semantic_version=old['semantic_version'], semantic=old.get('semantic'))
        else:
            missing.append(uid)
    if missing:
        label = 'local archive' if isinstance(s, ArchiveSource) else side
        log(f'  {label}: hashing {len(missing)} messages ({len(records)-len(missing)} cached)', flush=True)
        progress = HashProgress(len(missing), len(records)-len(missing), label)
        progress.show(True)
        with Prefetch(s, records, missing, opts, on_wait=progress.show) as stream:
            for rec in stream:
                journal.save(side, folder, validity, rec)
                records[rec['uid']] = {k:v for k,v in rec.items() if k != 'raw'}
                progress.hashed_message(len(rec['raw']))
        progress.show(True)
    return validity, records


def reconcile(dst, j, opts, *, retry_missing=False, source=None, resolve_uid=None):
    p = j.pending()
    if not p:
        if retry_missing or resolve_uid is not None:
            log('No pending APPEND to resolve.')
        return
    log(f"\nReconciling pending APPEND: {p['sf']!r} UID {p['su']}", flush=True)
    dv, records = inventory(dst, j, 'destination', p['df'], opts)
    if dv != p['dv']:
        raise RuntimeError('Pending APPEND has a different destination UIDVALIDITY; manual review required')
    candidates = {uid:r for uid,r in records.items() if uid > p['before']}
    assigned = p.get('accepted_uid')
    exact = [uid for uid,r in candidates.items() if r['canon'] == p['canon']]
    original = j.record('source', p['sf'], p['sv'], p['su'])
    comparable = original is not None and original['canon'] == p['canon']
    semantic = [uid for uid,r in candidates.items() if comparable and equivalent(original, r)]
    matches = ([assigned] if assigned in records else []) if assigned else exact
    if not assigned and len(matches) == 0 and len(semantic) == 1:
        matches = semantic
        log('  Recovered by one checked-equivalent candidate; original bytes differ.')
    log(f'  Destination UIDVALIDITY={dv}; before UID={p["before"]}; '
          f'new UID count={len(candidates)}; exact matches={exact}; '
          f'checked-equivalent matches={semantic}; acknowledged UID={assigned}', flush=True)
    if resolve_uid is not None:
        if resolve_uid not in candidates and resolve_uid != assigned:
            raise RuntimeError('--resolve-pending-uid must identify a new or acknowledged UID in this folder')
        if resolve_uid not in records:
            raise RuntimeError('Requested pending destination UID does not exist')
        owner = j.db.execute('SELECT sf,sv,su FROM bindings WHERE df=? AND dv=? AND du=?',
                             (p['df'],dv,resolve_uid)).fetchone()
        if owner is not None and tuple(owner) != (p['sf'],p['sv'],p['su']):
            raise RuntimeError('Requested UID is already mapped to another source message')
        if any(task['new_du'] == resolve_uid and task['df'] == p['df'] and task['dv'] == dv
               for task in j.repair_rows()):
            raise RuntimeError('Requested UID already belongs to an outstanding repair')
        log(f'  Explicitly accepting destination UID {resolve_uid} as the pending upload. '
              'Full verification will still report content/metadata differences.', flush=True)
        matches = [resolve_uid]
    if len(matches) != 1 and retry_missing:
        if candidates:
            raise RuntimeError('Pending upload has new destination UIDs; retry refused to prevent duplicates. '
                'Inspect their raw messages, then use --resolve-pending-uid UID only if you establish identity.')
        if source is None:
            raise RuntimeError('Pending retry requires the source connection')
        # Reconnect and inventory again before an explicitly requested retry.
        # This reduces uncertainty but cannot rule out a delayed server commit.
        dst.close()
        fresh_validity = dst.select(p['df'])
        fresh = dst.snapshot()
        if fresh_validity != p['dv'] or any(uid > p['before'] for uid in fresh):
            raise RuntimeError('Destination changed during pending retry checks; no retry sent')
        if source.select(p['sf']) != p['sv']:
            raise RuntimeError('Pending source UIDVALIDITY changed; no retry sent')
        rec = source.fetch([p['su']])[p['su']]
        if rec['canon'] != p['canon']:
            raise RuntimeError('Pending source content changed; no retry sent')
        if p.get('source_strict') and rec['strict'] != p['source_strict']:
            raise RuntimeError('Pending repair source bytes changed; no retry sent')
        j.save('source', p['sf'], p['sv'], rec)
        p.pop('accepted_uid', None)
        p['retry_count'] = p.get('retry_count', 0) + 1
        # Update in place: never discard the uncertain upload record before sending.
        with j.db:
            j.db.execute('UPDATE pending SET record=? WHERE id=1', (json.dumps(p),))
        log('  Explicit pending retry: no new destination messages found after reconnect. '
              'A delayed old upload could still cause a duplicate; inspect full verification.', flush=True)
        try:
            du = dst.append(p['df'], rec)
        except AppendRejected:
            j.reject()
            raise
        except TRANSIENT as exc:
            dst.close()
            raise RuntimeError(f'Pending retry lost its reply: {exc}. Record retained; no further retry sent.') from exc
        if du is not None:
            j.accepted(du)
        # Only one retry per explicit invocation, even when its result is uncertain.
        return reconcile(dst, j, opts)
    if len(matches) != 1:
        preview = list(sorted(candidates))[:20]
        raise RuntimeError(f'Pending APPEND remains unresolved: {len(matches)} identity matches; '
            f'{len(candidates)} new destination messages; candidate UIDs={preview}. '
            'Pending record retained; no retry sent. '
            + ('If you accept the residual risk of a delayed duplicate, rerun with --retry-pending. '
               'This option retries only when no new destination UIDs exist.' if not candidates else
               'Inspect candidates and use --resolve-pending-uid UID only after identifying the uploaded message.'))
    du = matches[0]
    if p.get('kind') == 'repair':
        j.stage_repair(p, du)
    else:
        j.bind(p['sf'], p['sv'], p['su'], p['df'], dv, du, clear=True)
    log(f'  Recovered destination UID {du}')


def mapping(folders, dest_folders, root):
    delims = {f.delimiter for f in dest_folders if f.delimiter is not None}
    if len(delims) != 1:
        raise RuntimeError('Destination has a flat or ambiguous namespace; cannot safely construct hierarchy')
    delim = next(iter(delims))
    if any(not x for x in root.split(delim)):
        raise RuntimeError('Destination root contains empty path components')
    result, seen = {}, {}
    for f in folders:
        parts = f.name.split(f.delimiter) if f.delimiter else [f.name]
        if any(delim in x or not x for x in parts):
            raise RuntimeError(f'Cannot map {f.name!r} losslessly using destination delimiter {delim!r}')
        name = delim.join([root] + parts)
        # Conservative: some servers treat names case-insensitively.
        key = name.casefold()
        if key in seen:
            raise RuntimeError(f'Folder mapping collision: {seen[key]!r} and {f.name!r}')
        result[f.name], seen[key] = name, f.name
    return delim, result


def ensure_folder(dst, name, delim, known=None):
    if known is None:
        known = {f.name for f in dst.folders()}
    parts = name.split(delim)
    for n in range(1, len(parts)+1):
        prefix = delim.join(parts[:n])
        if prefix not in known:
            log(f'  Creating destination folder {prefix!r}', flush=True)
            # CREATE is idempotent only after checking LIST following an uncertain result.
            try:
                data = dst.read('prepare CREATE', lambda c: c)
                typ, response = data.create(quote_mailbox(prefix))
                if typ != 'OK' and prefix not in {f.name for f in dst.folders()}:
                    raise RuntimeError(f'CREATE {prefix!r} failed: {response!r}')
            except TRANSIENT:
                dst.close()
                if prefix not in {f.name for f in dst.folders()}:
                    raise RuntimeError(f'CREATE outcome unknown for {prefix!r}; rerun')
            known.add(prefix)


def prepare_destination_tree(dst, names, delim, source=None, known_folders=None):
    log(f'\nPreparing destination tree: {len(names)} message folders', flush=True)
    started = time.monotonic()
    known = {f.name for f in (dst.folders() if known_folders is None else known_folders)}
    def keep_source_alive():
        if source is not None and time.monotonic()-source.last_activity > 30:
            log('  Keeping source connection alive (NOOP)', flush=True)
            source.read('preparation keepalive NOOP', lambda c: require_ok(c.noop(), 'source NOOP'))
    ordered = sorted(names.items(), key=lambda item: (item[1].count(delim), item[1]))
    for n, (sf, df) in enumerate(ordered, 1):
        check_stop()
        keep_source_alive()
        folder_started = time.monotonic()
        log(f'  Preparing [{n}/{len(ordered)}] {sf!r} -> {df!r}', flush=True)
        try:
            ensure_folder(dst, df, delim, known)
            # EXAMINE alone also succeeds on read-only mailboxes. SELECT checks
            # writable selection without changing messages; return to EXAMINE.
            keep_source_alive()
            log(f'  Checking writable SELECT: {df!r}', flush=True)
            dst.mailbox, dst.validity = None, None
            def writable(c):
                dst._select(df, readonly=False)
                readonly = c.response('READ-ONLY')[1]
                if readonly and any(value is not None for value in readonly):
                    raise RuntimeError(f'Destination selected {df!r} read-only')
            dst.read('check writable SELECT', writable)
            log(f'  Returning to read-only EXAMINE: {df!r}', flush=True)
            dst.select(df)
            log(f'  Ready [{n}/{len(ordered)}]: {df!r} '
                f'({time.monotonic()-folder_started:.2f}s)', flush=True)
        except (RuntimeError, imaplib.IMAP4.error, OSError) as exc:
            raise RuntimeError(f'Destination tree preparation failed: {sf!r} -> {df!r}: {exc}. '
                'No messages were uploaded/deleted in this run. Successfully created folders '
                'are retained for the next run; keep the journal.') from exc

    # Subscribe only after every message folder has passed preparation. Include
    # intermediate containers so clients can display the complete hierarchy.
    paths = set()
    for df in names.values():
        parts = df.split(delim)
        paths.update(delim.join(parts[:n]) for n in range(1, len(parts)+1))
    subscriptions = sorted(paths, key=lambda value: (value.count(delim), value))
    log(f'Subscribing {len(subscriptions)} destination folders...', flush=True)
    for n, name in enumerate(subscriptions, 1):
        keep_source_alive()
        log(f'  Subscribing [{n}/{len(subscriptions)}] {name!r}', flush=True)
        try:
            dst.read('SUBSCRIBE', lambda c: require_ok(
                c.subscribe(quote_mailbox(name)), f'SUBSCRIBE {name!r}'))
        except (RuntimeError, imaplib.IMAP4.error, OSError) as exc:
            log(f'WARNING: could not subscribe destination folder {name!r}: {exc}. '
                  'Subscribe manually in your mail client if it is hidden.', file=sys.stderr)
    keep_source_alive()
    log(f'Destination tree ready ({time.monotonic()-started:.2f}s); '
        'beginning migration/recovery.', flush=True)


class Progress:
    def __init__(self, total):
        self.total, self.done, self.copied, self.skipped = total, 0, 0, 0
        self.folder_total = self.folder_done = self.folder_copied = self.folder_mapped = 0
        self.uploaded, self.downloaded = 0, 0
        self.start, self.last = time.monotonic(), 0
        self.log_last = 0

    def begin_folder(self, total, mapped, destination_count, name='Folder'):
        self.folder_total, self.folder_done = total, mapped
        self.folder_copied, self.folder_mapped = 0, mapped
        self.done += mapped
        self.skipped += mapped
        if total == 0:
            log('  Empty source folder; destination empty. Nothing to copy.')
            return
        log(f'  Folder: {total} source messages; destination currently contains '
            f'{destination_count}; {mapped} already mapped; {total-mapped} to copy')
        self.show(True)

    def copied_message(self, size):
        self.copied += 1
        self.done += 1
        self.folder_copied += 1
        self.folder_done += 1
        self.uploaded += size

    def end_folder(self):
        pass

    def elapsed_seconds(self, now):
        return now-self.start

    def show(self, force=False):
        global _progress_rows, _progress_columns
        if self.folder_total == 0:
            return
        now = time.monotonic()
        live = sys.stdout.isatty() and os.environ.get('TERM') != 'dumb'
        if not force and now-self.last < (.25 if live else 10):
            return
        elapsed = max(self.elapsed_seconds(now), .001)
        lines = self.progress_lines(elapsed)
        with OUTPUT_LOCK:
            if LOG_FILE is not None and (force or now-self.log_last >= 10):
                log('\n'.join(lines), file_only=True)
                self.log_last = now
            columns = shutil.get_terminal_size().columns
            if live and _progress_rows and columns == _progress_columns:
                sys.stdout.write(f'\x1b[{_progress_rows}A\r\x1b[J')
            sys.stdout.write('\n'.join(lines)+'\n')
            sys.stdout.flush()
            _progress_rows = sum(max(1,(len(line)+columns-1)//columns) for line in lines) if live else 0
            _progress_columns = columns
        self.last = now

    def progress_lines(self, elapsed):
        return [
            f'  Folder: {self.folder_done}/{self.folder_total} processed | '
            f'{max(0,self.folder_total-self.folder_done)} left | copied={self.folder_copied} | '
            f'already mapped={self.folder_mapped}',
            f'  Overall: {self.done}/{self.total} processed | '
            f'{max(0,self.total-self.done)} left | copied={self.copied} | already mapped={self.skipped}',
            f'  Speed: {self.copied/elapsed:.1f} copied msg/s | uploaded={human_bytes(self.uploaded)} '
            f'({human_bytes(int(self.uploaded/elapsed))}/s) | fetched={human_bytes(self.downloaded)} | '
            f'copy elapsed={human_duration(elapsed)}']


def terminal_text_width(text):
    return sum(0 if unicodedata.combining(char) else
               2 if unicodedata.east_asian_width(char) in ('W', 'F') else 1
               for char in text)


class CopyProgress(Progress):
    """Compact copy-only counters; widths are chosen from the source census."""
    def __init__(self, total, folder_counts):
        super().__init__(total)
        self.names = {name: ''.join(char if char.isprintable() else repr(char)[1:-1]
                                   for char in name) for name in folder_counts}
        self.name_width = max([terminal_text_width('Total')] +
                              [terminal_text_width(name) for name in self.names.values()])
        self.number_width = max(len(f'{total:,}'),
                                max((len(f'{count:,}') for count in folder_counts.values()), default=1))
        self.folder_name = ''
        self.copy_seconds = 0.0
        self.copy_started = None
        self.show_stats = False

    def row(self, name, total, copied, left):
        padding = ' ' * max(0, self.name_width-terminal_text_width(name))
        count = f'({total:,})'
        return (f'{name}{padding} {count:>{self.number_width+2}}: '
                f'{copied:>{self.number_width},} | {left:>{self.number_width},} left')

    def begin_folder(self, total, mapped, destination_count, name='Folder'):
        self.folder_name = self.names.get(name, name)
        self.folder_total, self.folder_done = total, mapped
        self.folder_copied, self.folder_mapped = 0, mapped
        self.done += mapped
        self.skipped += mapped
        self.show_stats = total > mapped
        if not self.show_stats:
            log('  Nothing to upload; verification follows.')
            return
        self.copy_started = time.monotonic()
        self.show(True)

    def end_folder(self):
        if self.copy_started is not None:
            self.copy_seconds += time.monotonic()-self.copy_started
            self.copy_started = None

    def elapsed_seconds(self, now):
        return self.copy_seconds + (now-self.copy_started if self.copy_started is not None else 0)

    def show(self, force=False):
        if self.show_stats:
            super().show(force)

    def progress_lines(self, elapsed):
        return [
            self.row(self.folder_name, self.folder_total, self.folder_copied,
                     max(0, self.folder_total-self.folder_done)),
            self.row('Total', self.total, self.copied, max(0, self.total-self.done)),
            '',
            f'Transferred: {human_bytes(self.uploaded)} ↑ | {human_bytes(self.downloaded)} ↓',
            f'Speed:       {self.copied/elapsed:.1f} msg/s | '
            f'{human_bytes(int(self.uploaded/elapsed))}/s ↑',
            f'Copy time:   {human_duration(elapsed) if elapsed >= 1 else f"{elapsed:.2f}s"}']


OPERATION_STATE = threading.local()


class OperationStatus(Progress):
    """Display-only timer: never accesses a connection or changes retry behavior."""
    def __init__(self, description, report_completion=True):
        super().__init__(1)
        self.folder_total = 1
        self.description = description
        self.report_completion = report_completion
        self.finished = threading.Event()
        self.thread = None
        self.shown = False
        self.enabled = False

    def __enter__(self):
        # Prefetch already has its own consumer progress. Nested reads belong
        # to the outer operation, so only one timer can redraw the terminal.
        self.enabled = (threading.current_thread() is threading.main_thread()
                        and not getattr(OPERATION_STATE, 'active', False))
        if self.enabled:
            OPERATION_STATE.active = True
            self.thread = threading.Thread(target=self.run, daemon=True)
            try:
                self.thread.start()
            except BaseException:
                OPERATION_STATE.active = False
                raise
        return self

    def run(self):
        # Fast operations need no extra output. Event.wait permits instant exit.
        if self.finished.wait(2):
            return
        while not self.finished.is_set():
            self.shown = True
            self.show()
            if self.finished.wait(.5):
                return

    def progress_lines(self, elapsed):
        live = sys.stdout.isatty() and os.environ.get('TERM') != 'dumb'
        frame = '-\\|/'[int(elapsed*2) % 4]+' ' if live else ''
        return [f'  {frame}Waiting: {self.description} | operation elapsed={elapsed:.2f}s']

    def __exit__(self, exc_type, exc, tb):
        if self.enabled:
            self.finished.set()
            self.thread.join()
            OPERATION_STATE.active = False
            if self.shown and self.report_completion:
                result = 'completed' if exc_type is None else 'stopped'
                log(f'  {self.description}: {result} '
                    f'({time.monotonic()-self.start:.2f}s)')


class HashProgress(Progress):
    """Folder inventory progress, separate from copied-message counters."""
    def __init__(self, total, cached, label):
        super().__init__(total)
        self.folder_total = total
        self.cached, self.label = cached, label

    def hashed_message(self, size):
        self.done += 1
        self.downloaded += size
        self.show()

    def progress_lines(self, elapsed):
        # Bytes/counts advance only after a fetched batch has been parsed and
        # each record journaled. Elapsed time also advances while awaiting it.
        return [
            f'  {self.label}: {self.done}/{self.total} hashed | '
            f'{max(0,self.total-self.done)} left | {self.cached} cached',
            f'  Read={human_bytes(self.downloaded)} '
            f'({human_bytes(int(self.downloaded/elapsed))}/s) | '
            f'hashing elapsed={elapsed:.2f}s']


class ArchiveProgress(Progress):
    def __init__(self, total, operation='export'):
        super().__init__(total)
        self.operation = operation

    def begin_archive_folder(self, total):
        self.folder_total, self.folder_done = total, 0
        self.folder_copied = self.folder_mapped = 0
        if not total:
            log('  Empty folder; no message files to '+self.operation+'.')
            return
        log(f'  Folder: {total} messages to {self.operation}')
        self.show(True)

    def reused_message(self):
        self.done += 1
        self.skipped += 1
        self.folder_done += 1
        self.folder_mapped += 1
        self.show()

    def completed_message(self, size):
        self.copied_message(size)
        self.downloaded += size
        self.show()

    def progress_lines(self, elapsed):
        if self.operation == 'validate':
            return [
                f'  Folder: {self.folder_done}/{self.folder_total} checked | '
                f'{max(0,self.folder_total-self.folder_done)} left',
                f'  Overall: {self.done}/{self.total} checked | {max(0,self.total-self.done)} left',
                f'  Speed: {self.copied/elapsed:.1f} checked msg/s | read={human_bytes(self.downloaded)} '
                f'({human_bytes(int(self.downloaded/elapsed))}/s) | check elapsed={human_duration(elapsed)}']
        return [
            f'  Folder: {self.folder_done}/{self.folder_total} processed | '
            f'{max(0,self.folder_total-self.folder_done)} left | exported={self.folder_copied} | '
            f'already archived={self.folder_mapped}',
            f'  Overall: {self.done}/{self.total} processed | '
            f'{max(0,self.total-self.done)} left | exported={self.copied} | already archived={self.skipped}',
            f'  Speed: {self.copied/elapsed:.1f} exported msg/s | fetched={human_bytes(self.downloaded)} '
            f'({human_bytes(int(self.downloaded/elapsed))}/s) | export elapsed={human_duration(elapsed)}']


def human_bytes(n):
    value = float(n)
    for unit in ('B','KiB','MiB','GiB','TiB'):
        if value < 1024 or unit == 'TiB':
            return f'{value:.1f} {unit}'
        value /= 1024


MESSAGE_DIAGNOSTIC_LIMIT = 5  # Per folder/phase; provider rewriting may affect every message.


def message_header_bytes(raw):
    # Retain bounded source headers for upload diagnostics, not whole bodies.
    limit = min(len(raw), 64*1024)
    ends = [raw.find(separator, 0, limit) for separator in (b'\r\n\r\n', b'\n\n', b'\r\r')]
    end = min((pos for pos in ends if pos >= 0), default=limit)
    return raw[:end], end == 64*1024


def message_log_headers(raw):
    try:
        message = email.parser.BytesHeaderParser(policy=email.policy.default).parsebytes(raw)
        # repr() at output prevents header control characters corrupting the terminal.
        return [(name, str(value)) for name in ('Subject', 'Date')
                for value in message.get_all(name, [])]
    except Exception:
        return None


def show_message_identity(src, sf, sv, su, df, dv, du, reason, headers=None, file_only=False):
    log(f'  {reason}: source folder={sf!r} UIDVALIDITY={sv} UID={su}; '
        f'destination folder={df!r} UIDVALIDITY={dv} UID={du}', file_only=file_only)
    if isinstance(src, ArchiveSource):
        eml, _ = archive_message_paths(src.root, src.entries[sf], su)
        log(f'    Archived EML: {str(eml.absolute())!r}', file_only=file_only)
    if headers is None:
        log('    Source Subject/Date unavailable.', file_only=file_only)
    elif not headers:
        log('    Source has no Subject or Date header.', file_only=file_only)
    else:
        for name, value in headers:
            log(f'    Source {name}: {value!r}', file_only=file_only)


def migrate_folder(src, dst, j, f, df, delim, opts, progress):
    ensure_folder(dst, df, delim)
    dv, dest = inventory(dst, j, 'destination', df, opts)
    sv = src.select(f.name)
    source = src.snapshot()
    if not source and dest:
        log(f'  Empty source folder; destination contains {len(dest)} messages. '
            'Checking reconciliation; no destination messages will be deleted.')
    bound = j.bindings(f.name, sv, df, dv)
    # A missing mapped UID is an ordinary incomplete destination, not permission
    # to delete anything. Forget only that absent binding and upload the source again.
    missing = {su:du for su,du in bound.items() if su in source and du not in dest}
    if missing:
        log(f'  Restoring {len(missing)} missing journaled message(s); no destination deletion.', flush=True)
        with j.db:
            for su,du in missing.items():
                j.db.execute('DELETE FROM bindings WHERE sf=? AND sv=? AND su=? AND df=? AND dv=? AND du=?',
                             (f.name,sv,su,df,dv,du))
        bound = {su:du for su,du in bound.items() if su not in missing}
    bound = {su:du for su,du in bound.items() if su in source and du in dest}
    cached = j.cached('source', f.name, sv)
    for uid,r in source.items():
        old = cached.get(uid)
        if old and old['size'] == r['size'] and old.get('semantic_version') == SEMANTIC_VERSION:
            r.update(strict=old['strict'], canon=old['canon'],
                     semantic_version=old['semantic_version'], semantic=old.get('semantic'))
    unmatched = set(dest) - set(bound.values())
    if unmatched:
        # Preflight the complete folder before any upload. This also retains multiplicity.
        _, source = inventory(src, j, 'source', f.name, opts)
        available = collections.defaultdict(list)
        for uid in sorted(unmatched):
            available[dest[uid]['canon']].append(uid)
        adopted = []
        for uid in sorted(source):
            if uid in bound:
                continue
            matches = available[source[uid]['canon']]
            if matches:
                du = matches.pop(0)
                adopted.append((uid,du))
        remaining = sum(map(len, available.values()))
        if remaining:
            raise RuntimeError(f'{df!r}: {remaining} destination message(s) cannot be matched to source '
                'by content/multiplicity. No messages appended to this folder. '
                'Use a new destination_root for a clean migration, or review existing mail manually.')
        for su,du in adopted:
            j.bind(f.name, sv, su, df, dv, du)
            bound[su] = du
    todo = [uid for uid in sorted(source) if uid not in bound]
    progress.begin_folder(len(source), len(bound), len(dest), name=f.name)
    waiting = {}
    waiting_bytes = 0
    rewritten = 0
    diagnosed = False
    diagnostic_count = 0
    diagnostic_suppressed = False
    def check_uploaded():
        nonlocal waiting_bytes, rewritten, diagnosed, diagnostic_count, diagnostic_suppressed
        if not waiting:
            return
        fetched = dst.fetch(list(waiting))
        failures = []
        equivalent_count = 0
        for du,got in fetched.items():
            j.save('destination', df, dv, got)
            dest[du] = {k:v for k,v in got.items() if k != 'raw'}
            expected, summary = waiting[du]
            if expected['canon'] != got['canon']:
                failures.append(du)
                equivalent_count += equivalent(expected, got)
                if diagnostic_count < MESSAGE_DIAGNOSTIC_LIMIT or LOG_FILE is not None:
                    header_bytes, truncated = expected['log_header_bytes']
                    show_message_identity(src, f.name, sv, expected['uid'], df, dv, du,
                        'Accepted message changed bytes', message_log_headers(header_bytes),
                        file_only=diagnostic_count >= MESSAGE_DIAGNOSTIC_LIMIT)
                    if truncated:
                        log('    Source header diagnostic limited to the first 64 KiB.',
                            file_only=diagnostic_count >= MESSAGE_DIAGNOSTIC_LIMIT)
                if diagnostic_count >= MESSAGE_DIAGNOSTIC_LIMIT and not diagnostic_suppressed:
                    log(f'  Further changed-message details suppressed for {df!r} '
                        f'(limit {MESSAGE_DIAGNOSTIC_LIMIT}); counts and journal mappings are retained.')
                    diagnostic_suppressed = True
                diagnostic_count += 1
                if not diagnosed and summary is not None:
                    show_content_difference(summary, got['raw'])
                    diagnosed = True
        waiting.clear()
        waiting_bytes = 0
        if failures:
            rewritten += len(failures)
            log(f'\nNOTICE: {df!r}: {len(failures)} accepted message(s) changed bytes '
                  f'({rewritten} this pass): equivalent-under-checked-rules={equivalent_count}, '
                  f'changed-or-uncertain={len(failures)-equivalent_count}. '
                  'Continuing transfer; original bytes remain different.')
    if todo:
        with Prefetch(src, source, todo, opts, on_wait=progress.show) as stream:
            for rec in stream:
                check_stop()
                uid = rec['uid']
                progress.downloaded += len(rec['raw'])
                j.save('source', f.name, sv, rec)
                source[uid] = {k:v for k,v in rec.items() if k != 'raw'}
                if time.monotonic()-dst.last_activity > 30:
                    dst.read('pre-upload NOOP', lambda c: require_ok(c.noop(), 'NOOP'))
                p = dict(sf=f.name, sv=sv, su=uid, df=df, dv=dv,
                         before=max(dest, default=0), canon=rec['canon'])
                j.begin(p)  # FULL synchronous commit BEFORE handing the literal to IMAP.
                try:
                    du = dst.append(df, rec)
                except AppendRejected:
                    j.reject()  # Tagged rejection definitively reports that APPEND did not succeed.
                    raise
                except TRANSIENT as exc:
                    log(f'\nDestination APPEND connection lost: {exc}; reconciling without retry', file=sys.stderr)
                    dst.close()
                    reconcile(dst, j, opts)
                    du = j.bindings(f.name, sv, df, dv)[uid]
                else:
                    if du is not None:
                        j.accepted(du)  # Recovery by UID remains safe even if content is rewritten.
                    if du is None:
                        reconcile(dst, j, opts)
                        du = j.bindings(f.name, sv, df, dv)[uid]
                    else:
                        # Durable UID binding is enough for safe resume. Verify bodies in
                        # batches, avoiding a second round trip after every APPEND.
                        j.bind(f.name, sv, uid, df, dv, du, clear=True)
                dest[du] = {'uid': du}
                bound[uid] = du
                summary = content_summary(rec['raw']) if not diagnosed else None
                waiting[du] = ({key: rec[key] for key in
                    ('uid', 'canon', 'semantic', 'semantic_version')}, summary)
                if diagnostic_count < MESSAGE_DIAGNOSTIC_LIMIT or LOG_FILE is not None:
                    waiting[du][0]['log_header_bytes'] = message_header_bytes(rec['raw'])
                waiting_bytes += len(rec['raw'])
                if len(waiting) >= opts.batch_messages or waiting_bytes >= opts.batch_bytes:
                    check_uploaded()
                progress.copied_message(len(rec['raw']))
                progress.show()
        check_uploaded()
    progress.end_folder()
    progress.show(True)
    log()
    # The consumer may have spent minutes uploading the final source batch.
    # Session.read reconnects/reselects safely before using this connection again.
    current = src.snapshot()
    if current != {uid:{k:v for k,v in r.items() if k not in ('strict','canon','semantic','semantic_version')} for uid,r in source.items()}:
        raise RuntimeError(f'{f.name!r}: source changed during migration; rerun for a fresh snapshot')


def verify_folder(src, dst, j, f, df, opts):
    sv, source = inventory(src, j, 'source', f.name, opts, full=opts.full_verify)
    dv, dest = inventory(dst, j, 'destination', df, opts, full=opts.full_verify)
    def counter(records, metadata=False, strict=False):
        return collections.Counter((r['strict'] if strict else r['canon'],
                   tuple(r['flags']),r['epoch']) if metadata else
                   r['strict'] if strict else r['canon'] for r in records.values())
    bound = j.bindings(f.name, sv, df, dv)
    mapped = {su:du for su,du in bound.items() if su in source and du in dest}
    changed = 0
    equivalent_count = 0
    if len(mapped) == len(source):
        # Server-confirmed UID bindings distinguish transformed mail from absent mail.
        strict = sum(source[su]['strict'] == dest[du]['strict'] for su,du in mapped.items())
        changed = len(source) - strict
        equivalent_count = sum(source[su]['strict'] != dest[du]['strict'] and
            equivalent(source[su], dest[du]) for su,du in mapped.items())
        missing = 0
        extra = len(set(dest)-set(mapped.values()))
        meta = sum((source[su]['flags'],source[su]['epoch']) !=
                   (dest[du]['flags'],dest[du]['epoch']) for su,du in mapped.items())
    else:
        s,d = counter(source),counter(dest)
        sm,dm = counter(source,True),counter(dest,True)
        strict = sum((counter(source,strict=True) & counter(dest,strict=True)).values())
        missing,extra = sum((s-d).values()),sum((d-s).values())
        meta = sum((sm-dm).values()) + sum((dm-sm).values())
        changed = max(0, len(source)-strict-missing)
    # Report every known mismatching pair, even when other bindings are missing.
    # Never infer a pairing from a Message-ID or an unmatched content count.
    different_pairs = {}
    for su, du in sorted(mapped.items()):
        differences = []
        if source[su]['strict'] != dest[du]['strict']:
            differences.append('bytes (checked-equivalent)' if equivalent(source[su], dest[du])
                               else 'bytes (changed or uncertain)')
        if source[su]['flags'] != dest[du]['flags']:
            differences.append('flags')
        if source[su]['epoch'] != dest[du]['epoch']:
            differences.append('INTERNALDATE')
        if differences:
            different_pairs[su] = (du, differences)
    # Cap terminal output; a requested log receives every differing pair.
    # Fetch only headers, not another copy of every changed message body.
    terminal_uids = set(sorted(different_pairs)[:MESSAGE_DIAGNOSTIC_LIMIT])
    diagnostic_uids = sorted(different_pairs) if LOG_FILE is not None else sorted(terminal_uids)
    for ids in batches(source, diagnostic_uids, opts.batch_messages, opts.batch_bytes):
        diagnostics = src.fetch_headers(ids)
        for su in ids:
            du, differences = different_pairs[su]
            show_message_identity(src, f.name, sv, su, df, dv, du,
                'Verification difference: '+', '.join(differences),
                message_log_headers(diagnostics[su]), file_only=su not in terminal_uids)
    if len(different_pairs) > len(terminal_uids):
        log(f'  {df!r}: terminal details shown for {len(terminal_uids)}/{len(different_pairs)} '
            f'differing mapped messages; {len(different_pairs)-len(terminal_uids)} suppressed on terminal. '
            'All differences are counted; all UID mappings remain in the journal.')
    # Detect changes throughout body scans, including concurrent deletions/arrivals.
    def stable(s, recs):
        snap = s.snapshot()
        return snap == {uid:{k:v for k,v in r.items() if k not in ('strict','canon','semantic','semantic_version')} for uid,r in recs.items()}
    stable_source, stable_dest = stable(src,source),stable(dst,dest)
    ok = not (missing or extra or meta or changed) and stable_source and stable_dest
    checked_ok = not (missing or extra or meta or changed-equivalent_count) and stable_source and stable_dest
    status = 'BYTE IDENTICAL' if ok else ('EQUIVALENT UNDER CHECKED RULES; ORIGINAL BYTES CHANGED'
        if checked_ok else 'CHANGED OR UNCERTAIN / INCOMPLETE')
    log(f'  {f.name!r}: source={len(source)} destination={len(dest)} '
          f'journal-confirmed={len(mapped)} byte-identical={strict} '
          f'content-changed={changed} equivalent-under-checked-rules={equivalent_count} '
          f'changed-or-uncertain={changed-equivalent_count} missing={missing} extra={extra} '
          f'flag/date differences={meta} stable={stable_source and stable_dest}: '
          + status, flush=True)
    return ok


# ------------------------- verified repair -----------------------------------

def repair_matches(source, target):
    return ((source['strict'] == target['strict'] or equivalent(source, target))
            and source['flags'] == target['flags'] and source['epoch'] == target['epoch'])


def repair_content_description(source, target):
    if source['strict'] == target['strict']:
        return 'byte-identical'
    if source['canon'] == target['canon']:
        return 'identical except for line endings (original bytes differ)'
    if equivalent(source, target):
        return 'equivalent under checked rules (original bytes differ)'
    return 'changed or uncertain'


def show_repair_difference(src, task, source, target, diagnostic_state):
    count = diagnostic_state[0] if diagnostic_state is not None else 0
    file_only = count >= MESSAGE_DIAGNOSTIC_LIMIT
    if diagnostic_state is not None:
        diagnostic_state[0] += 1
    if file_only and count == MESSAGE_DIAGNOSTIC_LIMIT:
        log(f'  Further detailed repair diagnostics suppressed on terminal '
            f'(limit {MESSAGE_DIAGNOSTIC_LIMIT} per folder); '
            'use --log-file for every failed candidate. Failure summaries remain visible.')
    if file_only and LOG_FILE is None:
        return
    show_message_identity(src, task['sf'], task['sv'], task['su'], task['df'],
        task['dv'], task['new_du'], 'Repair candidate difference',
        message_log_headers(source['raw']), file_only=file_only)
    if source['flags'] != target['flags']:
        log(f'    Flags: source={source["flags"]!r}; candidate={target["flags"]!r}',
            file_only=file_only)
    if source['epoch'] != target['epoch']:
        log(f'    INTERNALDATE: source={source["date"]!r} (epoch {source["epoch"]}); '
            f'candidate={target["date"]!r} (epoch {target["epoch"]})', file_only=file_only)
    if source['strict'] != target['strict']:
        log(f'    Message size: source={len(source["raw"]):,} bytes; '
            f'candidate={len(target["raw"]):,} bytes', file_only=file_only)
        if source['canon'] == target['canon']:
            def endings(raw):
                crlf = raw.count(b'\r\n')
                lf, cr = raw.count(b'\n')-crlf, raw.count(b'\r')-crlf
                return f'CRLF={crlf:,}, bare LF={lf:,}, bare CR={cr:,}'
            log('    Line endings: source '+endings(source['raw'])+
                '; candidate '+endings(target['raw']), file_only=file_only)
            log('    Entire message matches after line-ending normalization; '
                'this is not byte-exact verification or signature validation.', file_only=file_only)
        else:
            show_content_difference(content_summary(source['raw']), target['raw'],
                file_only=file_only, label='Repair candidate')


def finish_repair(src, dst, journal, task, diagnostic_state=None):
    """Resume a candidate or committed cleanup without issuing another APPEND."""
    sf, sv, su, df, dv = (task[key] for key in ('sf','sv','su','df','dv'))
    if src.select(sf) != sv or dst.select(df) != dv:
        raise RuntimeError('Repair UIDVALIDITY changed; retained copies require review')
    if su not in src.snapshot():
        raise RuntimeError('Repair source message disappeared; no destination copy deleted')
    source = src.fetch([su])[su]
    if source['strict'] != task['source_strict']:
        raise RuntimeError('Repair source content changed; no destination copy deleted')
    destination = dst.snapshot()
    du = task['new_du']
    if du == task.get('old_du'):
        raise RuntimeError('Repair candidate equals old UID; no destination copy deleted')
    if du not in destination:
        # An acknowledged UID confirmed absent is not an ambiguous APPEND.
        # Retire an uncommitted candidate without deleting the old copy.
        # Committed cleanup may already have marked/expunged the old UID, so
        # it must not be rolled back as if no replacement had been committed.
        state = task.get('state')
        if state not in ('candidate', 'cleanup'):
            raise RuntimeError('Unknown repair state; missing candidate record retained')
        if state == 'cleanup':
            raise RuntimeError(f'Committed repair replacement UID {du} disappeared in {df!r}. '
                'Cleanup record retained; old copy may already be marked deleted or expunged. '
                'No automatic rollback/deletion attempted; review this committed repair separately.')
        current = journal.db.execute('SELECT df,dv,du FROM bindings WHERE sf=? AND sv=? AND su=?',
                                     (sf,sv,su)).fetchone()
        old = task.get('old_du')
        if current is not None:
            if current[:2] != (df,dv):
                raise RuntimeError('Repair binding destination changed; missing candidate record retained')
            if current[2] in destination and current[2] != old:
                raise RuntimeError('Repair binding points to another surviving UID; record retained')
        surviving_old = old if old is not None and old in destination else None
        if surviving_old is not None:
            previous = dst.fetch([old])[old]
            if previous['strict'] != task['old_strict']:
                raise RuntimeError('Old repair target content changed; missing candidate record retained')
        fresh = dst.snapshot()
        if du in fresh or (surviving_old is not None and surviving_old not in fresh):
            raise RuntimeError('Destination changed while confirming missing repair candidate; record retained')
        journal.save('source', sf, sv, source)
        journal.retire_missing_repair(task, surviving_old)
        log(f'  Repair candidate UID {du} is confirmed missing in {df!r} (state={state}). '
            + (f'Old UID {surviving_old} retained; ' if surviving_old is not None else 'No old copy remains; ')
            + f'source UID {su} will be uploaded again. No destination message deleted.')
        return None
    target = dst.fetch([du])[du]
    journal.save('source', sf, sv, source)
    journal.save('destination', df, dv, target)
    if not repair_matches(source, target):
        log(f'  REPAIR FAILED: {sf!r} source UID {su}, candidate UID {du}. '
              f'Content: {repair_content_description(source, target)}; '
              f'flags={"MATCH" if source["flags"] == target["flags"] else "DIFFER"}; '
              f'INTERNALDATE={"MATCH" if source["epoch"] == target["epoch"] else "DIFFER"}. '
              'Replacement failed content/flags/date verification; retained old and candidate copies. '
              'Reruns recheck this candidate without appending duplicates.', flush=True)
        show_repair_difference(src, task, source, target, diagnostic_state)
        return False
    old = task.get('old_du')
    if old is not None and old in destination:
        previous = dst.fetch([old])[old]
        if previous['strict'] != task['old_strict']:
            raise RuntimeError('Old repair target content changed; no destination copy deleted')
        dst.require_uid_expunge()
    # Commit the new mapping AND cleanup intent before any destructive command.
    # A crash after STORE/EXPUNGE can then only resume this exact UID cleanup.
    journal.commit_repair(task)
    if old is not None and old in destination:
        dst.delete_uid(df, dv, old)
    journal.finish_repair(task)
    log(f'  REPAIRED: {sf!r} source UID {su} -> destination UID {du}' +
          (f'; removed old UID {old}' if old is not None and old in destination else '; missing message restored'), flush=True)
    return True


def repair_folder(src, dst, journal, folder, df, delim, opts):
    ensure_folder(dst, df, delim)
    diagnostic_state = [0]
    retry_missing = set()
    # Reuse surviving candidates; confirmed missing candidates can be retired.
    for task in journal.repair_rows():
        if task['sf'] == folder.name:
            if task['df'] != df:
                raise RuntimeError('Repair destination differs from current folder mapping')
            if finish_repair(src, dst, journal, task, diagnostic_state) is None:
                retry_missing.add(task['su'])
    sv, source = inventory(src, journal, 'source', folder.name, opts, full=True)
    dv, destination = inventory(dst, journal, 'destination', df, opts, full=True)
    bound = journal.bindings(folder.name, sv, df, dv)
    staged = {task['su']:task for task in journal.repair_rows()
              if task['sf'] == folder.name and task['sv'] == sv and task['dv'] == dv}
    protected = {task['new_du'] for task in staged.values()}
    protected.update(task['old_du'] for task in staged.values() if task.get('old_du') is not None)
    unknown = set(destination) - set(bound.values()) - protected
    # Adopt legacy messages only under the existing exact-content/multiplicity rule.
    available = collections.defaultdict(list)
    for uid in sorted(unknown):
        available[destination[uid]['canon']].append(uid)
    adopted = []
    for uid in sorted(source):
        if uid in bound or uid in staged:
            continue
        matches = available[source[uid]['canon']]
        if matches:
            adopted.append((uid,matches.pop(0)))
    if any(available.values()):
        raise RuntimeError(f'{df!r}: untracked destination messages cannot be identified safely. '
                           'Repair will not delete or overwrite untracked messages.')
    for su,du in adopted:
        journal.bind(folder.name, sv, su, df, dv, du)
        bound[su] = du
    for su, expected in sorted(source.items()):
        check_stop()
        if su in staged:
            continue  # Already rechecked above; failed candidate remains journaled.
        old = bound.get(su)
        previous = destination.get(old)
        if su not in retry_missing and previous is not None and repair_matches(expected, previous):
            continue
        if previous is not None:
            dst.require_uid_expunge()  # Refuse before uploading if selective cleanup is impossible.
        rec = src.fetch([su])[su]
        if (rec['strict'],rec['flags'],rec['epoch']) != (expected['strict'],expected['flags'],expected['epoch']):
            raise RuntimeError('Source changed during repair; rerun for a fresh snapshot')
        journal.save('source', folder.name, sv, rec)
        log(f'  Uploading repair replacement: {folder.name!r} source UID {su} '
            f'-> {df!r}; {len(rec["raw"]):,} original bytes', flush=True)
        if time.monotonic()-dst.last_activity > 30:
            dst.read('pre-repair NOOP', lambda c: require_ok(c.noop(), 'NOOP'))
        task = dict(kind='repair', sf=folder.name, sv=sv, su=su, df=df, dv=dv,
                    before=max(destination, default=0), canon=rec['canon'],
                    source_strict=rec['strict'], old_du=old if previous is not None else None,
                    old_strict=previous['strict'] if previous is not None else None)
        journal.begin(task)
        try:
            du = dst.append(df, rec)
        except AppendRejected:
            journal.reject()
            raise
        except TRANSIENT as exc:
            log(f'  Repair APPEND connection lost: {exc}; reconciling without retry', flush=True)
            dst.close()
            reconcile(dst, journal, opts)
        else:
            if du is not None:
                journal.accepted(du)
                journal.stage_repair(journal.pending(), du)
            else:
                reconcile(dst, journal, opts)
        task = next(task for task in journal.repair_rows()
                    if (task['sf'],task['sv'],task['su']) == (folder.name,sv,su))
        finish_repair(src, dst, journal, task, diagnostic_state)
        # Refresh live metadata without discarding the content fingerprints.
        live = dst.snapshot()
        if set(live) - set(destination) - {task['new_du']}:
            raise RuntimeError('Destination changed during repair; review new untracked UIDs')
        destination = {uid:dict(destination[uid], **metadata) for uid,metadata in live.items()
                       if uid in destination}
        if task['new_du'] in live:
            cached = journal.record('destination', df, dv, task['new_du'])
            if cached is None:
                raise RuntimeError('Verified repair candidate missing from journal inventory')
            destination[task['new_du']] = dict(cached, **live[task['new_du']])



# ------------------------- portable EML archives -----------------------------

ARCHIVE_VERSION = 1
SNAPSHOT_KEYS = ('uid', 'size', 'date', 'epoch', 'flags', 'original_flags')


def archive_json(path):
    if path.is_symlink():
        raise RuntimeError(f'Archive refuses symbolic link: {path}')
    try:
        result = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, ValueError) as exc:
        raise RuntimeError(f'Cannot read archive metadata {path}: {exc}') from exc
    if not isinstance(result, dict):
        raise RuntimeError(f'Archive metadata is not an object: {path}')
    return result


def archive_bytes(obj):
    return (json.dumps(obj, ensure_ascii=False, indent=2, sort_keys=True)+'\n').encode('utf-8')


def sync_directory(path):
    if os.name != 'nt':
        fd = os.open(path, os.O_RDONLY)
        try:
            os.fsync(fd)
        finally:
            os.close(fd)


def archive_mkdir(path):
    missing = []
    current = path
    while not current.exists():
        missing.append(current)
        current = current.parent
    for directory in reversed(missing):
        directory.mkdir(mode=0o700)
        sync_directory(directory.parent)
    if not path.is_dir():
        raise RuntimeError(f'Archive directory path is not a directory: {path}')


def durable_replace(path, payload, epoch=None):
    """Commit bytes before metadata; never expose partial final files."""
    temp = path.with_name(path.name+'.partial')
    if path.is_symlink() or temp.is_symlink():
        raise RuntimeError(f'Archive refuses symbolic link: {path}')
    with temp.open('wb') as stream:
        os.chmod(temp, 0o600)
        stream.write(payload)
        stream.flush()
        if epoch is not None:
            os.utime(temp, (epoch, epoch))
        os.fsync(stream.fileno())
    os.replace(temp, path)
    sync_directory(path.parent)


def archive_path(root, relative):
    # Metadata may be hand edited: refuse traversal and symlink descendants.
    if not isinstance(relative, str) or not relative or '\\' in relative:
        raise RuntimeError('Invalid archive relative path')
    parts = relative.split('/')
    if any(not portable_component(part) for part in parts):
        raise RuntimeError(f'Unsafe archive relative path: {relative!r}')
    path = root
    for part in parts:
        path = path / part
        if path.is_symlink():
            raise RuntimeError(f'Archive refuses symbolic link: {path}')
    return path


def portable_component(name):
    reserved = {'CON', 'PRN', 'AUX', 'NUL', 'CONIN$', 'CONOUT$'}
    reserved.update(f'{prefix}{digit}' for prefix in ('COM', 'LPT')
                    for digit in '123456789¹²³')
    return (isinstance(name, str) and bool(name) and name not in ('.', '..')
            and not name.endswith((' ', '.'))
            and not any(ord(ch) < 32 or ord(ch) == 127 or ch in '<>:"/\\|?*' for ch in name)
            and name.split('.')[0].upper() not in reserved
            and len(name.encode('utf-8')) <= 100)


def portable_key(name):
    return unicodedata.normalize('NFC', name).casefold()


class ArchiveLock:
    def __init__(self, root):
        archive_mkdir(root)
        path = root / 'archive.lock'
        if path.is_symlink():
            raise RuntimeError('Archive lock is a symbolic link')
        self.stream = path.open('a+b')
        os.chmod(path, 0o600)
        try:
            if os.name == 'nt':
                import msvcrt
                self.stream.seek(0)
                self.stream.write(b'0')
                self.stream.flush()
                self.stream.seek(0)
                msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            self.stream.close()
            raise RuntimeError('Another process is using this archive') from exc

    def close(self):
        self.stream.close()


def assign_archive_folder(root, archive, folder):
    previous = archive['folders'].get(folder.name)
    if previous:
        if previous['delimiter'] != folder.delimiter:
            raise RuntimeError(f'{folder.name!r}: source hierarchy delimiter changed')
        archive_mkdir(archive_path(root, previous['path']))
        previous['flags'] = sorted(folder.flags)
        return previous
    parent = ''
    parts = folder.name.split(folder.delimiter) if folder.delimiter else [folder.name]
    nodes = archive['nodes']
    used = {portable_key(path) for path in nodes.values()}
    for index, part in enumerate(parts, 1):
        key = json.dumps([folder.delimiter, parts[:index]], ensure_ascii=True)
        if key in nodes:
            parent = nodes[key]
            archive_mkdir(archive_path(root, parent))
            continue
        candidate = '/'.join(filter(None, (parent, part)))
        reserved = part.casefold().removesuffix('.partial') in ('archive.metadata', 'snapshot.metadata', 'archive.lock',
                                      'folder.metadata') or re.fullmatch(r'uid-\d+\.(eml|metadata)(\.partial)?', part, re.I)
        if (not portable_component(part) or reserved or portable_key(candidate) in used
                or len(candidate.encode('utf-8')) > 180):
            alias = 'folder-'+hashlib.sha256(key.encode('ascii')).hexdigest()[:24]
            candidate = '/'.join(filter(None, (parent, alias)))
            if len(candidate.encode('utf-8')) > 180:
                candidate = alias  # Deep trees remain represented exactly in metadata.
        if portable_key(candidate) in used:
            raise RuntimeError('Archive folder alias collision; no files overwritten')
        directory = archive_path(root, candidate)
        if directory.exists():
            raise RuntimeError(f'Unregistered archive directory already exists: {directory}')
        nodes[key] = candidate
        # Reserve the mapping durably before creating its directory, so a crash
        # cannot leave an unregistered directory that prevents resuming.
        durable_replace(root / 'archive.metadata', archive_bytes(archive))
        archive_mkdir(directory)
        used.add(portable_key(candidate))
        parent = candidate
    entry = dict(name=folder.name, delimiter=folder.delimiter, path=parent,
                 flags=sorted(folder.flags), validity=None)
    archive['folders'][folder.name] = entry
    return entry


def archive_message_paths(root, entry, uid):
    directory = archive_path(root, entry['path'])
    stem = f'uid-{uid:012d}'
    eml, meta = directory / (stem+'.eml'), directory / (stem+'.metadata')
    if eml.is_symlink() or meta.is_symlink():
        raise RuntimeError(f'Archive message is a symbolic link: {eml}')
    return eml, meta


def read_archive_message(root, entry, uid, raw_required=True, fingerprint=True):
    eml, meta = archive_message_paths(root, entry, uid)
    rec = archive_json(meta)
    if (rec.get('uid') != uid or rec.get('validity') != entry['validity']
            or rec.get('folder') != entry['name'] or not isinstance(rec.get('date'), str)
            or rec.get('epoch') != date_epoch(rec['date'])
            or type(rec.get('size')) is not int or rec['size'] < 0
            or not isinstance(rec.get('flags'), list)
            or any(flag not in STANDARD_FLAGS.values() for flag in rec['flags'])
            or rec['flags'] != sorted(set(rec['flags']))
            or not isinstance(rec.get('original_flags'), list)
            or any(not isinstance(flag, str) for flag in rec['original_flags'])):
        raise RuntimeError(f'Invalid message metadata: {meta}')
    result = {key: rec[key] for key in SNAPSHOT_KEYS}
    if raw_required:
        raw = eml.read_bytes()
        strict, canon = hashes(raw)
        if len(raw) != rec.get('bytes') or strict != rec.get('sha256'):
            raise RuntimeError(f'Archive byte length/hash mismatch: {eml}')
        result.update(raw=raw, strict=strict, canon=canon)
        if fingerprint:
            result.update(semantic_version=SEMANTIC_VERSION, semantic=semantic_fingerprint(raw))
    return result



def validate_archive_header(root, archive):
    if (archive.get('format') != ARCHIVE_VERSION
            or not isinstance(archive.get('archive_id'), str)
            or not re.fullmatch(r'[0-9a-f]{32}', archive['archive_id'])
            or not isinstance(archive.get('folders'), dict)
            or not isinstance(archive.get('nodes'), dict)):
        raise RuntimeError('Invalid or unsupported archive header')
    paths = list(archive['nodes'].values())
    for relative in paths:
        archive_path(root, relative)
    if len({portable_key(path) for path in paths}) != len(paths):
        raise RuntimeError('Archive directory mapping collision')
    for name, entry in archive['folders'].items():
        if (not isinstance(entry, dict) or entry.get('name') != name
                or entry.get('path') not in paths
                or not isinstance(entry.get('flags'), list)
                or any(not isinstance(flag, str) for flag in entry['flags'])
                or (entry.get('delimiter') is not None and
                    (not isinstance(entry['delimiter'], str) or not entry['delimiter']))
                or (entry.get('validity') is not None and
                    (type(entry['validity']) is not int or not 0 < entry['validity'] < 2**32))):
            raise RuntimeError(f'Invalid archive folder entry: {name!r}')



def timestamp_archive_folders(root, archive, snapshots):
    log('Timestamping archive folders from their newest current messages...')
    explicit = {entry['path'] for entry in archive['folders'].values()}
    times = {}
    for name, records in snapshots.items():
        if not records:
            continue  # An empty folder has no source message timestamp to use.
        entry = archive['folders'][name]
        epoch = max(rec['epoch'] for rec in records.values())
        directory = archive_path(root, entry['path'])
        metadata = directory / 'folder.metadata'
        if metadata.is_symlink():
            raise RuntimeError(f'Archive refuses symbolic link: {metadata}')
        os.utime(metadata, (epoch, epoch))
        with metadata.open('rb') as stream:
            os.fsync(stream.fileno())
        times[entry['path']] = epoch
        parts = entry['path'].split('/')
        for length in range(1, len(parts)):
            prefix = '/'.join(parts[:length])
            if prefix not in explicit:
                times[prefix] = max(times.get(prefix, epoch), epoch)
    for relative, epoch in sorted(times.items(), key=lambda item: item[0].count('/'), reverse=True):
        directory = archive_path(root, relative)
        os.utime(directory, (epoch, epoch))
        sync_directory(directory)


def export_archive(src, root, opts):
    identity = [src.ep.server.casefold(), src.ep.user.casefold()]
    header = root / 'archive.metadata'
    if header.is_symlink():
        raise RuntimeError('Archive header is a symbolic link')
    if header.exists():
        archive = archive_json(header)
        if archive.get('format') != ARCHIVE_VERSION or archive.get('source') != identity:
            raise RuntimeError('Archive belongs to another source or unsupported format')
    else:
        if any(path.name not in ('archive.lock', 'archive.metadata.partial') for path in root.iterdir()):
            raise RuntimeError('New archive requires an empty directory')
        archive = dict(format=ARCHIVE_VERSION, archive_id=uuid.uuid4().hex,
                       source=identity, folders={}, nodes={}, complete=False)
    validate_archive_header(root, archive)
    log(f'Exporting source to {root}; existing historical message files are retained.')
    folders = [f for f in src.folders() if r'\noselect' not in f.flags]
    folders.sort(key=lambda f: (f.name.casefold() != 'inbox', f.name.casefold()))
    if not folders:
        raise RuntimeError('Source has no selectable folders')
    log('Counting source messages and checking folder identities before updating archive...')
    validities, initial_counts = {}, {}
    for index, folder in enumerate(folders, 1):
        log(f'  Counting [{index}/{len(folders)}] {folder.name!r}')
        validity = src.select(folder.name)
        previous = archive['folders'].get(folder.name)
        if previous and previous['delimiter'] != folder.delimiter:
            raise RuntimeError(f'{folder.name!r}: source hierarchy delimiter changed; use a new archive')
        if previous and previous['validity'] is not None and previous['validity'] != validity:
            raise RuntimeError(f'{folder.name!r}: UIDVALIDITY changed; use a new archive directory')
        validities[folder.name] = validity
        initial_counts[folder.name] = len(src.snapshot())
    progress = ArchiveProgress(sum(initial_counts.values()))
    archive['complete'] = False
    durable_replace(header, archive_bytes(archive))
    snapshot = dict(format=ARCHIVE_VERSION, archive_id=archive['archive_id'], folders={})
    expected_snapshots = {}
    for index, folder in enumerate(folders, 1):
        check_stop()
        log(f'Export [{index}/{len(folders)}] {folder.name!r}')
        entry = assign_archive_folder(root, archive, folder)
        validity = src.select(folder.name)
        if validity != validities[folder.name] or (entry['validity'] is not None and entry['validity'] != validity):
            raise RuntimeError(f'{folder.name!r}: UIDVALIDITY changed; use a new archive directory')
        entry['validity'] = validity
        durable_replace(header, archive_bytes(archive))
        folder_metadata = archive_path(root, entry['path']) / 'folder.metadata'
        if folder_metadata.is_symlink():
            raise RuntimeError(f'Archive refuses symbolic link: {folder_metadata}')
        folder_payload = archive_bytes(entry)
        if not folder_metadata.exists() or folder_metadata.read_bytes() != folder_payload:
            durable_replace(folder_metadata, folder_payload)
        records = src.snapshot()
        expected_snapshots[folder.name] = records
        progress.total += len(records)-initial_counts[folder.name]
        progress.begin_archive_folder(len(records))
        todo, cached, known_hashes = [], 0, {}
        for uid, rec in records.items():
            check_stop()
            if time.monotonic()-src.last_activity > 30:
                src.read('export keepalive NOOP', lambda c: require_ok(c.noop(), 'source NOOP'))
            eml, meta = archive_message_paths(root, entry, uid)
            if not (eml.exists() and meta.exists()):
                todo.append(uid)
                continue
            old = read_archive_message(root, entry, uid, fingerprint=False)
            if old['size'] != rec['size'] or old['epoch'] != rec['epoch']:
                raise RuntimeError(f'{folder.name!r} UID {uid}: immutable source metadata changed')
            known_hashes[uid] = old['strict']
            if opts.full_verify:
                todo.append(uid)
                continue
            metadata = dict(rec, folder=folder.name, validity=validity,
                            bytes=len(old['raw']), sha256=old['strict'])
            payload = archive_bytes(metadata)
            if meta.read_bytes() != payload:
                durable_replace(meta, payload, rec['epoch'])
            else:
                os.utime(meta, (rec['epoch'], rec['epoch']))
            os.utime(eml, (rec['epoch'], rec['epoch']))
            cached += 1
            progress.reused_message()
        if records:
            log(f'  {cached} already archived; {len(todo)} to fetch')
        with Prefetch(src, records, todo, opts) as stream:
            for rec in stream:
                uid = rec['uid']
                if {key:rec[key] for key in SNAPSHOT_KEYS} != records[uid]:
                    raise RuntimeError(f'{folder.name!r} UID {uid}: changed during export; rerun')
                if uid in known_hashes and known_hashes[uid] != rec['strict']:
                    raise RuntimeError(f'{folder.name!r} UID {uid}: source bytes changed under the same UID; archive retained')
                eml, meta = archive_message_paths(root, entry, uid)
                # Existing orphan EMLs have no commit metadata and may be replaced.
                if uid not in known_hashes:
                    durable_replace(eml, rec['raw'], rec['epoch'])
                else:
                    os.utime(eml, (rec['epoch'], rec['epoch']))
                metadata = dict(records[uid], folder=folder.name, validity=validity,
                                bytes=len(rec['raw']), sha256=rec['strict'])
                durable_replace(meta, archive_bytes(metadata), rec['epoch'])
                progress.completed_message(len(rec['raw']))
        progress.show(True)
        log()
        if src.snapshot() != records:
            raise RuntimeError(f'{folder.name!r}: source changed during export; rerun')
        snapshot['folders'][folder.name] = dict(validity=validity, uids=sorted(records))
    log('Checking exported source snapshot before committing completion...')
    if {f.name for f in src.folders() if r'\noselect' not in f.flags} != set(expected_snapshots):
        raise RuntimeError('Source folder set changed during export; rerun')
    for folder in folders:
        if src.select(folder.name) != archive['folders'][folder.name]['validity'] or src.snapshot() != expected_snapshots[folder.name]:
            raise RuntimeError(f'{folder.name!r}: source changed before export completed; rerun')
    timestamp_archive_folders(root, archive, expected_snapshots)
    payload = archive_bytes(snapshot)
    durable_replace(root / 'snapshot.metadata', payload)
    archive.update(complete=True, snapshot_sha256=hashlib.sha256(payload).hexdigest(),
                   completed_utc=_dt.datetime.now(_dt.timezone.utc).isoformat())
    durable_replace(header, archive_bytes(archive))
    log(f'EXPORT COMPLETE: {sum(len(item["uids"]) for item in snapshot["folders"].values())} '
        f'messages in {len(folders)} folders. Archive retained at {root}.')
    return 0


class ExportSession(Session):
    def fetch(self, ids, bodies=True):
        fields = '(UID FLAGS INTERNALDATE RFC822.SIZE' + (' BODY.PEEK[]' if bodies else '') + ')'
        def run(c):
            data = require_ok(c.uid('FETCH', ','.join(map(str, ids)), fields), 'export FETCH')
            records = fetch_records(data, bodies, all_flags=True)
            if any(uid not in records for uid in ids):
                raise RuntimeError('Source FETCH omitted requested UID during export')
            return {uid:records[uid] for uid in ids}
        return self.read('export FETCH', run)


class ArchiveSource:
    """Read-only Session interface; no source network connection is opened."""
    def __init__(self, root):
        self.root = root
        header = archive_path(root, 'archive.metadata')
        self.archive = archive_json(header)
        validate_archive_header(root, self.archive)
        if self.archive.get('format') != ARCHIVE_VERSION or self.archive.get('complete') is not True:
            raise RuntimeError('Archive is unsupported/incomplete; finish --export before --import')
        snapshot_path = archive_path(root, 'snapshot.metadata')
        payload = snapshot_path.read_bytes()
        if hashlib.sha256(payload).hexdigest() != self.archive.get('snapshot_sha256'):
            raise RuntimeError('Archive snapshot hash mismatch')
        try:
            self.snapshot_manifest = json.loads(payload)
        except ValueError as exc:
            raise RuntimeError('Invalid archive snapshot JSON') from exc
        if not isinstance(self.snapshot_manifest, dict):
            raise RuntimeError('Archive snapshot is not an object')
        if (self.snapshot_manifest.get('format') != ARCHIVE_VERSION
                or self.snapshot_manifest.get('archive_id') != self.archive.get('archive_id')
                or not isinstance(self.snapshot_manifest.get('folders'), dict)):
            raise RuntimeError('Invalid archive snapshot')
        identity = self.archive.get('source')
        if not isinstance(identity, list) or len(identity) != 2 or not all(isinstance(x, str) and x for x in identity):
            raise RuntimeError('Invalid archive source identity')
        self.ep = EndpointConfig(*identity, '', '', '', '')
        self.mailbox, self.validity = None, None
        self.last_activity = time.monotonic()
        self.records = {}
        self.content_hashes = {}
        self.entries = {}
        used_paths = set()
        log('Validating archive message pairs and SHA-256 hashes before destination connection...')
        validation = ArchiveProgress(sum(len(snap['uids']) for snap in
            self.snapshot_manifest['folders'].values()), operation='validate')
        for name, snap in self.snapshot_manifest['folders'].items():
            check_stop()
            entry = self.archive['folders'][name]
            if entry['name'] != name or entry['validity'] != snap['validity'] or type(entry['validity']) is not int or entry['validity'] < 1:
                raise RuntimeError('Archive folder identity mismatch')
            key = portable_key(entry['path'])
            if key in used_paths:
                raise RuntimeError('Archive folder path collision')
            used_paths.add(key)
            directory = archive_path(root, entry['path'])
            if archive_json(directory / 'folder.metadata') != entry:
                raise RuntimeError(f'Folder metadata mismatch: {name!r}')
            if entry['delimiter'] is not None and (not isinstance(entry['delimiter'], str) or not entry['delimiter']):
                raise RuntimeError('Invalid archive folder delimiter')
            uids = snap['uids']
            if not isinstance(uids, list) or any(type(uid) is not int or not 0 < uid < 2**32 for uid in uids) or uids != sorted(set(uids)):
                raise RuntimeError('Invalid archive snapshot UID list')
            self.entries[name] = entry
            records = {}
            self.content_hashes[name] = {}
            log(f'  Validating {name!r}: {len(uids)} messages')
            validation.begin_archive_folder(len(uids))
            for uid in uids:
                check_stop()
                rec = read_archive_message(root, entry, uid, fingerprint=False)
                records[uid] = {key:rec[key] for key in SNAPSHOT_KEYS}
                self.content_hashes[name][uid] = rec['strict']
                validation.completed_message(len(rec['raw']))
            validation.show(True)
            log()
            self.records[name] = records

    def folders(self):
        return [Folder(name, entry['delimiter'], frozenset(entry['flags'])) for name, entry in self.entries.items()]

    def select(self, mailbox):
        if mailbox not in self.entries:
            raise RuntimeError(f'Folder absent from archive snapshot: {mailbox!r}')
        self.mailbox = mailbox
        self.validity = self.entries[mailbox]['validity']
        return self.validity

    def snapshot(self):
        return {uid:dict(rec) for uid,rec in self.records[self.mailbox].items()}

    def fetch(self, ids, bodies=True):
        result = {}
        for uid in ids:
            check_stop()
            rec = read_archive_message(self.root, self.entries[self.mailbox], uid, raw_required=bodies)
            if bodies and rec['strict'] != self.content_hashes[self.mailbox][uid]:
                raise RuntimeError('Archive message bytes changed during import')
            if {key:rec[key] for key in SNAPSHOT_KEYS} != self.records[self.mailbox][uid]:
                raise RuntimeError('Archive metadata changed during import')
            result[uid] = rec
        self.last_activity = time.monotonic()
        return result

    def fetch_headers(self, ids):
        headers = {}
        for uid in ids:
            check_stop()
            eml, _ = archive_message_paths(self.root, self.entries[self.mailbox], uid)
            with eml.open('rb') as stream:
                headers[uid] = stream.read(64*1024)
        return headers

    def read(self, operation, func):
        check_stop()
        return func(self)

    def noop(self):
        return 'OK', []

    def close(self):
        pass



def main(argv=None):
    global CONFIG_PATH, _stop_requested, RUN_STARTED, _progress_rows, LOG_FILE
    RUN_STARTED = time.monotonic()
    _progress_rows = 0
    parser = argparse.ArgumentParser(description=__doc__,
        epilog='Default: copy/resume from actual source, destination and journal state; no deletion.\n'
               '--repair: copy/resume plus verified replacement/deletion; always full verification.\n'
               '--verify-only: compare without uploading/deleting.\n'
               '--retry-pending: explicit recovery of an ambiguous upload, not a completion/resume switch.',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    archive_modes = parser.add_mutually_exclusive_group()
    archive_modes.add_argument('--export', dest='export_archive', action='store_true',
        help='Export/refresh a resumable EML and per-message metadata backup; source connection only')
    archive_modes.add_argument('--import', dest='import_archive', action='store_true',
        help='Upload the latest completed archive snapshot; destination connection only')
    parser.add_argument('--path', type=Path, help='Archive directory, required with --export/--import')
    parser.add_argument('--provider-notes', action='store_true', help='Show observed provider preservation behavior and exit')
    parser.add_argument('--config', type=Path, help='INI path (default: beside script)')
    parser.add_argument('--log-file', type=Path,
        help='Append timestamped status, periodic progress and all changed-message identities/Subjects/Dates to this UTF-8 file')
    parser.add_argument('--skip-destination-tree-deploy', action='store_true',
        help='Skip upfront folder creation, writable-selection checks and subscriptions; folders are still checked/created as reached during copying or repair')
    parser.add_argument('--full-verify', action='store_true', help='Re-download both sides for independent content verification')
    parser.add_argument('--retry-pending', action='store_true', help='Explicitly retry an unresolved upload only if no new destination UIDs exist; a delayed duplicate remains possible')
    parser.add_argument('--resolve-pending-uid', type=int, metavar='UID', help='Explicitly identify a pending uploaded message after inspecting its candidate UID')
    parser.add_argument('--repair', action='store_true', help='Copy/resume plus verified repair/deletion of damaged copies; implies --full-verify')
    parser.add_argument('--verify-only', action='store_true', help='Verify without creating folders or uploading messages')
    args = parser.parse_args(argv)
    if (args.export_archive or args.import_archive) != (args.path is not None):
        parser.error('--path is required with --export/--import and cannot be used without them')
    if args.export_archive and (args.repair or args.verify_only or args.retry_pending
                               or args.resolve_pending_uid is not None or args.skip_destination_tree_deploy):
        parser.error('--export cannot be combined with destination operation options')
    if args.provider_notes and (args.export_archive or args.import_archive):
        parser.error('--provider-notes cannot be combined with --export/--import')
    if args.retry_pending and args.resolve_pending_uid is not None:
        parser.error('--retry-pending and --resolve-pending-uid are mutually exclusive')
    if args.verify_only and (args.retry_pending or args.resolve_pending_uid is not None):
        parser.error('Pending recovery options cannot be combined with --verify-only')
    if args.resolve_pending_uid is not None and args.resolve_pending_uid < 1:
        parser.error('--resolve-pending-uid must be positive')
    if args.repair and args.verify_only:
        parser.error('--repair and --verify-only are mutually exclusive')
    if args.provider_notes:
        show_provider_notes()
        return 0
    CONFIG_PATH = args.config.expanduser().resolve() if args.config else None
    signal.signal(signal.SIGINT, _request_stop)
    signal.signal(signal.SIGTERM, _request_stop)
    src = dst = journal = archive_lock = None
    # Applies to the SQLite journal and sidecars, without changing existing INI permissions.
    previous_umask = os.umask(0o077)
    try:
        source_cfg, dest_cfg, root = load_config('export' if args.export_archive else
                                                  'import' if args.import_archive else 'migration')
        opts = options(args)
        if args.log_file:
            log_path = args.log_file.expanduser().absolute()
            if log_path.resolve() in (ini_path().resolve(), opts.journal.resolve()):
                raise RuntimeError('Log file must differ from configuration and journal')
            descriptor = os.open(log_path, os.O_WRONLY | os.O_CREAT | os.O_APPEND, 0o600)
            LOG_FILE = os.fdopen(descriptor, 'a', encoding='utf-8')
            log('\nRun started at '+_dt.datetime.now(_dt.timezone.utc).isoformat())
            log(f'Log file: {str(log_path)!r}; includes message Subjects and Dates.')
        if args.export_archive or args.import_archive:
            archive_root = args.path.expanduser().resolve()
            if args.import_archive and not archive_root.is_dir():
                raise RuntimeError(f'Archive directory does not exist: {archive_root}')
            archive_lock = ArchiveLock(archive_root)
            if args.export_archive:
                log(f'Configuration: {ini_path()}')
                src = ExportSession(source_cfg, 'source', opts.retries)
                return export_archive(src, archive_root, opts)
            src = ArchiveSource(archive_root)
            source_cfg = src.ep
            opts.full_verify = True
            archive_id = src.archive['archive_id']
            if not isinstance(archive_id, str) or not re.fullmatch(r'[0-9a-f]{32}', archive_id):
                raise RuntimeError('Invalid archive identifier')
            opts.journal = opts.journal.with_name(opts.journal.stem+'-import-'+archive_id[:12]+opts.journal.suffix)
            log('Import uses the completed local snapshot; full byte verification is enabled.')
        if args.repair:
            opts.full_verify = True
        mode = ('verify only; no upload/deletion' if args.verify_only else
                'copy/resume and verified repair; full verification; selective deletion enabled'
                if args.repair else 'copy/resume; restore missing mail; no destination deletion')
        log('Mode: ' + mode)
        if args.repair and args.full_verify:
            log('Note: --repair already implies --full-verify.')
        show_provider_notes(dest_cfg.server)
        if not args.import_archive and (source_cfg.server.casefold(), source_cfg.user.casefold()) == (dest_cfg.server.casefold(), dest_cfg.user.casefold()):
            raise RuntimeError('Source and destination must be different accounts')
        journal = Journal(opts.journal, dict(version=2, source=[source_cfg.server.casefold(), source_cfg.user.casefold()],
            destination=[dest_cfg.server.casefold(),dest_cfg.user.casefold()], root=root,
            **({'archive_id': src.archive['archive_id']} if args.import_archive else {})))
        log(f'Configuration: {ini_path()}\nJournal: {opts.journal}', flush=True)
        if src is None:
            src = Session(source_cfg,'source',opts.retries)
        dst = Session(dest_cfg,'destination',opts.retries)
        log('Enumerating archive folders...' if args.import_archive else 'Enumerating source folders...', flush=True)
        folders = [f for f in src.folders() if r'\noselect' not in f.flags]
        if not folders:
            raise RuntimeError('Source has no selectable folders')
        folders.sort(key=lambda f:(f.name.casefold() != 'inbox',f.name.casefold()))
        log('Enumerating destination folders and mapping the source tree...', flush=True)
        destination_folders = dst.folders()
        delim, names = mapping(folders,destination_folders,root)
        if not args.verify_only:
            check_provider_layout(dest_cfg.server, names, delim)
            if args.skip_destination_tree_deploy:
                log('Skipping destination tree preparation/subscriptions (--skip-destination-tree-deploy). '
                    'Folders will be checked/created as reached during copying or repair.')
            else:
                prepare_destination_tree(dst, names, delim,
                    source=None if args.import_archive else src, known_folders=destination_folders)
        if not args.verify_only:
            reconcile(dst, journal, opts, retry_missing=args.retry_pending,
                      source=src, resolve_uid=args.resolve_pending_uid)
            tasks = journal.repair_rows()
            if tasks and not args.repair:
                raise RuntimeError('A previous repair has outstanding candidates/cleanup. '
                    'Default copy/resume will not finish a deletion: rerun with --repair '
                    '(full verification is implied), or inspect with --verify-only --full-verify.')
            if any(task['sf'] not in names or task['df'] != names[task['sf']] for task in tasks):
                raise RuntimeError('Outstanding repair falls outside current migration folder mappings')
            log('Counting source messages...', flush=True)
            total = 0
            folder_counts = {}
            for n, f in enumerate(folders, 1):
                log(f'  Counting [{n}/{len(folders)}] {f.name!r}', flush=True)
                src.select(f.name)
                folder_counts[f.name] = len(src.snapshot())
                total += folder_counts[f.name]
            log(f'Source: {total:,} messages in {len(folders):,} folders. Destination root: {root!r}')
            progress = CopyProgress(total, folder_counts)
            for n,f in enumerate(folders,1):
                log(f'\n[{n}/{len(folders)}] {f.name!r} -> {names[f.name]!r}', flush=True)
                if args.repair:
                    repair_folder(src,dst,journal,f,names[f.name],delim,opts)
                else:
                    migrate_folder(src,dst,journal,f,names[f.name],delim,opts,progress)
        elif journal.pending():
            raise RuntimeError('A pending APPEND exists. Resume migration to reconcile it first.')
        log('\nVerification: '+('independent full body download' if opts.full_verify else 'live UID/metadata checks with cached content hashes'))
        ok = True
        for n, f in enumerate(folders, 1):
            log(f'  Verifying [{n}/{len(folders)}] {f.name!r} -> {names[f.name]!r}', flush=True)
            result = verify_folder(src,dst,journal,f,names[f.name],opts)
            ok = result and ok
        # New folders also invalidate the result; no silent omission of arriving mailboxes.
        final_folders = {f.name for f in src.folders() if r'\noselect' not in f.flags}
        ok = ok and final_folders == {f.name for f in folders}
        if journal.repair_rows():
            ok = False
            log('Outstanding repair candidates remain; no failed replacement was used to delete an old copy.')
        log('\nFINAL RESULT: '+('VERIFIED' if ok else 'VERIFICATION DIFFERENCES DETECTED (see folder results)'))
        return 0 if ok else 2
    except KeyboardInterrupt:
        message = ('Stopped. Archive files are retained; rerun --export with the same --path to resume.'
                   if args.export_archive else 'Stopped. Keep the journal and rerun to resume safely.')
        log('\n'+message, file=sys.stderr)
        return 130
    except (RuntimeError, imaplib.IMAP4.error, OSError, ValueError, KeyError, TypeError, sqlite3.Error, configparser.Error) as exc:
        advice = ('Source messages were not deleted. Archive files are retained; interrupted exports '
                  'can resume with --export and the same --path.' if args.export_archive else
                  'Source messages were not deleted. Keep the journal when resuming.')
        log(f'\nERROR: {exc}\n{advice}', file=sys.stderr)
        return 1
    finally:
        for item in (src,dst,journal,archive_lock):
            if item:
                item.close()
        if LOG_FILE is not None:
            with contextlib.suppress(OSError):
                LOG_FILE.close()
            LOG_FILE = None
        os.umask(previous_umask)


if __name__ == '__main__':
    sys.exit(main())
