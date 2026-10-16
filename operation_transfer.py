"""Conservative, insert-only transfer of database operations after a workbook import.

Discovery and review always use read-only transactions. The only production DML
is a parameterized INSERT into the four operation tables in apply_transfer().
"""

from collections import Counter, defaultdict
from contextlib import contextmanager
from datetime import datetime, timezone, timedelta
import hashlib
import hmac
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import uuid
from decimal import Decimal

import fdb
import fdb.services

from config_utils import KEY_PATH, is_developer_mode_enabled
from merge_processor import read_merge_file


OPERATIONS = {
    "CROSSCUTS": ("CC_ID", "LENGTE", "SPLIT", "HOLES_L", "HOLES_S", "LSWING"),
    "CUTOUTS": ("CUTOUT1", "CUTOUT2"),
    "MITRES": ("MITRE_L", "MITRE_W"),
    "GROOVES": ("GROOVE_L", "GROOVE_W"),
}
LOCK_TABLES = ("QUOTE", "CUSTOMER", "CUTLIST", "CUT_LIST_DETAIL", "BOARD_TYPES",
               "BOARD_CATEGORY", "EDGE_TYPES", "EDGE_CATEGORY", *OPERATIONS)
SUPPORTED_SCHEMA = 24


class TransferError(ValueError):
    pass


class CommitOutcomeUnknown(TransferError):
    pass


def _text(value):
    return "" if value is None else str(value).strip()


def _json(value):
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True)


def digest(value):
    return hashlib.sha256(_json(value).encode()).hexdigest()


def file_hash(path):
    result = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            result.update(chunk)
    return result.hexdigest()


def database_identity(config):
    return {"host": _text(config.get("ip")).casefold(),
            "port": str(config.get("port") or "3050"),
            "path": str(config["filepath"]).replace("\\", "/").casefold()}


def _require_developer(config):
    if not is_developer_mode_enabled(config or {}):
        raise TransferError("A valid Developer Password is required for extra-operation transfers.")


def firebird_tool(name):
    candidates = []
    for root in [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]:
        if root:
            candidates.extend(sorted((Path(root) / "Firebird").glob(f"Firebird_3*/{name}"), reverse=True))
    found = shutil.which(name)
    if found:
        candidates.append(Path(found))
    for path in candidates:
        if path.is_file():
            return str(path)
    raise TransferError(f"Firebird 3 tool {name} was not found. A verified backup is required before transferring.")


def _connect(config, *, write=False):
    library = config.get("fb_library") or firebird_tool("fbclient.dll")
    kwargs = dict(user=config["username"], password=config["password"],
                  charset=config.get("charset") or "UTF8", fb_library_name=library,
                  no_db_triggers=True, no_gc=True,
                  isolation_level=fdb.ISOLATION_LEVEL_SNAPSHOT if write else fdb.ISOLATION_LEVEL_READ_COMMITED_RO)
    if config.get("ip"):
        kwargs["dsn"] = f"{config['ip']}/{config.get('port') or '3050'}:{config['filepath']}"
    else:
        kwargs["database"] = str(config["filepath"])
    return fdb.connect(**kwargs)


@contextmanager
def readonly(config):
    connection = _connect(config)
    try:
        connection.begin(tpb=fdb.ISOLATION_LEVEL_READ_COMMITED_RO)
        yield connection.cursor()
    finally:
        try:
            connection.rollback()
        finally:
            connection.close()


def _schema(cursor):
    cursor.execute("SELECT VERSION FROM DB_INFO")
    if cursor.fetchall() != [(SUPPORTED_SCHEMA,)]:
        raise TransferError("This transfer supports Cutting Manager database schema 24 only; no changes were made.")
    cursor.execute("SELECT RDB$TRIGGER_NAME FROM RDB$TRIGGERS WHERE COALESCE(RDB$SYSTEM_FLAG,0)=0 AND COALESCE(RDB$TRIGGER_INACTIVE,0)=0")
    if cursor.fetchone() is not None:
        raise TransferError("Custom database triggers require review before extra-operation transfers can be enabled.")
    for table, fields in OPERATIONS.items():
        cursor.execute("SELECT RDB$RELATION_TYPE, RDB$EXTERNAL_FILE FROM RDB$RELATIONS WHERE RDB$RELATION_NAME=?", (table,))
        relation = cursor.fetchone()
        if relation is None or relation[0] not in (None, 0) or relation[1] is not None:
            raise TransferError(f"{table} is not a normal database table.")
        cursor.execute("SELECT RDB$FIELD_NAME FROM RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME=?", (table,))
        actual = {_text(row[0]) for row in cursor.fetchall()}
        if actual != {"QUOTE_NR", "CUTLIST_ID", "ITEM_ID", *fields}:
            raise TransferError(f"Unexpected {table} schema; a transfer is unsafe.")


def _quote(cursor, number):
    cursor.execute("SELECT q.QUOTE_NR, q.JOB_NAME, c.NAME, q.SIZEMODE FROM QUOTE q JOIN CUSTOMER c ON c.CUSTOMER_ID=q.CUSTOMER_ID WHERE q.QUOTE_NR=?", (number,))
    header = cursor.fetchall()
    if len(header) != 1:
        raise TransferError(f"Quote {number} was not found uniquely in the configured database.")
    cursor.execute("SELECT CUTLIST_ID, STATUS, OPTIMIZE_MODE, BUILD_TYPE FROM CUTLIST WHERE QUOTE_NR=? ORDER BY CUTLIST_ID", (number,))
    lists = [list(row) for row in cursor.fetchall()]
    cursor.execute("""SELECT d.CUTLIST_ID,d.ITEM_ID,d.UNIT_ID,d.LENGTE,d.WYDTE,d.QTY,
        d.EDGEL,d.EDGEW,d.HOLESL,d.HOLESW,d.REMARKS,d.BOARD_ID,d.EDGE_ID,
        b.BOARD_NAME,b.BOARD_THICKNESS,bc.BCAT_NAME,e.EDGE_NAME,ec.ECAT_NAME
        FROM CUT_LIST_DETAIL d
        JOIN BOARD_TYPES b ON b.BOARD_ID=d.BOARD_ID
        JOIN BOARD_CATEGORY bc ON bc.BCAT_ID=b.BCAT_ID
        LEFT JOIN EDGE_TYPES e ON e.EDGE_ID=d.EDGE_ID
        LEFT JOIN EDGE_CATEGORY ec ON ec.ECAT_ID=e.ECAT_ID
        WHERE d.QUOTE_NR=? ORDER BY d.CUTLIST_ID,d.ITEM_ID""", (number,))
    rows = [list(row) for row in cursor.fetchall()]
    cursor.execute("SELECT COUNT(*) FROM CUT_LIST_DETAIL WHERE QUOTE_NR=?", (number,))
    if cursor.fetchone()[0] != len(rows):
        raise TransferError(f"Quote {number} contains invalid or missing material references.")
    keys = [(r[0], r[1]) for r in rows]
    if len(keys) != len(set(keys)) or len(lists) != len({r[0] for r in lists}):
        raise TransferError(f"Quote {number} contains duplicate database item/cutlist identifiers.")
    operations = {}
    for table, fields in OPERATIONS.items():
        cursor.execute(f"SELECT CUTLIST_ID,ITEM_ID,{','.join(fields)} FROM {table} WHERE QUOTE_NR=?", (number,))
        values = [list(row) for row in cursor.fetchall()]
        if any((r[0], r[1]) not in set(keys) for r in values):
            raise TransferError(f"Quote {number} has {table} records without a matching piece.")
        operations[table] = sorted(values, key=_json)
    return {"header": list(header[0]), "cutlists": lists, "rows": rows, "operations": operations}


def _signature(values, board, category, thickness, edge_name, edge_category):
    # Native DB dimensions/counts are integers. Never round an export to make it fit.
    numbers = []
    for value in values[:8]:
        try:
            number = int(value)
            if isinstance(value, bool) or float(value) != number:
                raise ValueError
        except (TypeError, ValueError, OverflowError):
            raise TransferError(f"Non-integer or invalid piece value {value!r}; no safe database match.") from None
        numbers.append(number)
    try:
        if isinstance(thickness, bool) or float(thickness) != int(thickness) or int(thickness) <= 0:
            raise ValueError
    except (TypeError, ValueError, OverflowError):
        raise TransferError("Invalid board thickness; no safe database match.") from None
    edged = numbers[4] > 0 or numbers[5] > 0
    return (tuple(numbers), _text(values[8]), _text(board).casefold(), _text(category).casefold(),
            int(thickness), _text(edge_name).casefold() if edged else "",
            _text(edge_category).casefold() if edged else "")


def _db_signature(row):
    return _signature(row[2:11], row[13], row[15], row[14], row[16], row[17])


def _export_signature(source, row_number, *, unit_id=None):
    sheet = source.sheet
    values = [sheet.cell(row_number, c).value for c in range(1, 12)]
    if unit_id is not None:
        values[0] = unit_id
    return _signature([*values[:8], values[10]], sheet["E4"].value, sheet["C4"].value,
                      sheet["J4"].value, values[9] or sheet["E5"].value,
                      values[8] or sheet["C5"].value)


def _item_operations(snapshot, cutlist, item):
    return {table: [row[2:] for row in rows if row[:2] == [cutlist, item]]
            for table, rows in snapshot["operations"].items()}


def _validate_operations(operations):
    for table, rows in operations.items():
        if table not in OPERATIONS:
            raise TransferError("Unknown extra-operation type.")
        if table != "CROSSCUTS" and len(rows) > 1:
            raise TransferError(f"Multiple {table} records on one piece require manual review.")
        seen = set()
        for row in rows:
            if len(row) != len(OPERATIONS[table]):
                raise TransferError(f"Invalid {table} record.")
            if table == "CROSSCUTS":
                if (any(not isinstance(v, int) or isinstance(v, bool) or v < 0 for v in row[:5])
                        or row[1] == 0 or row[5] not in ("0", "1") or row[0] in seen):
                    raise TransferError("Invalid or duplicate crosscut data requires manual review.")
                seen.add(row[0])
            elif any(v is not None and (not isinstance(v, str) or len(v) > 100) for v in row):
                raise TransferError(f"Unexpected {table} data requires manual review.")


def inspect_sources(files, config, overrides=None):
    """Read-only discovery. An unverified source never counts as 'no extra cuts'."""
    _require_developer(config)
    overrides = overrides or {}
    result = {"database": database_identity(config), "sources": [], "errors": {}, "high_water": 0}
    with readonly(config) as cursor:
        _schema(cursor)
        cursor.execute("SELECT MAX(QUOTE_NR) FROM QUOTE")
        result["high_water"] = cursor.fetchone()[0] or 0
        for file in files:
            path = str(file.path)
            try:
                match = re.match(r"^(\d+)\s*-", file.path.name)
                number = overrides.get(path) or (int(match[1]) if match else None)
                if not isinstance(number, int) or number <= 0:
                    raise TransferError("Set this file's original Cutting Manager quote number.")
                snapshot = _quote(cursor, number)
                if _text(snapshot["header"][1]) != file.job:
                    raise TransferError("The workbook job name does not match the source quote.")
                if {r[0] for r in snapshot["cutlists"]} != {int(s.sheet.title[7:]) for s in file.cutlists}:
                    raise TransferError("Workbook cutlists do not match the complete source quote.")
                mapping = []
                for source in file.cutlists:
                    cutlist = int(source.sheet.title[7:])
                    candidates = defaultdict(list)
                    for row in snapshot["rows"]:
                        if row[0] == cutlist:
                            candidates[_db_signature(row)].append(row)
                    export_counts = Counter(_export_signature(source, r) for r in source.rows)
                    if export_counts != Counter({k: len(v) for k, v in candidates.items()}):
                        raise TransferError(f"{source.sheet.title}: workbook pieces differ from the database; export a fresh copy.")
                    for signature, rows in candidates.items():
                        variants = {digest(_item_operations(snapshot, r[0], r[1])) for r in rows}
                        if len(variants) != 1:
                            raise TransferError(f"{source.sheet.title}: identical pieces have different extra operations; source identity is ambiguous.")
                    used = Counter()
                    for row_number in source.rows:
                        signature = _export_signature(source, row_number)
                        row = candidates[signature][used[signature]]
                        used[signature] += 1
                        operations = _item_operations(snapshot, cutlist, row[1])
                        _validate_operations(operations)
                        mapping.append({"sheet": source.sheet.title, "row": row_number,
                                        "original_unit_id": row[2], "item": row[1], "cutlist": cutlist,
                                        "operations": operations})
                result["sources"].append({"path": path, "file_hash": file_hash(file.path), "quote": number,
                                          "snapshot_hash": digest(snapshot), "size_mode": snapshot["header"][3],
                                          "settings": snapshot["cutlists"], "mapping": mapping})
            except (TransferError, OSError) as exc:
                result["errors"][path] = str(exc)
    return result


def operation_counts(inspection):
    counts = Counter()
    for source in inspection["sources"]:
        for item in source["mapping"]:
            for table, rows in item["operations"].items():
                counts[table] += len(rows)
    return dict(counts)


def _key():
    # Existing encrypted app configuration supplies the per-device signing key.
    return Path(KEY_PATH).read_bytes()


def _write_signed(path, payload, *, replace=True):
    envelope = {"payload": payload, "signature": hmac.new(_key(), _json(payload).encode(), hashlib.sha256).hexdigest()}
    path = Path(path)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as handle:
            temporary = Path(handle.name)
            json.dump(envelope, handle, indent=2, ensure_ascii=True)
            handle.flush()
            os.fsync(handle.fileno())
        if replace:
            os.replace(temporary, path)
        else:
            os.link(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def read_manifest(path, config):
    _require_developer(config)
    envelope = json.loads(Path(path).read_text(encoding="utf-8"))
    payload = envelope["payload"]
    signature = hmac.new(_key(), _json(payload).encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(signature, envelope.get("signature", "")):
        raise TransferError("The transfer record was changed or belongs to another device.")
    if payload.get("version") != 1 or payload.get("database") != database_identity(config):
        raise TransferError("The transfer record does not match the configured database.")
    workbook = Path(path).parent / payload["workbook"]
    if workbook.parent.resolve() != Path(path).parent.resolve():
        raise TransferError("The transfer record contains an invalid workbook location.")
    if not workbook.is_file():
        raise TransferError(f"The merged workbook was not found: {workbook}. Keep it beside the transfer record.")
    if file_hash(workbook) != payload["workbook_hash"]:
        _validate_rewritten_workbook(workbook, payload, config)
    return payload


def _workbook_content_hash(file):
    """Hash import cell values, independently of ZIP timestamps and Excel styles."""
    content = []
    for source in file.cutlists:
        cells = []
        piece_rows = set(source.rows)
        for row in source.sheet.iter_rows(max_col=11):
            for cell in row:
                if (cell.row in piece_rows and cell.column in (9, 10)
                        and not Decimal(_text(source.sheet.cell(cell.row, 5).value) or "0")
                        and not Decimal(_text(source.sheet.cell(cell.row, 6).value) or "0")):
                    # Cutting Manager can replace unused edging on unedged pieces.
                    # It has no cutting effect; the signed piece signature also omits it.
                    continue
                value = _text(cell.value)
                if not value:
                    continue
                numeric = ((cell.row >= 8 and cell.column <= 8)
                           or cell.coordinate in ("J4", "I6"))
                if numeric:
                    try:
                        number = Decimal(value)
                        if number.is_finite():
                            value = format(number.normalize(), "f")
                    except ArithmeticError:
                        pass
                cells.append([cell.coordinate, value.casefold()])
        content.append([source.sheet.title.casefold(), cells])
    return digest(content)


def _validate_rewritten_workbook(path, manifest, config):
    """Allow import-time re-saving and verified EID substitutions, never piece edits.

    Legacy v5.3.0 records contain complete signed piece signatures. New records
    additionally bind all import cell values. Target validation remains mandatory.
    This function only reads the workbook and, for EID tokens, source DB records.
    """
    try:
        file = read_merge_file(path)
    except (ValueError, OSError) as exc:
        raise TransferError(f"The merged workbook cannot be verified: {exc}") from exc
    actual = [(source, row) for source in file.cutlists for row in source.rows]
    expected = manifest["rows"]
    if len(actual) != len(expected):
        raise TransferError("The merged workbook's piece count changed; create a new merge.")
    expected_lists = list(dict.fromkeys(row["cutlist"] for row in expected))
    if [int(s.sheet.title[7:]) for s in file.cutlists] != expected_lists:
        raise TransferError("The merged workbook's cutlist order changed; create a new merge.")
    for source in file.cutlists:
        if (_text(source.sheet["C2"].value).casefold() != manifest["customer"].casefold()
                or _text(source.sheet["C3"].value).casefold() != manifest["job"].casefold()):
            raise TransferError(f"{source.sheet.title}: merged customer/job changed; create a new merge.")

    tokens = []
    for (source, row), original in zip(actual, expected):
        if int(source.sheet.title[7:]) != original["cutlist"]:
            raise TransferError("The merged workbook's cutlist pieces changed; create a new merge.")
        edge = source.sheet.cell(row, 10)
        token = re.fullmatch(r"EID_(\d+)", _text(edge.value), re.IGNORECASE)
        if token:
            tokens.append((source, edge, int(token[1]), original))
    for source in file.cutlists:
        cell = source.sheet["E5"]
        token = re.fullmatch(r"EID_(\d+)", _text(cell.value), re.IGNORECASE)
        if token:
            tokens.append((source, cell, int(token[1]), None))
    if tokens:
        with readonly(config) as cursor:
            _schema(cursor)
            snapshots = {}
            for source, cell, edge_id, original in tokens:
                if original is not None:
                    number = original["source_quote"]
                    if number not in snapshots:
                        snapshot = _quote(cursor, number)
                        records = [s for s in manifest["sources"] if s["quote"] == number]
                        if not records or any(digest(snapshot) != s["snapshot_hash"] for s in records):
                            raise TransferError(f"Source quote {number} changed after the merge. Create a fresh merge.")
                        snapshots[number] = {(r[0], r[1]): r for r in snapshot["rows"]}
                    key = (int(original["source_sheet"][7:]), original["source_item"])
                    piece = snapshots[number].get(key)
                    if piece is None or ((piece[6] or piece[7]) and piece[12] != edge_id):
                        raise TransferError(f"{source.sheet.title}, row {cell.row}: imported edging ID differs from the original piece.")
                cursor.execute("SELECT EDGE_NAME FROM EDGE_TYPES WHERE EDGE_ID=?", (edge_id,))
                names = cursor.fetchall()
                if len(names) != 1:
                    raise TransferError(f"{source.sheet.title}, row {cell.row}: imported edging ID cannot be verified.")
                # Only normalize our in-memory copy; never save over the imported workbook.
                cell.value = names[0][0]

    for (source, row), original in zip(actual, expected):
        if _json(_export_signature(source, row)) != _json(original["signature"]):
            raise TransferError(f"{source.sheet.title}, row {row}: merged piece data changed; create a new merge.")
    content_hash = manifest.get("workbook_content_hash")
    if content_hash and _workbook_content_hash(file) != content_hash:
        raise TransferError("The merged workbook's headings, controls or import data changed; create a new merge.")


def prepare_manifest(groups, inspection, customer, job):
    if inspection["errors"]:
        raise TransferError("Every source must be verified before extra operations can be transferred.")
    sources = {s["path"]: s for s in inspection["sources"]}
    if set(sources) != {str(s.path) for g in groups for s in g.sources}:
        raise TransferError("Source files changed after the database check. Check them again.")
    modes = {s["size_mode"] for s in sources.values()}
    if len(modes) != 1 or None in modes:
        raise TransferError("Source quotes use different or unknown dimension modes; transfer requires manual review.")
    settings = []
    rows, unit_ids = [], {}
    uid = 0
    for index, group in enumerate(groups, 1):
        group_settings = {tuple(row[2:]) for source in group.sources
                          for row in sources[str(source.path)]["settings"]
                          if row[0] == int(source.sheet.title[7:])}
        if len(group_settings) != 1:
            raise TransferError("Combined boards have different database cutting settings; transfer requires manual review.")
        settings.append([index, *next(iter(group_settings))])
        for source in group.sources:
            inspected = sources[str(source.path)]
            if file_hash(source.path) != inspected["file_hash"]:
                raise TransferError(f"{source.path.name} changed after checking; reload it.")
            mapping = {(m["sheet"], m["row"]): m for m in inspected["mapping"]}
            for row in source.rows:
                uid += 1
                unit_ids[(str(source.path), source.sheet.title, row)] = uid
                original = mapping[(source.sheet.title, row)]
                rows.append({"uid": uid, "cutlist": index, "signature": _export_signature(source, row, unit_id=uid),
                             "source_path": str(source.path), "source_sheet": source.sheet.title,
                             "source_quote": inspected["quote"], "source_item": original["item"],
                             "source_row": row, "original_unit_id": original["original_unit_id"],
                             "operations": original["operations"]})
    payload = {"version": 1, "id": uuid.uuid4().hex, "created": datetime.now(timezone.utc).isoformat(),
               "database": inspection["database"], "high_water": inspection["high_water"],
               "sources": inspection["sources"], "customer": customer.strip(), "job": job.strip(), "rows": rows,
               "size_mode": next(iter(modes)), "settings": settings}
    return payload, unit_ids


def save_manifest(payload, workbook_path):
    path = Path(workbook_path).with_suffix(".transfer.json")
    if path.exists():
        raise TransferError("A transfer record already exists at this filename. Choose a new workbook filename.")
    payload = {**payload, "workbook": Path(workbook_path).name, "workbook_hash": file_hash(workbook_path),
               "workbook_content_hash": _workbook_content_hash(read_merge_file(workbook_path))}
    _write_signed(path, payload, replace=False)
    return path


def _plan(cursor, manifest, number):
    _schema(cursor)
    if number <= manifest["high_water"] or number in {s["quote"] for s in manifest["sources"]}:
        raise TransferError("Use the new quote created after this merge; existing/source quotes cannot be modified.")
    for source in manifest["sources"]:
        if digest(_quote(cursor, source["quote"])) != source["snapshot_hash"]:
            raise TransferError(f"Source quote {source['quote']} changed after the merge. Create a fresh merge.")
    target = _quote(cursor, number)
    if target["header"][3] != manifest["size_mode"]:
        raise TransferError("The imported quote uses a different dimension mode from the source quotes.")
    review_differences = []
    settings = {row[0]: row[1:] for row in manifest["settings"]}
    for cutlist, status, mode, build in target["cutlists"]:
        original = settings.get(cutlist)
        if original is None or mode != original[0]:
            raise TransferError(f"CutList{cutlist}: imported optimization mode differs from the sources; cutting settings need review in Cutting Manager.")
        if build != original[1]:
            # A new import resets Build_Type/Status independently of piece geometry.
            # Accept only the observed fresh-import default, with explicit review.
            if build != "3" or original[1] not in (None, "2"):
                raise TransferError(f"CutList{cutlist}: unrecognized build-type difference; cutting settings require manual review.")
            review_differences.append({"field": "Build type", "cutlist": cutlist,
                                       "source": original[1], "target": build})
    if _text(target["header"][2]).casefold() != manifest["customer"].casefold():
        raise TransferError("The new quote's customer/job do not match the merged workbook: customer differs.")
    if _text(target["header"][1]).casefold() != manifest["job"].casefold():
        if not _text(target["header"][1]):
            raise TransferError("The imported quote needs a job name before its identity can be reviewed.")
        review_differences.append({"field": "Job name", "source": manifest["job"],
                                   "target": _text(target["header"][1])})
    if any(row[1] != 0 for row in target["cutlists"]):
        raise TransferError("The target cutlists must be newly imported and not optimized/processed.")
    if {row[0] for row in target["cutlists"]} != {row["cutlist"] for row in manifest["rows"]}:
        raise TransferError("The target quote has different cutlists from the merged workbook.")
    by_uid = {r[2]: r for r in target["rows"]}
    if len(by_uid) != len(target["rows"]) or len(by_uid) != len(manifest["rows"]):
        raise TransferError("Target piece IDs are missing, repeated or changed. Import the exact merged workbook.")
    inserts = []
    for expected in manifest["rows"]:
        row = by_uid.get(expected["uid"])
        if row is None or row[0] != expected["cutlist"] or _json(_db_signature(row)) != _json(expected["signature"]):
            raise TransferError(f"Target piece ID {expected['uid']} does not match the merged workbook.")
        _validate_operations(expected["operations"])
        for table, operations in expected["operations"].items():
            for operation in operations:
                inserts.append({"table": table, "uid": expected["uid"], "cutlist": row[0],
                                "item": row[1], "values": operation, "source_quote": expected["source_quote"]})
    planned = {table: sorted([[i["cutlist"], i["item"], *i["values"]] for i in inserts if i["table"] == table], key=_json) for table in OPERATIONS}
    existing = target["operations"]
    applied = existing == planned
    if any(existing.values()) and not applied:
        raise TransferError("The target already contains different or partial extra operations. Nothing will be replaced.")
    return {"manifest_id": manifest["id"], "quote": number, "header": target["header"],
            "target_hash": digest(target), "inserts": inserts, "already_applied": applied,
            "database": manifest["database"], "manifest_hash": digest(manifest),
            "review_differences": review_differences}


def check_target(manifest_path, quote_number, config):
    manifest = read_manifest(manifest_path, config)
    if not isinstance(quote_number, int) or isinstance(quote_number, bool) or quote_number <= 0:
        raise TransferError("Enter a valid new quote number.")
    with readonly(config) as cursor:
        return _plan(cursor, manifest, quote_number)


def _run_gbak(arguments, config):
    env = {**os.environ, "ISC_USER": config["username"], "ISC_PASSWORD": config["password"]}
    options = {"creationflags": subprocess.CREATE_NO_WINDOW} if os.name == "nt" else {}
    result = subprocess.run([firebird_tool("gbak.exe"), *arguments], env=env,
                            capture_output=True, text=True, errors="replace", timeout=300, **options)
    if result.returncode:
        raise TransferError("Firebird could not create/restore a verified backup: " + result.stderr[-1800:])


def create_verified_backup(manifest_path, plan, config, folder):
    """Logical backup, restored only to a newly created temporary database."""
    manifest = read_manifest(manifest_path, config)
    if plan["manifest_hash"] != digest(manifest) or plan["already_applied"]:
        raise TransferError("Check the target quote again before backing up.")
    folder = Path(folder).resolve()
    folder.mkdir(parents=True, exist_ok=True)
    backup = folder / f"before-transfer-{plan['quote']}-{uuid.uuid4().hex}.fbk"
    service = None
    try:
        if config.get("ip"):
            fdb.load_api(config.get("fb_library") or firebird_tool("fbclient.dll"))
            service = fdb.services.connect(host=f"{config['ip']}/{config.get('port') or '3050'}",
                                           user=config["username"], password=config["password"])
            with backup.open("xb") as stream:
                service.local_backup(config["filepath"], stream, collect_garbage=0, no_db_triggers=1)
                stream.flush()
                os.fsync(stream.fileno())
        else:
            _run_gbak(["-backup_database", "-garbage_collect", str(config["filepath"]), str(backup)], config)
        if not backup.is_file() or backup.stat().st_size == 0:
            raise TransferError("The backup is empty; transfer is blocked.")
        with tempfile.TemporaryDirectory(prefix="cutting-verify-") as temporary:
            restored = Path(temporary).resolve() / "verified-copy.fdb"
            # -create_database refuses an existing destination. Never restore over a real DB.
            _run_gbak(["-create_database", str(backup), str(restored)],
                      {"username": "SYSDBA", "password": "isolated-verification-only"})
            local_config = {**config, "ip": "", "filepath": str(restored),
                            "username": "SYSDBA", "password": "isolated-verification-only"}
            with readonly(local_config) as cursor:
                restored_plan = _plan(cursor, manifest, plan["quote"])
            if digest(restored_plan) != digest(plan):
                raise TransferError("The database changed during backup. Check the new quote again.")
        receipt = {"manifest_id": manifest["id"], "plan_hash": digest(plan),
                   "database": manifest["database"], "path": str(backup), "sha256": file_hash(backup),
                   "verified_at": datetime.now(timezone.utc).isoformat()}
        receipt_path = backup.with_suffix(".verified.json")
        _write_signed(receipt_path, receipt)
        return str(receipt_path)
    finally:
        if service is not None:
            service.close()


def _read_backup(receipt_path, plan):
    envelope = json.loads(Path(receipt_path).read_text(encoding="utf-8"))
    receipt = envelope["payload"]
    expected = hmac.new(_key(), _json(receipt).encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected, envelope.get("signature", "")) or receipt["plan_hash"] != digest(plan):
        raise TransferError("This backup was not verified for the reviewed transfer.")
    age = datetime.now(timezone.utc) - datetime.fromisoformat(receipt["verified_at"])
    if age < timedelta(0) or age > timedelta(minutes=15) or file_hash(receipt["path"]) != receipt["sha256"]:
        raise TransferError("The verified backup expired or changed. Create a fresh backup.")
    return receipt


def apply_transfer(manifest_path, plan, receipt_path, config, *, confirmed_quote, cutting_manager_closed,
                   confirmed_review_hash=None):
    """All-or-nothing insertion after locked revalidation; never edit existing rows."""
    manifest = read_manifest(manifest_path, config)
    if confirmed_quote != plan["quote"] or not cutting_manager_closed:
        raise TransferError("Confirm the target quote and close Cutting Manager before applying additions.")
    differences = plan.get("review_differences", [])
    if differences and confirmed_review_hash != digest(differences):
        raise TransferError("Review and explicitly accept the displayed import differences before transferring operations.")
    receipt = _read_backup(receipt_path, plan)
    if plan["manifest_hash"] != digest(manifest) or receipt["database"] != database_identity(config):
        raise TransferError("The reviewed transfer does not match this database/manifest.")
    journal_path = Path(manifest_path).with_suffix(f".{uuid.uuid4().hex}.audit.json")
    journal = {"manifest_id": manifest["id"], "plan": plan, "backup": receipt,
               "status": "prepared", "time": datetime.now(timezone.utc).isoformat()}
    _write_signed(journal_path, journal)
    connection = _connect(config, write=True)
    committing = False
    try:
        tpb = fdb.TPB()
        tpb.isolation_level = fdb.isc_tpb_consistency
        tpb.lock_resolution = fdb.isc_tpb_nowait
        for table in LOCK_TABLES:
            tpb.table_reservation[table] = (fdb.isc_tpb_protected, fdb.isc_tpb_lock_write)
        connection.begin(tpb=tpb)
        cursor = connection.cursor()
        fresh = _plan(cursor, manifest, plan["quote"])
        if fresh["already_applied"]:
            connection.rollback()
            return "already_applied"
        if digest(fresh) != digest(plan):
            raise TransferError("The database changed since review. Nothing was added; check the quote again.")
        for item in fresh["inserts"]:
            table = item["table"]
            fields = ("QUOTE_NR", "CUTLIST_ID", "ITEM_ID", *OPERATIONS[table])
            values = (plan["quote"], item["cutlist"], item["item"], *item["values"])
            cursor.execute(f"INSERT INTO {table} ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})", values)
        verified = _plan(cursor, manifest, plan["quote"])
        if not verified["already_applied"]:
            raise TransferError("Inserted operations did not match the transfer record; all additions will be rolled back.")
        committing = True
        connection.commit()
    except Exception as exc:
        if committing:
            raise CommitOutcomeUnknown("The connection failed during commit. Use Check quote to determine whether it completed; do not retry blindly.") from exc
        connection.rollback()
        raise
    finally:
        # Cleanup errors must not hide a successful commit or an unknown outcome.
        try:
            connection.close()
        except Exception:
            pass
    # If receipt writing fails after commit, retain the prepared journal for recovery.
    try:
        _write_signed(journal_path, {**journal, "status": "committed"})
    except OSError:
        return "committed_audit_pending"
    return "committed"
