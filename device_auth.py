import json
import os
import urllib.error
import urllib.request
import winreg

from dataclasses import dataclass
from datetime import datetime, time, timedelta, timezone

DEFAULT_AUTHORIZATION_URL = (
    "https://raw.githubusercontent.com/LianJordaan/Cutting-Generator/refs/heads/master/authorized_machines.json"
)
AUTHORIZATION_URL = os.environ.get("CUTTING_GENERATOR_AUTH_URL", DEFAULT_AUTHORIZATION_URL)
AUTH_CACHE_DIR = os.path.join(os.path.expanduser("~"), ".cutting_generator")
AUTH_CACHE_PATH = os.path.join(AUTH_CACHE_DIR, "device_auth.json")
AUTH_CACHE_WINDOW = timedelta(hours=24)
INVALID_LICENSE_MESSAGE = "You do not have a valid licence. Please contact ByteBuilders Hosting to have it resolved."


@dataclass
class AuthorizationResult:
    allowed: bool
    machine_guid: str
    message: str
    checked_remotely: bool
    valid_until_utc: str | None = None


@dataclass
class MachineLicenseEntry:
    machine_guid: str
    expires_utc: datetime | None = None
    comment: str = ""


def normalize_machine_guid(machine_guid: str) -> str:
    return str(machine_guid).strip().lower()


def get_machine_guid() -> str:
    if os.name != "nt":
        raise RuntimeError("Machine GUID authorization is only supported on Windows.")

    access_modes = [winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0), winreg.KEY_READ]
    last_error = None

    for access_mode in access_modes:
        try:
            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography",
                0,
                access_mode,
            ) as registry_key:
                machine_guid, _ = winreg.QueryValueEx(registry_key, "MachineGuid")
                if machine_guid:
                    return normalize_machine_guid(machine_guid)
        except OSError as exc:
            last_error = exc

    raise RuntimeError("Unable to read the Windows MachineGuid from the registry.") from last_error


def load_auth_cache() -> dict:
    if not os.path.exists(AUTH_CACHE_PATH):
        return {}

    try:
        with open(AUTH_CACHE_PATH, "r", encoding="utf-8") as cache_file:
            return json.load(cache_file)
    except (OSError, json.JSONDecodeError, ValueError):
        return {}


def save_auth_cache(cache_data: dict) -> None:
    os.makedirs(AUTH_CACHE_DIR, exist_ok=True)
    with open(AUTH_CACHE_PATH, "w", encoding="utf-8") as cache_file:
        json.dump(cache_data, cache_file, indent=2, sort_keys=True)


def parse_utc_timestamp(timestamp: str | None) -> datetime | None:
    if not timestamp:
        return None

    normalized_timestamp = timestamp.strip()
    if normalized_timestamp.endswith("Z"):
        normalized_timestamp = normalized_timestamp[:-1] + "+00:00"

    try:
        parsed = datetime.fromisoformat(normalized_timestamp)
    except ValueError:
        return None

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)

    return parsed.astimezone(timezone.utc)


def format_utc_timestamp(timestamp: datetime) -> str:
    return timestamp.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def parse_license_expiry(entry: dict) -> datetime | None:
    for key in ("expires_at_utc", "expires_at", "expires_on", "expiry_date", "expires"):
        expiry_value = entry.get(key)
        if not expiry_value:
            continue

        expiry_text = str(expiry_value).strip()
        if not expiry_text:
            continue

        if len(expiry_text) == 10 and expiry_text.count("-") == 2:
            try:
                expiry_date = datetime.strptime(expiry_text, "%Y-%m-%d").date()
            except ValueError as exc:
                raise ValueError(f"Invalid expiry date '{expiry_text}' for machine GUID entry.") from exc
            return datetime.combine(expiry_date, time.max, tzinfo=timezone.utc)

        parsed_expiry = parse_utc_timestamp(expiry_text)
        if parsed_expiry is None:
            raise ValueError(f"Invalid expiry timestamp '{expiry_text}' for machine GUID entry.")
        return parsed_expiry

    return None


def extract_authorized_machine_entries(payload) -> dict[str, MachineLicenseEntry]:
    authorized_machine_entries: dict[str, MachineLicenseEntry] = {}
    found_supported_structure = False

    def add_machine_guid(entry) -> None:
        if isinstance(entry, str):
            normalized_guid = normalize_machine_guid(entry)
            if normalized_guid:
                authorized_machine_entries[normalized_guid] = MachineLicenseEntry(machine_guid=normalized_guid)
            return

        if not isinstance(entry, dict):
            return

        if entry.get("enabled", True) is False:
            return

        machine_guid = entry.get("machine_guid") or entry.get("machineGuid") or entry.get("id")
        if machine_guid:
            normalized_guid = normalize_machine_guid(machine_guid)
            if not normalized_guid:
                return
            authorized_machine_entries[normalized_guid] = MachineLicenseEntry(
                machine_guid=normalized_guid,
                expires_utc=parse_license_expiry(entry),
                comment=str(entry.get("comment", "")).strip(),
            )

    if isinstance(payload, list):
        found_supported_structure = True
        for item in payload:
            add_machine_guid(item)
    elif isinstance(payload, dict):
        for key in ("authorized_machine_guids", "authorized_devices", "devices"):
            value = payload.get(key)
            if isinstance(value, list):
                found_supported_structure = True
                for item in value:
                    add_machine_guid(item)

    if not found_supported_structure:
        raise ValueError(
            "Authorization allowlist must be a JSON list or an object with authorized_machine_guids/authorized_devices/devices."
        )

    return authorized_machine_entries


def fetch_authorized_machine_entries(authorization_url: str = AUTHORIZATION_URL) -> dict[str, MachineLicenseEntry]:
    request = urllib.request.Request(
        authorization_url,
        headers={"User-Agent": "CuttingGeneratorDeviceAuth"},
    )

    with urllib.request.urlopen(request, timeout=10) as response:
        payload = json.load(response)

    return extract_authorized_machine_entries(payload)


def compute_access_valid_until_utc(now_utc: datetime, expires_utc: datetime | None) -> datetime:
    cache_valid_until_utc = now_utc + AUTH_CACHE_WINDOW
    if expires_utc is None:
        return cache_valid_until_utc
    return min(cache_valid_until_utc, expires_utc)


def is_cached_authorization_valid(
    cache_data: dict,
    machine_guid: str,
    authorization_url: str,
    now_utc: datetime,
) -> bool:
    cached_machine_guid = normalize_machine_guid(cache_data.get("machine_guid", ""))
    cached_authorization_url = cache_data.get("authorization_url")
    cached_access_valid_until_utc = parse_utc_timestamp(cache_data.get("cached_access_valid_until_utc"))

    if (
        cached_machine_guid != machine_guid
        or cached_authorization_url != authorization_url
        or cached_access_valid_until_utc is None
    ):
        return False

    return now_utc < cached_access_valid_until_utc


def ensure_device_is_authorized(authorization_url: str = AUTHORIZATION_URL) -> AuthorizationResult:
    machine_guid = get_machine_guid()
    now_utc = datetime.now(timezone.utc)
    cache_data = load_auth_cache()

    if is_cached_authorization_valid(cache_data, machine_guid, authorization_url, now_utc):
        return AuthorizationResult(
            allowed=True,
            machine_guid=machine_guid,
            message="Cached device licence is still valid.",
            checked_remotely=False,
            valid_until_utc=cache_data.get("cached_access_valid_until_utc"),
        )

    cache_data["machine_guid"] = machine_guid
    cache_data["authorization_url"] = authorization_url
    cache_data["last_checked_at_utc"] = format_utc_timestamp(now_utc)

    try:
        authorized_machine_entries = fetch_authorized_machine_entries(authorization_url)
    except (OSError, ValueError, urllib.error.URLError) as exc:
        cache_data["last_failure_at_utc"] = format_utc_timestamp(now_utc)
        cache_data["last_failure_reason"] = f"Unable to refresh device licence: {exc}"
        save_auth_cache(cache_data)
        return AuthorizationResult(
            allowed=False,
            machine_guid=machine_guid,
            message=INVALID_LICENSE_MESSAGE,
            checked_remotely=True,
        )

    matching_entry = authorized_machine_entries.get(machine_guid)

    if matching_entry is None:
        cache_data["last_failure_at_utc"] = format_utc_timestamp(now_utc)
        cache_data["last_failure_reason"] = "Machine GUID is not present in the licence allowlist."
        save_auth_cache(cache_data)
        return AuthorizationResult(
            allowed=False,
            machine_guid=machine_guid,
            message=INVALID_LICENSE_MESSAGE,
            checked_remotely=True,
        )

    if matching_entry.expires_utc is not None and now_utc >= matching_entry.expires_utc:
        cache_data["last_failure_at_utc"] = format_utc_timestamp(now_utc)
        cache_data["last_failure_reason"] = (
            f"Machine licence expired at {format_utc_timestamp(matching_entry.expires_utc)}."
        )
        cache_data["licensed_until_utc"] = format_utc_timestamp(matching_entry.expires_utc)
        cache_data["license_comment"] = matching_entry.comment or None
        save_auth_cache(cache_data)
        return AuthorizationResult(
            allowed=False,
            machine_guid=machine_guid,
            message=INVALID_LICENSE_MESSAGE,
            checked_remotely=True,
        )

    access_valid_until_utc = compute_access_valid_until_utc(now_utc, matching_entry.expires_utc)

    cache_data["last_successful_validation_at_utc"] = format_utc_timestamp(now_utc)
    cache_data["cached_access_valid_until_utc"] = format_utc_timestamp(access_valid_until_utc)
    cache_data["licensed_until_utc"] = (
        format_utc_timestamp(matching_entry.expires_utc) if matching_entry.expires_utc is not None else None
    )
    cache_data["license_comment"] = matching_entry.comment or None
    cache_data["last_failure_at_utc"] = None
    cache_data["last_failure_reason"] = None
    save_auth_cache(cache_data)

    return AuthorizationResult(
        allowed=True,
        machine_guid=machine_guid,
        message="Device licence refreshed successfully.",
        checked_remotely=True,
        valid_until_utc=format_utc_timestamp(access_valid_until_utc),
    )