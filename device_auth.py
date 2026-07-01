import json
import os
import urllib.error
import urllib.request
import winreg

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone

DEFAULT_AUTHORIZATION_URL = (
    "https://raw.githubusercontent.com/LianJordaan/Cutting-Generator/refs/heads/master/authorized_machines.json"
)
AUTHORIZATION_URL = os.environ.get("CUTTING_GENERATOR_AUTH_URL", DEFAULT_AUTHORIZATION_URL)
AUTH_CACHE_DIR = os.path.join(os.path.expanduser("~"), ".cutting_generator")
AUTH_CACHE_PATH = os.path.join(AUTH_CACHE_DIR, "device_auth.json")
AUTH_CACHE_WINDOW = timedelta(hours=24)


@dataclass
class AuthorizationResult:
    allowed: bool
    machine_guid: str
    message: str
    checked_remotely: bool
    valid_until_utc: str | None = None


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


def extract_authorized_machine_guids(payload) -> set[str]:
    authorized_machine_guids = set()
    found_supported_structure = False

    def add_machine_guid(entry) -> None:
        if isinstance(entry, str):
            normalized_guid = normalize_machine_guid(entry)
            if normalized_guid:
                authorized_machine_guids.add(normalized_guid)
            return

        if not isinstance(entry, dict):
            return

        if entry.get("enabled", True) is False:
            return

        machine_guid = entry.get("machine_guid") or entry.get("machineGuid") or entry.get("id")
        if machine_guid:
            authorized_machine_guids.add(normalize_machine_guid(machine_guid))

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
            "Authorization allowlist must be a JSON list or an object with authorized_machine_guids/authorized_devices."
        )

    return authorized_machine_guids


def fetch_authorized_machine_guids(authorization_url: str = AUTHORIZATION_URL) -> set[str]:
    request = urllib.request.Request(
        authorization_url,
        headers={"User-Agent": "CuttingGeneratorDeviceAuth"},
    )

    with urllib.request.urlopen(request, timeout=10) as response:
        payload = json.load(response)

    return extract_authorized_machine_guids(payload)


def is_cached_authorization_valid(
    cache_data: dict,
    machine_guid: str,
    authorization_url: str,
    now_utc: datetime,
) -> bool:
    cached_machine_guid = normalize_machine_guid(cache_data.get("machine_guid", ""))
    cached_authorization_url = cache_data.get("authorization_url")
    last_validation_utc = parse_utc_timestamp(cache_data.get("last_successful_validation_at_utc"))

    if (
        cached_machine_guid != machine_guid
        or cached_authorization_url != authorization_url
        or last_validation_utc is None
    ):
        return False

    return now_utc - last_validation_utc < AUTH_CACHE_WINDOW


def ensure_device_is_authorized(authorization_url: str = AUTHORIZATION_URL) -> AuthorizationResult:
    machine_guid = get_machine_guid()
    now_utc = datetime.now(timezone.utc)
    cache_data = load_auth_cache()

    if is_cached_authorization_valid(cache_data, machine_guid, authorization_url, now_utc):
        last_validation_utc = parse_utc_timestamp(cache_data.get("last_successful_validation_at_utc"))
        valid_until_utc = format_utc_timestamp(last_validation_utc + AUTH_CACHE_WINDOW)
        return AuthorizationResult(
            allowed=True,
            machine_guid=machine_guid,
            message="Cached device authorization is still valid.",
            checked_remotely=False,
            valid_until_utc=valid_until_utc,
        )

    cache_data["machine_guid"] = machine_guid
    cache_data["authorization_url"] = authorization_url
    cache_data["last_checked_at_utc"] = format_utc_timestamp(now_utc)

    try:
        authorized_machine_guids = fetch_authorized_machine_guids(authorization_url)
    except (OSError, ValueError, urllib.error.URLError) as exc:
        cache_data["last_failure_at_utc"] = format_utc_timestamp(now_utc)
        cache_data["last_failure_reason"] = f"Unable to refresh device authorization: {exc}"
        save_auth_cache(cache_data)
        return AuthorizationResult(
            allowed=False,
            machine_guid=machine_guid,
            message=(
                f"Unable to refresh device authorization from {authorization_url}. "
                f"A successful validation is required at least once every 24 hours. Error: {exc}"
            ),
            checked_remotely=True,
        )

    if machine_guid not in authorized_machine_guids:
        cache_data["last_failure_at_utc"] = format_utc_timestamp(now_utc)
        cache_data["last_failure_reason"] = "Machine GUID is not present in the authorization allowlist."
        save_auth_cache(cache_data)
        return AuthorizationResult(
            allowed=False,
            machine_guid=machine_guid,
            message=(
                f"This machine is not authorized in {authorization_url}. "
                f"Add the Machine GUID below to the allowlist and try again."
            ),
            checked_remotely=True,
        )

    cache_data["last_successful_validation_at_utc"] = format_utc_timestamp(now_utc)
    cache_data["last_failure_at_utc"] = None
    cache_data["last_failure_reason"] = None
    save_auth_cache(cache_data)

    return AuthorizationResult(
        allowed=True,
        machine_guid=machine_guid,
        message="Device authorization refreshed successfully.",
        checked_remotely=True,
        valid_until_utc=format_utc_timestamp(now_utc + AUTH_CACHE_WINDOW),
    )