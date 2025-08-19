import json
import os
from dataclasses import asdict, dataclass
from typing import Optional, List


USERS_FILE_NAME = "users.txt"


@dataclass
class UserRecord:
    user_id: int
    username: Optional[str]
    first_name: Optional[str]
    second_name: Optional[str]
    master_filename: Optional[str] = None


def _users_file_path() -> str:
    return os.path.join(os.path.dirname(__file__), USERS_FILE_NAME)


def _read_all_users() -> List[UserRecord]:
    path = _users_file_path()
    if not os.path.exists(path):
        return []
    users: List[UserRecord] = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                data = json.loads(line)
                users.append(
                    UserRecord(
                        user_id=int(data.get("user_id")),
                        username=data.get("username"),
                        first_name=data.get("first_name"),
                        second_name=data.get("second_name"),
                        master_filename=data.get("master_filename"),
                    )
                )
            except Exception:
                # Skip malformed lines
                continue
    return users


def _write_all_users(users: List[UserRecord]) -> None:
    path = _users_file_path()
    with open(path, "w", encoding="utf-8") as f:
        for user in users:
            f.write(json.dumps(asdict(user), ensure_ascii=False) + "\n")


def load_user_by_id(user_id: int) -> Optional[UserRecord]:
    for user in _read_all_users():
        if user.user_id == user_id:
            return user
    return None


def upsert_user(record: UserRecord) -> UserRecord:
    users = _read_all_users()
    replaced = False
    for idx, u in enumerate(users):
        if u.user_id == record.user_id:
            users[idx] = record
            replaced = True
            break
    if not replaced:
        users.append(record)
    _write_all_users(users)
    return record


def get_or_create_user_from_telegram(user) -> UserRecord:
    existing = load_user_by_id(user.id)
    if existing:
        return existing
    # Create new user. Use Telegram last_name as second_name if present.
    created = UserRecord(
        user_id=user.id,
        username=getattr(user, "username", None),
        first_name=getattr(user, "first_name", None),
        second_name=getattr(user, "last_name", None),
    )
    return upsert_user(created)


def update_user_second_name(user_id: int, second_name: str) -> UserRecord:
    users = _read_all_users()
    for idx, u in enumerate(users):
        if u.user_id == user_id:
            users[idx].second_name = second_name
            _write_all_users(users)
            return users[idx]
    # If user was not there, create it minimalistically
    created = UserRecord(user_id=user_id, username=None, first_name=None, second_name=second_name)
    return upsert_user(created)


def build_report_filename(second_name: Optional[str]) -> str:
    name = (second_name or "user").strip()
    # Simple sanitization: replace spaces, keep unicode letters/digits
    name = name.replace(" ", "_")
    if not name:
        name = "user"
    return f"finance_report_{name}.xlsx"


def set_master_filename(user_id: int, filename: str) -> UserRecord:
    users = _read_all_users()
    for idx, u in enumerate(users):
        if u.user_id == user_id:
            users[idx].master_filename = filename
            _write_all_users(users)
            return users[idx]
    created = UserRecord(user_id=user_id, username=None, first_name=None, second_name=None, master_filename=filename)
    return upsert_user(created)

