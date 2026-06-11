# -*- coding: utf-8 -*-
import hashlib, os
from .models import User, KPI

def hash_password(pw, salt=None):
    salt = salt or os.urandom(16).hex()
    h = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 120000).hex()
    return h, salt

def verify_password(pw, pw_hash, salt):
    h, _ = hash_password(pw, salt)
    return h == pw_hash

ROLE_LABEL = {
    "admin": "مدير النظام (صلاحيات كاملة)",
    "executive": "الإدارة التنفيذية",
    "manager": "مدير إدارة",
    "employee": "موظف مسؤول عن مؤشر",
}

def current_user(request, db):
    uid = request.session.get("uid")
    if not uid: return None
    return db.query(User).get(uid)

def can_edit_definition(user):
    """تعديل المؤشر/المستهدف/الوزن/الهيكل — للأدمن فقط."""
    return user is not None and user.role == "admin"

def can_edit_value(user, kpi: KPI):
    """تغذية الأرقام الشهرية — الأدمن، أو المسؤول عن مؤشره/إدارته فقط (دون المساس بالتعريف)."""
    if user is None: return False
    if user.role == "admin": return True
    if kpi.owner_user_id and kpi.owner_user_id == user.id: return True
    if user.department_id and user.department_id == kpi.department_id and user.role in ("executive","manager","employee"):
        return True
    return False
