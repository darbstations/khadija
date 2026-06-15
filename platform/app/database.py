# -*- coding: utf-8 -*-
import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_URL = os.environ.get("DATABASE_URL", "")
if DB_URL:
    # قاعدة بيانات مركزية مشتركة (PostgreSQL على الاستضافة)
    if DB_URL.startswith("postgres://"):
        DB_URL = DB_URL.replace("postgres://", "postgresql://", 1)
    engine = create_engine(DB_URL, pool_pre_ping=True)
else:
    # محلي: SQLite
    DB_PATH = os.environ.get("DARB_DB_PATH", os.path.join(BASE_DIR, "darb.db"))
    os.makedirs(os.path.dirname(os.path.abspath(DB_PATH)), exist_ok=True)
    engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
