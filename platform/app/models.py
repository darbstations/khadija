# -*- coding: utf-8 -*-
from sqlalchemy import (Column, Integer, String, Float, ForeignKey, UniqueConstraint, Text)
from sqlalchemy.orm import relationship
from .database import Base

class Department(Base):
    __tablename__ = "departments"
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    key = Column(String, unique=True, nullable=False)
    kpis = relationship("KPI", back_populates="department", order_by="KPI.order")

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True)
    username = Column(String, unique=True, nullable=False)
    full_name = Column(String, default="")
    pw_hash = Column(String, nullable=False)
    salt = Column(String, nullable=False)
    role = Column(String, nullable=False)          # admin | executive | manager | employee
    department_id = Column(Integer, ForeignKey("departments.id"), nullable=True)
    department = relationship("Department")

class KPI(Base):
    __tablename__ = "kpis"
    id = Column(Integer, primary_key=True)
    department_id = Column(Integer, ForeignKey("departments.id"))
    order = Column(Integer, default=0)
    axis = Column(String)
    name = Column(String, nullable=False)
    unit = Column(String)
    polarity = Column(String)        # ↑ / ↓
    agg = Column(String)             # SUM / AVG / LAST
    target = Column(Float, nullable=True)
    target_text = Column(String)
    fmt = Column(String)             # pct/int/num1/rial
    pillar = Column(String)
    project = Column(String)
    perspective = Column(String)
    kpitype = Column(String)         # قائد / لاحق
    weight = Column(Float, default=0)
    baseline = Column(Float, nullable=True)
    owner_user_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    department = relationship("Department", back_populates="kpis")
    values = relationship("KPIValue", back_populates="kpi", cascade="all,delete-orphan")

class KPIValue(Base):
    __tablename__ = "kpi_values"
    id = Column(Integer, primary_key=True)
    kpi_id = Column(Integer, ForeignKey("kpis.id"))
    month = Column(Integer)          # 1..12
    actual = Column(Float, nullable=True)
    kpi = relationship("KPI", back_populates="values")
    __table_args__ = (UniqueConstraint("kpi_id", "month", name="uix_kpi_month"),)

class Initiative(Base):
    __tablename__ = "initiatives"
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    pillar = Column(String, default="")
    project = Column(String, default="")
    kpi_link = Column(String, default="")
    owner = Column(String, default="")
    start = Column(String, default="")
    end = Column(String, default="")
    budget = Column(Float, nullable=True)
    progress = Column(Float, default=0)   # 0..1
    notes = Column(Text, default="")

class Risk(Base):
    __tablename__ = "risks"
    id = Column(Integer, primary_key=True)
    title = Column(String, nullable=False)
    area = Column(String, default="")
    likelihood = Column(Integer, default=0)   # 1..3
    impact = Column(Integer, default=0)       # 1..3
    response = Column(Text, default="")
    owner = Column(String, default="")

class Setting(Base):
    __tablename__ = "settings"
    key = Column(String, primary_key=True)
    value = Column(String, default="")
