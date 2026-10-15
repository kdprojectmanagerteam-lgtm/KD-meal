import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, DeclarativeBase

# Cloud Run 用環境變數連 Cloud SQL，本地開發用 SQLite
DB_CONNECTION = os.getenv("DB_CONNECTION")

if DB_CONNECTION:
    SQLALCHEMY_DATABASE_URL = DB_CONNECTION
    connect_args = {}
else:
    SQLALCHEMY_DATABASE_URL = "sqlite:///./meal_order.db"
    connect_args = {"check_same_thread": False}

engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class Base(DeclarativeBase):
    pass


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
