import os
import prestodb
import pandas as pd


def get_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


PRESTO_HOST = os.getenv("PRESTO_HOST", "presto.dbc.ludp.lenovo.com")
PRESTO_PORT = int(os.getenv("PRESTO_PORT", "30060"))
PRESTO_USER = get_required_env("ITcode")
PRESTO_PASSWORD = get_required_env("Password")
PRESTO_CERT_PATH = get_required_env("PRESTO_CERT_PATH")


def get_presto_conn():
    conn = prestodb.dbapi.connect(
        host=PRESTO_HOST,
        port=PRESTO_PORT,
        user=PRESTO_USER,
        catalog="hive",
        schema="default",
        http_scheme="https",
        auth=prestodb.auth.BasicAuthentication(
            PRESTO_USER,
            PRESTO_PASSWORD
        )
    )

    conn._http_session.verify = PRESTO_CERT_PATH
    return conn


def query_df(sql: str) -> pd.DataFrame:
    conn = get_presto_conn()
    try:
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        cols = [c[0] for c in cur.description]
        return pd.DataFrame(rows, columns=cols)
    finally:
        conn.close()
