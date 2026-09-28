"""只读获取指定日期、站点的人民币汇率；不记录连接凭据。"""
from datetime import date
from decimal import Decimal
from pathlib import Path
import re

DEFAULT_CONFIG = Path(r"C:\codex_projects\log_working\.venv\password.md")


def read_config(path: Path) -> dict:
    try:
        text = Path(path).read_text(encoding="utf-8-sig")
        values = {}
        for line in text.splitlines():
            match = re.match(r"\s*[-*]?\s*([\w\u4e00-\u9fff]+)\s*[:=：]\s*(.*)", line)
            if match:
                values[match[1]] = match[2].strip().strip("`")
        result = {"host": values["主机"], "user": values["用户名"], "password": values["密码"]}
        result["port"] = int(values.get("端口", "3306"))
        if values.get("数据库"):
            result["database"] = values["数据库"]
        return result
    except (OSError, KeyError, ValueError):
        raise ValueError("数据库配置读取失败：需要主机、用户名、密码，可选端口和数据库") from None


def fetch_rates(day: date, countries: set[str], config_path: Path = DEFAULT_CONFIG) -> dict[str, Decimal]:
    if not countries:
        return {}
    config = read_config(config_path)
    try:
        import pymysql
    except ImportError:
        raise ValueError("价格计算需要 PyMySQL，请运行 python -m pip install -r requirements-price.txt") from None
    connection = None
    try:
        connection = pymysql.connect(**config, charset="utf8mb4", connect_timeout=10,
                                     read_timeout=15, write_timeout=15, autocommit=False)
        with connection.cursor() as cursor:
            if "database" not in config:
                cursor.execute("SELECT TABLE_SCHEMA FROM information_schema.TABLES WHERE TABLE_NAME=%s", ("exchange_rates",))
                schemas = [row[0] for row in cursor.fetchall()]
                if len(schemas) != 1:
                    raise ValueError("无法唯一定位 exchange_rates，请在配置文件指定数据库")
                connection.select_db(schemas[0])
            placeholders = ",".join(["%s"] * len(countries))
            cursor.execute("SELECT country,currency_pair,exchange_rate FROM exchange_rates "
                           "WHERE record_date=%s AND country IN (" + placeholders + ")",
                           (day, *sorted(countries)))
            rows = cursor.fetchall()
    except pymysql.MySQLError:
        raise ValueError("MySQL 汇率读取失败，请检查连接、权限和 exchange_rates 表；凭据及服务器错误详情已隐藏") from None
    finally:
        if connection is not None:
            connection.close()
    rates = {}
    for country, pair, value in rows:
        if country not in countries:
            continue
        if not re.fullmatch(r"[A-Z]{3}/CNY", str(pair).strip().upper()):
            raise ValueError(f"{country} 的兑换方向不是外币/CNY，无法直接相乘")
        if country in rates:
            raise ValueError(f"{country} 在 {day} 存在多条汇率，无法唯一匹配")
        rate = Decimal(str(value))
        if not rate.is_finite() or rate <= 0:
            raise ValueError(f"{country} 汇率必须为正数")
        rates[country] = rate
    return rates
