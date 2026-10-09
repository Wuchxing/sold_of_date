"""按运行开始月份追加 JSONL；不记录异常原文或数据库配置。"""
from datetime import datetime
import json
import os
from pathlib import Path
import time


def now():
    return datetime.now().astimezone()


class RunLog:
    def __init__(self, source, output, all_calculations):
        self.started = now()
        self.timer = time.perf_counter()
        self.stats = {}
        self.output = str(Path(output).resolve()) if output is not None else None
        self.stage = '输入检查'
        self.source = Path(source).resolve()
        self.mode = '日销与价格' if all_calculations else '仅日销'

    def finish(self, error=None):
        ended = now()
        stats = self.stats
        record = dict(
            started_at=self.started.isoformat(), ended_at=ended.isoformat(),
            duration_seconds=round(time.perf_counter() - self.timer, 6),
            input_path=str(self.source), input_filename=self.source.name,
            output_path=self.output, mode=self.mode,
            total_rows=stats.get('total_rows'),
            daily_calculated=stats.get('calculated', 0),
            price_calculated=stats.get('price_calculated', 0),
            sales_corrected=stats.get('corrected', 0),
            daily_skipped_no_history=stats.get('blank', 0),
            daily_skipped_no_sales=stats.get('skipped', 0),
            price_skipped_no_rate=stats.get('price_skipped_rate', 0),
            price_skipped_empty=stats.get('price_skipped_empty', 0),
            status='success' if error is None else 'failed',
            stage=self.stage,
            error=None if error is None else dict(
                type=type(error).__name__,
                message=f'{self.stage}失败；请检查输入数据、文件权限或运行依赖。异常原文未记录。'),
        )
        base = Path(os.environ.get('LOCALAPPDATA') or Path.home() / 'AppData' / 'Local')
        path = base / 'sold_of_date' / 'logs' / f'{self.started:%Y-%m}.jsonl'
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open('a', encoding='utf-8') as stream:
            stream.write(json.dumps(record, ensure_ascii=False) + '\n')
        return path
