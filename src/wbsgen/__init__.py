"""wbsgen — 空の WBS (ガントチャート用紙) を作るツール。

期間を指定すると、月と週 (または日 / 月) の日程表と、記入用の空行が並んだ
Excel を書き出す。中身は入れない。
"""

__version__ = "2.5.0"

from .blank import BlankWBS, build  # noqa: E402
from .workbook import write  # noqa: E402

__all__ = ["BlankWBS", "build", "write", "__version__"]
