#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager, rcParams

EXCLUDED_CATEGORIES = {"금융/투자"}


def _set_korean_font() -> None:
    preferred = ["AppleGothic", "Malgun Gothic", "NanumGothic"]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in preferred:
        if name in available:
            rcParams["font.family"] = name
            break
    rcParams["axes.unicode_minus"] = False


def _plot_pie(series: pd.Series, title: str, out_path: Path) -> None:
    values = series.values.astype(float)
    labels = series.index.astype(str).tolist()
    total = float(series.sum())
    legend_labels = [
        f"{label}: {(value / total) * 100:.1f}% ({value:,.0f}원)" for label, value in zip(labels, values)
    ]

    fig_height = max(8.0, 4.0 + 0.30 * len(labels))
    fig, ax = plt.subplots(figsize=(15, fig_height))
    wedges, _ = ax.pie(
        values,
        labels=None,
        startangle=90,
        counterclock=False,
        wedgeprops={"linewidth": 1, "edgecolor": "white"},
    )
    ax.legend(
        wedges,
        legend_labels,
        title="세부 카테고리",
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=False,
        fontsize=9,
        title_fontsize=11,
    )
    ax.set_title(f"{title}\n총액: {total:,.0f}원", fontsize=14)
    ax.axis("equal")
    fig.subplots_adjust(right=0.62)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200)
    plt.close(fig)


def main() -> None:
    _set_korean_font()

    src = Path("통합지출내역.xlsx")
    out_path = Path("charts/13_세부카테고리_전체_파이차트.png")
    out_path.parent.mkdir(exist_ok=True)

    if not src.exists():
        raise FileNotFoundError(f"입력 파일이 없습니다: {src}")

    df = pd.read_excel(src)
    required = ["지출금액", "카테고리", "세부 카테고리"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}")

    df["지출금액"] = pd.to_numeric(df["지출금액"], errors="coerce")
    df["카테고리"] = df["카테고리"].fillna("").astype(str).str.strip()
    df["세부 카테고리"] = df["세부 카테고리"].fillna("").astype(str).str.strip()

    df = df[df["지출금액"].notna() & (df["지출금액"] > 0)]
    df = df[df["카테고리"] != ""]
    df = df[df["카테고리"].str.lower() != "nan"]
    df = df[~df["카테고리"].isin(EXCLUDED_CATEGORIES)]
    df.loc[df["세부 카테고리"] == "", "세부 카테고리"] = "미분류"

    sub_sum = df.groupby("세부 카테고리", as_index=True)["지출금액"].sum().sort_values(ascending=False)
    _plot_pie(sub_sum, "세부 카테고리별 지출 비중", out_path)
    print(out_path.as_posix())


if __name__ == "__main__":
    main()
