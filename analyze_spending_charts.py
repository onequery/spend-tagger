#!/usr/bin/env python3
from __future__ import annotations

import re
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


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name).strip("_")


def _plot_pie(series: pd.Series, title: str, out_path: Path) -> None:
    values = series.values.astype(float)
    labels = series.index.astype(str).tolist()
    total = float(series.sum())
    legend_labels = [
        f"{label}: {(value / total) * 100:.1f}% ({value:,.0f}원)" for label, value in zip(labels, values)
    ]

    fig_height = max(8.0, 4.0 + 0.45 * len(labels))
    fig, ax = plt.subplots(figsize=(14, fig_height))
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
        title="항목",
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=False,
        fontsize=10,
        title_fontsize=11,
    )
    ax.set_title(f"{title}\n총액: {total:,.0f}원", fontsize=14)
    ax.axis("equal")
    fig.subplots_adjust(right=0.65)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200)
    plt.close(fig)


def main() -> None:
    _set_korean_font()

    src = Path("통합지출내역.xlsx")
    if not src.exists():
        raise FileNotFoundError(f"입력 파일이 없습니다: {src}")

    out_dir = Path("charts")
    out_dir.mkdir(exist_ok=True)

    df = pd.read_excel(src)
    required = ["지출날짜", "사용처", "지출금액", "카테고리", "세부 카테고리"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}")

    df["지출금액"] = pd.to_numeric(df["지출금액"], errors="coerce")
    df = df[df["지출금액"].notna() & (df["지출금액"] > 0)].copy()
    df["카테고리"] = df["카테고리"].fillna("").astype(str).str.strip()
    df = df[df["카테고리"] != ""]
    df = df[df["카테고리"].str.lower() != "nan"]
    df = df[~df["카테고리"].isin(EXCLUDED_CATEGORIES)]
    df["세부 카테고리"] = df["세부 카테고리"].fillna("미분류").astype(str).str.strip()
    df.loc[df["세부 카테고리"] == "", "세부 카테고리"] = "미분류"

    category_sum = df.groupby("카테고리", as_index=True)["지출금액"].sum().sort_values(ascending=False)
    _plot_pie(
        category_sum,
        "카테고리별 지출 비중",
        out_dir / "카테고리_파이차트.png",
    )

    created = [out_dir / "카테고리_파이차트.png"]

    for category in category_sum.index:
        sub_sum = (
            df[df["카테고리"] == category]
            .groupby("세부 카테고리", as_index=True)["지출금액"]
            .sum()
            .sort_values(ascending=False)
        )
        file_name = f"세부카테고리_파이차트_{_safe_filename(category)}.png"
        out_path = out_dir / file_name
        _plot_pie(sub_sum, f"{category} - 세부 카테고리별 지출 비중", out_path)
        created.append(out_path)

    print(f"총 {len(created)}개 차트 생성 완료")
    for p in created:
        print(p.as_posix())


if __name__ == "__main__":
    main()
