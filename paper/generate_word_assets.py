# -*- coding: utf-8 -*-
"""Generate 300-DPI CVPR/TIP-style minimalist visual figures for PRISM_图像去雾大作业报告_精排终稿.docx."""
from pathlib import Path
import csv
import json
import fitz
import numpy as np
from PIL import Image, ImageDraw
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# Use SimSun / Times New Roman for classic academic aesthetic
plt.rcParams["font.sans-serif"] = ["SimSun", "STSong", "Microsoft YaHei", "Times New Roman"]
plt.rcParams["font.serif"] = ["Times New Roman", "SimSun"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["mathtext.fontset"] = "stix"

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parent
COURSE10 = PROJECT / "course_10"
FIG_DIR = ROOT / "figures"
OUT_DIR = ROOT / "word_assets"
OUT_DIR.mkdir(parents=True, exist_ok=True)

COURSE_DIRS = {
    "input": (COURSE10 / "input", "_original.jpg"),
    "base": (COURSE10 / "output_base", "_original.png"),
    "bilalora": (COURSE10 / "output_bilalora", "_original.png"),
    "ours": (COURSE10 / "output_ours", "_original.png"),
}

INK_BLACK = "#1E293B"


def load_course_metrics():
    data = {}
    for m_key in ["input", "base", "bilalora", "ours"]:
        with (COURSE10 / "metrics" / f"fade_biqme_{m_key}.csv").open(encoding="utf-8") as f:
            for row in csv.DictReader(f):
                stem = row["filename"].split("_original")[0]
                data.setdefault(stem, {}).setdefault(m_key, {})
                data[stem][m_key]["FADE"] = float(row["FADE"])
                data[stem][m_key]["BIQME"] = float(row["BIQME"])
                data[stem][m_key]["Entropy"] = float(row["Entropy_RGB_pooled"])
        with (COURSE10 / "metrics" / f"musiq_{m_key}.csv").open(encoding="utf-8") as f:
            for row in csv.DictReader(f):
                stem = row["filename"].split("_original")[0]
                data[stem][m_key]["MUSIQ"] = float(row["MUSIQ"])
    return data


def get_course_img_path(m_key: str, stem: str) -> Path:
    folder, suffix = COURSE_DIRS[m_key]
    return folder / f"{stem}{suffix}"


COURSE_METRICS = load_course_metrics()


def fit_crop(img: Image.Image, target_w: int, target_h: int) -> Image.Image:
    """Center-crop to match aspect ratio and resize with LANCZOS."""
    w, h = img.size
    target_ratio = target_w / target_h
    cur_ratio = w / h
    if cur_ratio > target_ratio:
        new_w = int(round(h * target_ratio))
        left = (w - new_w) // 2
        img = img.crop((left, 0, left + new_w, h))
    elif cur_ratio < target_ratio:
        new_h = int(round(w / target_ratio))
        top = (h - new_h) // 2
        img = img.crop((0, top, w, top + new_h))
    return img.resize((target_w, target_h), Image.Resampling.LANCZOS)


def format_metric_label(stem: str, m_key: str) -> str:
    """Format FADE / MUSIQ label in black with only the best score across the 4 methods bolded."""
    m = COURSE_METRICS[stem][m_key]
    best_fade = min(COURSE_METRICS[stem][k]["FADE"] for k in ["input", "base", "bilalora", "ours"])
    best_musiq = max(COURSE_METRICS[stem][k]["MUSIQ"] for k in ["input", "base", "bilalora", "ours"])
    f_str = f"\\mathbf{{{m['FADE']:.4f}}}" if abs(m["FADE"] - best_fade) < 1e-6 else f"{m['FADE']:.4f}"
    q_str = f"\\mathbf{{{m['MUSIQ']:.2f}}}" if abs(m["MUSIQ"] - best_musiq) < 1e-4 else f"{m['MUSIQ']:.2f}"
    return f"${f_str}\\ /\\ {q_str}$"


def make_cover_teaser():
    """Generate minimalist Cover Teaser strip: tiananmen & neighbour across 4 methods."""
    scenes = [
        ("tiananmen", "tiananmen"),
        ("neighbour", "neighbour"),
    ]
    methods = [
        ("input", "有雾原图"),
        ("base", "DEA-Net Base"),
        ("bilalora", "BiLaLoRA"),
        ("ours", "PRISM (本文)"),
    ]
    fig, axes = plt.subplots(2, 4, figsize=(10.5, 3.85), dpi=300)
    plt.subplots_adjust(left=0.025, right=0.998, top=0.91, bottom=0.05, wspace=0.012, hspace=0.14)

    for r, (stem, row_label) in enumerate(scenes):
        for c, (m_key, m_title) in enumerate(methods):
            ax = axes[r, c]
            img_path = get_course_img_path(m_key, stem)
            with Image.open(img_path) as im:
                cropped = fit_crop(im.convert("RGB"), 720, 440)
            ax.imshow(cropped)
            ax.set_xticks([])
            ax.set_yticks([])
            for spine in ax.spines.values():
                spine.set_visible(False)
            if r == 0:
                ax.set_title(m_title, fontsize=10.0, fontweight="bold", color=INK_BLACK, pad=5)
            if c == 0:
                ax.set_ylabel(row_label, fontsize=9.5, fontweight="bold", family="Times New Roman", color=INK_BLACK, labelpad=4)
            ax.set_xlabel(
                format_metric_label(stem, m_key),
                fontsize=8.5,
                color=INK_BLACK,
                labelpad=2,
            )

    out_path = OUT_DIR / "cover_teaser.png"
    fig.savefig(out_path, dpi=300, bbox_inches="tight", pad_inches=0.02, facecolor="white")
    plt.close(fig)
    print("Saved:", out_path)


def make_fig1_architecture():
    """Render standalone figure1.pdf at 360 DPI."""
    pdf_path = ROOT / "figure1.pdf"
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=360, alpha=False)
    out_path = OUT_DIR / "fig1_architecture.png"
    pix.save(str(out_path))
    doc.close()
    print("Saved:", out_path)


def make_fig2_landscape():
    """Render Fig. 2 (56-combination scatter landscape across 3 seeds) at 300 DPI."""
    fig, axes = plt.subplots(1, 3, figsize=(11.2, 3.35), sharey=True, dpi=300)
    seeds = [3407, 3408, 3409]
    for ax, seed in zip(axes, seeds):
        probe_dir = PROJECT / f"quality_selection/probes/eva_centered_{seed}/full"
        decision = json.loads((probe_dir / "frozen_selection.json").read_text(encoding="utf-8"))
        rows = [json.loads(line) for line in (probe_dir / "subsets32.jsonl").read_text(encoding="utf-8").splitlines() if line.strip()]
        ref = decision["published"]
        chosen = decision["quality_constrained"]
        h2c_best = decision["h2c_best"]
        pareto_keys = {tuple(item) for item in decision["pareto_frontier"]}

        xs = [r["h2c32_loss"] for r in rows]
        ys = [r["musiq"] for r in rows]
        tau = ref["musiq"]

        x_min, x_max = min(xs) - 0.0005, max(xs) + 0.0006
        y_min, y_max = min(ys) - 0.35, max(ys) + 0.45
        ax.set_xlim(x_min, x_max)
        ax.set_ylim(y_min, y_max)

        ax.axhspan(y_min, tau, color="#FEE2E2", alpha=0.45, zorder=0)
        ax.axhspan(tau, y_max, color="#CCFBF1", alpha=0.42, zorder=0)
        ax.axhline(tau, color="#334155", linestyle="--", linewidth=1.0, zorder=2)

        infeasible = [r for r in rows if r["musiq"] < tau]
        feasible = [r for r in rows if r["musiq"] >= tau]
        pareto = [r for r in rows if tuple(r["selected"]) in pareto_keys]

        ax.scatter(
            [r["h2c32_loss"] for r in infeasible],
            [r["musiq"] for r in infeasible],
            s=20,
            c="#94A3B8",
            alpha=0.75,
            label="不可行子集 (< τ)",
            zorder=3,
        )
        ax.scatter(
            [r["h2c32_loss"] for r in feasible],
            [r["musiq"] for r in feasible],
            s=24,
            c="#0D9488",
            alpha=0.80,
            label="感知可行子集 (≥ τ)",
            zorder=4,
        )
        pareto_sorted = sorted(pareto, key=lambda r: r["h2c32_loss"])
        ax.plot(
            [r["h2c32_loss"] for r in pareto_sorted],
            [r["musiq"] for r in pareto_sorted],
            color="#0F766E",
            linewidth=1.2,
            linestyle="-.",
            label="帕累托前沿",
            zorder=4,
        )

        ax.scatter([ref["h2c32_loss"]], [ref["musiq"]], s=75, marker="s", c="#1E40AF", edgecolors="white", linewidths=0.8, label="参考组合 {1,7,8}", zorder=6)
        ax.scatter([h2c_best["h2c32_loss"]], [h2c_best["musiq"]], s=85, marker="^", c="#B91C1C", edgecolors="white", linewidths=0.8, label="H2C 最优 {5,6,8}", zorder=6)
        ax.scatter([chosen["h2c32_loss"]], [chosen["musiq"]], s=115, marker="*", c="#047857", edgecolors="#022C22", linewidths=0.7, label="PRISM 选中 {6,7,8}", zorder=7)

        ax.annotate("{1,7,8}", (ref["h2c32_loss"], ref["musiq"]), xytext=(4, -11), textcoords="offset points", fontsize=8.0, color="#1E293B", fontweight="bold")
        ax.annotate("{5,6,8}", (h2c_best["h2c32_loss"], h2c_best["musiq"]), xytext=(4, 4), textcoords="offset points", fontsize=8.0, color="#991B1B", fontweight="bold")
        ax.annotate("{6,7,8}", (chosen["h2c32_loss"], chosen["musiq"]), xytext=(5, -3), textcoords="offset points", fontsize=8.5, color="#065F46", fontweight="bold")

        ax.text(x_max - 0.00012, y_max - 0.08, "Perceptually Feasible Region\n(感知安全可行域)", fontsize=7.2, color="#0F766E", ha="right", va="top", fontweight="bold")
        ax.text(x_max - 0.00012, y_min + 0.08, "Perceptual Degradation Zone\n(语义过拟合退化区)", fontsize=7.2, color="#991B1B", ha="right", va="bottom", fontweight="bold")
        ax.set_title(f"随机种子 {seed}（可行组合：{len(feasible)}/56）", fontsize=9.5, fontweight="bold", color=INK_BLACK)
        ax.set_xlabel("选层集 32 幅雾图平均 H2C 损失 h(S) ↓", fontsize=8.5, color=INK_BLACK)
        ax.grid(True, linestyle=":", alpha=0.4)

    axes[0].set_ylabel("选层集 32 幅整图平均 MUSIQ p(S) ↑", fontsize=8.8, color=INK_BLACK)
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="upper center", ncol=6, fontsize=8.0, frameon=True, bbox_to_anchor=(0.5, 1.03))
    plt.tight_layout(rect=[0, 0, 1, 0.92])
    out_path = OUT_DIR / "fig2_landscape.png"
    fig.savefig(out_path, dpi=300, bbox_inches="tight", pad_inches=0.03, facecolor="white")
    plt.close(fig)
    print("Saved:", out_path)


def make_course10_matrix(stems, out_name):
    """Generate large, ultra-clear full-page 5x4 minimalist visual matrix for 5 Course-10 images."""
    methods = [
        ("input", "有雾原图"),
        ("base", "DEA-Net Base"),
        ("bilalora", "BiLaLoRA"),
        ("ours", "PRISM (本文)"),
    ]
    # Give tall portrait 'bench' extra row height so its sky and canopy are not over-cropped
    has_bench = stems[0] == "bench"
    ratios = [1.45, 1.0, 1.0, 1.0, 1.0] if has_bench else [1.0, 1.0, 1.0, 1.0, 1.0]
    fig, axes = plt.subplots(
        len(stems),
        4,
        figsize=(10.5, 13.4),
        dpi=300,
        gridspec_kw={"height_ratios": ratios},
    )
    plt.subplots_adjust(left=0.025, right=0.998, top=0.972, bottom=0.022, wspace=0.015, hspace=0.145)

    for r, stem in enumerate(stems):
        target_wh = (640, 680) if stem == "bench" else (640, 480)
        for c, (m_key, m_title) in enumerate(methods):
            ax = axes[r, c]
            img_path = get_course_img_path(m_key, stem)
            with Image.open(img_path) as im:
                cropped = fit_crop(im.convert("RGB"), target_wh[0], target_wh[1])
            ax.imshow(cropped)
            ax.set_xticks([])
            ax.set_yticks([])
            for spine in ax.spines.values():
                spine.set_visible(False)
            if r == 0:
                ax.set_title(m_title, fontsize=11.0, fontweight="bold", color=INK_BLACK, pad=6)
            if c == 0:
                ax.set_ylabel(stem, fontsize=10.5, fontweight="bold", family="Times New Roman", color=INK_BLACK, labelpad=4)
            ax.set_xlabel(
                format_metric_label(stem, m_key),
                fontsize=9.0,
                color=INK_BLACK,
                labelpad=2.5,
            )

    out_path = OUT_DIR / out_name
    fig.savefig(out_path, dpi=300, bbox_inches="tight", pad_inches=0.02, facecolor="white")
    plt.close(fig)
    print("Saved:", out_path)


def make_course10_zoom():
    """Generate minimalist 3x5 local zoom & 8x difference figure for neighbour, tiananmen, pumpkins."""
    configs = [
        ("neighbour", (0.28, 0.18, 0.68, 0.68)),
        ("tiananmen", (0.22, 0.25, 0.72, 0.75)),
        ("pumpkins", (0.18, 0.12, 0.68, 0.62)),
    ]
    col_titles = [
        "有雾原图 (含放大框)",
        "DEA-Net Base (局部)",
        "BiLaLoRA (局部)",
        "PRISM (本文, 局部)",
        "8× |PRISM - BiLaLoRA|",
    ]
    fig, axes = plt.subplots(3, 5, figsize=(11.2, 4.95), dpi=300)
    plt.subplots_adjust(left=0.022, right=0.998, top=0.925, bottom=0.04, wspace=0.012, hspace=0.14)

    for r, (stem, (x0, y0, x1, y1)) in enumerate(configs):
        imgs = {}
        for k in ["input", "base", "bilalora", "ours"]:
            with Image.open(get_course_img_path(k, stem)) as im:
                imgs[k] = im.convert("RGB")
        w, h = imgs["input"].size
        box = (int(w * x0), int(h * y0), int(w * x1), int(h * y1))

        full_vis = imgs["input"].copy()
        draw = ImageDraw.Draw(full_vis)
        bw = max(3, w // 140)
        draw.rectangle(box, outline=(220, 38, 38), width=bw)
        full_vis = fit_crop(full_vis, 480, 330)
        axes[r, 0].imshow(full_vis)
        axes[r, 0].set_ylabel(stem, fontsize=9.8, fontweight="bold", family="Times New Roman", color=INK_BLACK, labelpad=4)
        axes[r, 0].set_xlabel(format_metric_label(stem, "input"), fontsize=8.2, color=INK_BLACK, labelpad=2)

        for c, k in enumerate(["base", "bilalora", "ours"], start=1):
            crop = imgs[k].crop(box).resize((480, 330), Image.Resampling.LANCZOS)
            axes[r, c].imshow(crop)
            axes[r, c].set_xlabel(format_metric_label(stem, k), fontsize=8.2, color=INK_BLACK, labelpad=2)

        arr_ours = np.asarray(imgs["ours"].crop(box).resize((480, 330), Image.Resampling.LANCZOS), dtype=np.float32)
        arr_bila = np.asarray(imgs["bilalora"].crop(box).resize((480, 330), Image.Resampling.LANCZOS), dtype=np.float32)
        diff = np.clip(np.abs(arr_ours - arr_bila) * 8.0, 0, 255).astype(np.uint8)
        axes[r, 4].imshow(diff)
        d_fade = COURSE_METRICS[stem]["ours"]["FADE"] - COURSE_METRICS[stem]["bilalora"]["FADE"]
        d_musiq = COURSE_METRICS[stem]["ours"]["MUSIQ"] - COURSE_METRICS[stem]["bilalora"]["MUSIQ"]
        axes[r, 4].set_xlabel(
            f"$\\Delta\\mathrm{{FADE}}={d_fade:+.4f},\\ \\Delta\\mathrm{{MUSIQ}}={d_musiq:+.2f}$",
            fontsize=8.0,
            color=INK_BLACK,
            labelpad=2,
        )

        for c in range(5):
            ax = axes[r, c]
            ax.set_xticks([])
            ax.set_yticks([])
            for spine in ax.spines.values():
                spine.set_visible(False)
            if r == 0:
                ax.set_title(col_titles[c], fontsize=9.6, fontweight="bold", color=INK_BLACK, pad=5)

    out_path = OUT_DIR / "fig5_course10_zoom.png"
    fig.savefig(out_path, dpi=300, bbox_inches="tight", pad_inches=0.02, facecolor="white")
    plt.close(fig)
    print("Saved:", out_path)


def make_fattal_qualitative():
    """Generate minimalist 3x5 Fattal qualitative comparison (swan, snow, buildings) at 300 DPI."""
    rows = [
        ("swan", ["$0.3081\\ /\\ 71.14$", "$0.2503\\ /\\ 70.52$", "$\\mathbf{0.2423}\\ /\\ \\mathbf{71.97}$", "$\\Delta\\mathrm{MUSIQ}=+1.45$"]),
        ("snow", ["$0.7007\\ /\\ 65.42$", "$0.6558\\ /\\ 67.45$", "$\\mathbf{0.6231}\\ /\\ \\mathbf{68.23}$", "$\\Delta\\mathrm{MUSIQ}=+0.79$"]),
        ("buildings", ["$0.4828\\ /\\ 69.91$", "$0.4109\\ /\\ 70.21$", "$\\mathbf{0.3934}\\ /\\ \\mathbf{70.84}$", "$\\Delta\\mathrm{MUSIQ}=+0.62$"]),
    ]
    col_titles = [
        "有雾原图 (2:1 局部)",
        "BiLaLoRA 官方权重",
        "参考位置 {1,7,8}+EVA",
        "PRISM (本文, {6,7,8})",
        "8× |PRISM - 参考位置|",
    ]
    suffixes = [
        "input_crop.png",
        "official_bilalora_crop.png",
        "warm_published_3407_crop.png",
        "warm_quality_constrained_3407_crop.png",
        "difference.png",
    ]
    fig, axes = plt.subplots(3, 5, figsize=(11.2, 4.85), dpi=300)
    plt.subplots_adjust(left=0.022, right=0.998, top=0.925, bottom=0.04, wspace=0.012, hspace=0.14)

    for r, (stem, labels) in enumerate(rows):
        all_labels = [f"{stem}"] + labels
        for c, suf in enumerate(suffixes):
            ax = axes[r, c]
            with Image.open(FIG_DIR / f"{stem}_{suf}") as im:
                ax.imshow(im.convert("RGB"))
            ax.set_xticks([])
            ax.set_yticks([])
            for spine in ax.spines.values():
                spine.set_visible(False)
            if r == 0:
                ax.set_title(col_titles[c], fontsize=9.6, fontweight="bold", color=INK_BLACK, pad=5)
            if c == 0:
                ax.set_ylabel(stem, fontsize=9.8, fontweight="bold", family="Times New Roman", color=INK_BLACK, labelpad=4)
            ax.set_xlabel(all_labels[c], fontsize=8.2, color=INK_BLACK, labelpad=2)

    out_path = OUT_DIR / "fig6_fattal_qualitative.png"
    fig.savefig(out_path, dpi=300, bbox_inches="tight", pad_inches=0.02, facecolor="white")
    plt.close(fig)
    print("Saved:", out_path)


def make_train_boundary():
    """Generate minimalist 1x5 Fattal train boundary case figure at 300 DPI."""
    col_titles = [
        "有雾原图 (train)",
        "BiLaLoRA 官方权重",
        "参考位置 {1,7,8}+EVA",
        "PRISM (本文, {6,7,8})",
        "8× |PRISM - 参考位置|",
    ]
    suffixes = [
        "input_crop.png",
        "official_bilalora_crop.png",
        "warm_published_3407_crop.png",
        "warm_quality_constrained_3407_crop.png",
        "difference.png",
    ]
    labels = [
        "train (局部)",
        "$0.1953\\ /\\ \\mathbf{63.67}$",
        "$0.1814\\ /\\ 58.88$",
        "$\\mathbf{0.1722}\\ /\\ 57.92$",
        "$\\Delta\\mathrm{FADE}=-0.0092,\\ \\Delta\\mathrm{MUSIQ}=-0.95$",
    ]
    fig, axes = plt.subplots(1, 5, figsize=(11.2, 1.95), dpi=300)
    plt.subplots_adjust(left=0.015, right=0.998, top=0.86, bottom=0.14, wspace=0.012)

    for c, suf in enumerate(suffixes):
        ax = axes[c]
        with Image.open(FIG_DIR / f"train_{suf}") as im:
            ax.imshow(im.convert("RGB"))
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.set_title(col_titles[c], fontsize=9.5, fontweight="bold", color=INK_BLACK, pad=4)
        ax.set_xlabel(labels[c], fontsize=8.2, color=INK_BLACK, labelpad=2)

    out_path = OUT_DIR / "fig7_train_boundary.png"
    fig.savefig(out_path, dpi=300, bbox_inches="tight", pad_inches=0.02, facecolor="white")
    plt.close(fig)
    print("Saved:", out_path)


if __name__ == "__main__":
    make_cover_teaser()
    make_fig1_architecture()
    make_fig2_landscape()
    make_course10_matrix(["bench", "cones", "dubai", "forest", "herzeliya"], "fig3_course10_part1.png")
    make_course10_matrix(["neighbour", "pumpkins", "road", "tiananmen", "train"], "fig4_course10_part2.png")
    make_course10_zoom()
    make_fattal_qualitative()
    make_train_boundary()
    print("All 8 minimalist CVPR/TIP-style 300-DPI figures generated successfully.")
