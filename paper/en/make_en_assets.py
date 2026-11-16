# -*- coding: utf-8 -*-
"""Generate pure-English figures (zero Chinese characters) for project/paper_en/."""
from pathlib import Path
import json
import sys
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFile

ImageFile.MAXBLOCK = 4 * 1024 * 1024

plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial", "Helvetica"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["mathtext.fontset"] = "stix"

EN_ROOT = Path(__file__).resolve().parent
CN_ROOT = EN_ROOT.parent / "paper_cn_v3"
PROJECT = EN_ROOT.parent
PROBE_ROOT = PROJECT / "quality_selection/probes"
COURSE10 = PROJECT / "course_10"
FIG_DIR = EN_ROOT / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(CN_ROOT))
import generate_word_assets as gwa  # reuse verified metric loader & lossless crops

plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial", "Helvetica"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["mathtext.fontset"] = "stix"


def save_q92_rgb(fig, out_jpg: Path, max_w: int = 2300):
    tmp_png = out_jpg.with_suffix(".tmp.png")
    fig.savefig(tmp_png, dpi=300, bbox_inches="tight", pad_inches=0.02, facecolor="#FFFFFF")
    plt.close(fig)
    im = Image.open(tmp_png).convert("RGB")
    if im.width > max_w:
        nh = max(1, round(im.height * (max_w / im.width)))
        im = im.resize((max_w, nh), Image.Resampling.LANCZOS)
    im.save(out_jpg, format="JPEG", quality=92, subsampling=0, optimize=True, dpi=(300, 300))
    tmp_png.unlink(missing_ok=True)
    print("Saved English JPEG:", out_jpg)


def make_landscape_en():
    """True-vector PDF 56-subset feasibility landscape across 3 seeds (100% English)."""
    fig, axes = plt.subplots(1, 3, figsize=(11.2, 3.45), sharey=True, dpi=300)
    seeds = [3407, 3408, 3409]
    for ax, seed in zip(axes, seeds):
        probe_dir = PROBE_ROOT / f"eva_centered_{seed}/full"
        decision = json.loads((probe_dir / "frozen_selection.json").read_text(encoding="utf-8"))
        rows = [
            json.loads(line)
            for line in (probe_dir / "subsets32.jsonl").read_text(encoding="utf-8").splitlines()
            if line.strip()
        ]
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

        ax.axhspan(y_min, tau, color="#FEE2E2", alpha=0.48, zorder=0)
        ax.axhspan(tau, y_max, color="#CCFBF1", alpha=0.45, zorder=0)
        ax.axhline(tau, color="#334155", linestyle="--", linewidth=1.05, zorder=2)

        infeasible = [r for r in rows if r["musiq"] < tau]
        feasible = [r for r in rows if r["musiq"] >= tau]
        pareto = [r for r in rows if tuple(r["selected"]) in pareto_keys]

        ax.scatter(
            [r["h2c32_loss"] for r in infeasible],
            [r["musiq"] for r in infeasible],
            s=20, c="#94A3B8", alpha=0.75, label=r"Infeasible ($<\tau$)", zorder=3,
        )
        ax.scatter(
            [r["h2c32_loss"] for r in feasible],
            [r["musiq"] for r in feasible],
            s=24, c="#0D9488", alpha=0.82, label=r"Feasible ($\geq\tau$)", zorder=4,
        )
        pareto_sorted = sorted(pareto, key=lambda r: r["h2c32_loss"])
        ax.plot(
            [r["h2c32_loss"] for r in pareto_sorted],
            [r["musiq"] for r in pareto_sorted],
            color="#0F766E", linewidth=1.2, linestyle="-.", label="Pareto Frontier", zorder=4,
        )

        ax.scatter([ref["h2c32_loss"]], [ref["musiq"]], s=75, marker="s", c="#1E40AF", edgecolors="white", linewidths=0.8, label="Ref {1,7,8}", zorder=6)
        ax.scatter([h2c_best["h2c32_loss"]], [h2c_best["musiq"]], s=85, marker="^", c="#B91C1C", edgecolors="white", linewidths=0.8, label="H2C-min {5,6,8}", zorder=6)
        ax.scatter([chosen["h2c32_loss"]], [chosen["musiq"]], s=115, marker="*", c="#047857", edgecolors="#022C22", linewidths=0.7, label="PRISM {6,7,8}", zorder=7)

        ax.annotate("{1,7,8}", (ref["h2c32_loss"], ref["musiq"]), xytext=(4, -11), textcoords="offset points", fontsize=8.0, color="#1E293B", fontweight="bold")
        ax.annotate("{5,6,8}", (h2c_best["h2c32_loss"], h2c_best["musiq"]), xytext=(4, 4), textcoords="offset points", fontsize=8.0, color="#991B1B", fontweight="bold")
        ax.annotate("{6,7,8}", (chosen["h2c32_loss"], chosen["musiq"]), xytext=(5, -3), textcoords="offset points", fontsize=8.5, color="#065F46", fontweight="bold")

        ax.text(x_max - 0.00012, y_max - 0.07, "Perceptually Feasible Region", fontsize=7.5, color="#0F766E", ha="right", va="top", fontweight="bold")
        ax.text(x_max - 0.00012, min(ys) - 0.25, "Perceptual Degradation Zone", fontsize=7.5, color="#991B1B", ha="right", va="bottom", fontweight="bold")
        ax.set_title(f"Seed {seed} (Feasible: {len(feasible)}/56)", fontsize=9.8, fontweight="bold", color="#0F172A")
        ax.set_xlabel(r"Selection-Set Mean H2C Loss $h(\mathcal{S})$ $\downarrow$", fontsize=8.6, color="#0F172A")
        ax.grid(True, linestyle=":", alpha=0.4)

    axes[0].set_ylabel(r"Whole-Image MUSIQ $p(\mathcal{S})$ $\uparrow$", fontsize=8.8, color="#0F172A")
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="upper center", ncol=6, fontsize=8.1, frameon=True, bbox_to_anchor=(0.5, 1.03))
    plt.tight_layout(rect=[0, 0, 1, 0.92])
    out_pdf = FIG_DIR / "fig_landscape_en.pdf"
    fig.savefig(out_pdf, bbox_inches="tight", pad_inches=0.03, facecolor="#FFFFFF")
    plt.close(fig)
    print("Saved English Vector PDF:", out_pdf)


def make_classic10_matrix_en(stems, out_name):
    methods = [
        ("input", "Hazy Input"),
        ("base", "DEA-Net Base"),
        ("bilalora", "BiLaLoRA"),
        ("ours", "PRISM (Ours)"),
    ]
    has_bench = stems[0] == "bench"
    ratios = [1.45, 1.0, 1.0, 1.0, 1.0] if has_bench else [1.0, 1.0, 1.0, 1.0, 1.0]
    fig, axes = plt.subplots(len(stems), 4, figsize=(10.5, 13.4), dpi=300, gridspec_kw={"height_ratios": ratios})
    plt.subplots_adjust(left=0.025, right=0.998, top=0.972, bottom=0.022, wspace=0.015, hspace=0.145)

    for r, stem in enumerate(stems):
        target_wh = (640, 680) if stem == "bench" else (640, 480)
        for c, (m_key, m_title) in enumerate(methods):
            ax = axes[r, c]
            img_path = gwa.get_course_img_path(m_key, stem)
            with Image.open(img_path) as im:
                cropped = gwa.fit_crop(im.convert("RGB"), target_wh[0], target_wh[1])
            ax.imshow(cropped)
            ax.set_xticks([])
            ax.set_yticks([])
            for sp in ax.spines.values():
                sp.set_visible(False)
            if r == 0:
                ax.set_title(m_title, fontsize=11.0, fontweight="bold", color="#0F172A", pad=6)
            if c == 0:
                ax.set_ylabel(stem, fontsize=10.5, fontweight="bold", family="Times New Roman", color="#0F172A", labelpad=4)
            ax.set_xlabel(gwa.format_metric_label(stem, m_key), fontsize=9.0, color="#0F172A", labelpad=2.5)

    save_q92_rgb(fig, FIG_DIR / out_name, max_w=2150)


def make_classic10_zoom_en():
    configs = [
        ("neighbour", (0.28, 0.18, 0.68, 0.68)),
        ("tiananmen", (0.22, 0.25, 0.72, 0.75)),
        ("pumpkins", (0.18, 0.12, 0.68, 0.62)),
    ]
    col_titles = [
        "Hazy Input (Zoom ROI)",
        "DEA-Net Base (Crop)",
        "BiLaLoRA (Crop)",
        "PRISM (Ours, Crop)",
        "8× |PRISM − BiLaLoRA|",
    ]
    fig, axes = plt.subplots(3, 5, figsize=(11.2, 4.95), dpi=300)
    plt.subplots_adjust(left=0.022, right=0.998, top=0.925, bottom=0.04, wspace=0.012, hspace=0.14)

    for r, (stem, (x0, y0, x1, y1)) in enumerate(configs):
        imgs = {}
        for k in ["input", "base", "bilalora", "ours"]:
            with Image.open(gwa.get_course_img_path(k, stem)) as im:
                imgs[k] = im.convert("RGB")
        w, h = imgs["input"].size
        box = (int(w * x0), int(h * y0), int(w * x1), int(h * y1))

        full_vis = imgs["input"].copy()
        draw = ImageDraw.Draw(full_vis)
        bw = max(3, w // 140)
        draw.rectangle(box, outline=(220, 38, 38), width=bw)
        full_vis = gwa.fit_crop(full_vis, 480, 330)
        axes[r, 0].imshow(full_vis)
        axes[r, 0].set_ylabel(stem, fontsize=9.8, fontweight="bold", family="Times New Roman", color="#0F172A", labelpad=4)
        axes[r, 0].set_xlabel(gwa.format_metric_label(stem, "input"), fontsize=8.2, color="#0F172A", labelpad=2)

        for c, k in enumerate(["base", "bilalora", "ours"], start=1):
            crop = imgs[k].crop(box).resize((480, 330), Image.Resampling.LANCZOS)
            axes[r, c].imshow(crop)
            axes[r, c].set_xlabel(gwa.format_metric_label(stem, k), fontsize=8.2, color="#0F172A", labelpad=2)

        arr_ours = np.asarray(imgs["ours"].crop(box).resize((480, 330), Image.Resampling.LANCZOS), dtype=np.float32)
        arr_bila = np.asarray(imgs["bilalora"].crop(box).resize((480, 330), Image.Resampling.LANCZOS), dtype=np.float32)
        diff = np.clip(np.abs(arr_ours - arr_bila) * 8.0, 0, 255).astype(np.uint8)
        axes[r, 4].imshow(diff)
        d_fade = gwa.COURSE_METRICS[stem]["ours"]["FADE"] - gwa.COURSE_METRICS[stem]["bilalora"]["FADE"]
        d_musiq = gwa.COURSE_METRICS[stem]["ours"]["MUSIQ"] - gwa.COURSE_METRICS[stem]["bilalora"]["MUSIQ"]
        axes[r, 4].set_xlabel(f"$\\Delta\\mathrm{{FADE}}={d_fade:+.4f},\\ \\Delta\\mathrm{{MUSIQ}}={d_musiq:+.2f}$", fontsize=8.0, color="#0F172A", labelpad=2)

        for c in range(5):
            ax = axes[r, c]
            ax.set_xticks([])
            ax.set_yticks([])
            for sp in ax.spines.values():
                sp.set_visible(False)
            if r == 0:
                ax.set_title(col_titles[c], fontsize=9.6, fontweight="bold", color="#0F172A", pad=5)

    save_q92_rgb(fig, FIG_DIR / "fig_classic10_zoom_en.jpg", max_w=2250)


if __name__ == "__main__":
    make_landscape_en()
    make_classic10_matrix_en(["bench", "cones", "dubai", "forest", "herzeliya"], "fig_classic10_part1_en.jpg")
    make_classic10_matrix_en(["neighbour", "pumpkins", "road", "tiananmen", "train"], "fig_classic10_part2_en.jpg")
    make_classic10_zoom_en()
    print("All pure-English assets generated successfully.")
