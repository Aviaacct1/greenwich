"""
Avia chart style (WS11), matplotlib mirror of the canonical engine format module
C:\\Avia\\avia_forecast_build\\avia_forecast\\outputs\\chart_format.py, which
implements "Avia Chart Format and Chart Catalogue": Arial/Ebrima, heading 20pt
bold centred, singular Source line, 18pt axes, legend at the bottom, data labels
bold in the line colour, no borders or gridlines, pinned Office 2024 palette,
A/F year suffixes. Constants mirrored, not reinvented; if chart_format.py
changes, change this to match. Author: Avia Solutions.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# Mirrored from chart_format.py (pinned Office 2024 accents)
PINNED_PALETTE = ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"]
SIZES = {"heading": 20, "source": 16, "axis": 18, "axis_title": 18, "legend": 18,
         "data_label": 18, "marker": 10}
FONTS = ["Arial", "Ebrima", "Liberation Sans", "DejaVu Sans"]

def source_line(sources="OAG, AviaSolutions analysis"):
    return f"Source: {sources}"          # singular by house rule

def year_label(year, base_year):
    return f"{year}{'A' if year <= base_year else 'F'}"

def avia_style():
    plt.rcParams.update({
        "font.family": "sans-serif", "font.sans-serif": FONTS,
        "axes.grid": False,                       # no gridlines
        "axes.spines.top": False, "axes.spines.right": False,
        "axes.spines.left": False,                # no borders beyond the x axis line
        "axes.prop_cycle": plt.cycler(color=PINNED_PALETTE),
        "figure.facecolor": "white", "axes.facecolor": "white",
        "legend.frameon": False,
    })

def avia_line_chart(series, title, sources, out_path, base_year=None,
                    end_labels=True, size=(12.8, 7.2)):
    """series: dict name -> (years, values). base_year drives A/F axis labels.
    Heading 20pt bold centred; legend bottom; end-of-line data labels bold in the
    line colour; Source line mandatory and singular."""
    sl = source_line(sources)
    if not sl.startswith("Source:"):
        raise ValueError("house rule: singular 'Source:' line required")
    avia_style()
    fig, ax = plt.subplots(figsize=size, dpi=150)
    for k, (name, (yrs, vals)) in enumerate(series.items()):
        colour = PINNED_PALETTE[k % len(PINNED_PALETTE)]
        ax.plot(yrs, vals, linewidth=2.5, label=name, color=colour)
        if end_labels:
            ax.annotate(f"{vals[-1]:,.1f}", xy=(yrs[-1], vals[-1]),
                        xytext=(6, 0), textcoords="offset points",
                        fontsize=SIZES["data_label"], fontweight="bold", color=colour)
    ax.set_title(title, fontsize=SIZES["heading"], fontweight="bold", loc="center", pad=14)
    if base_year is not None:
        step = max(1, len(next(iter(series.values()))[0]) // 12)
        yrs0 = next(iter(series.values()))[0]
        ticks = [y for i, y in enumerate(yrs0) if i % step == 0]
        ax.set_xticks(ticks)
        ax.set_xticklabels([year_label(y, base_year) for y in ticks],
                           fontsize=SIZES["axis"])
    ax.tick_params(axis="y", labelsize=SIZES["axis"])
    ax.legend(fontsize=SIZES["legend"], loc="upper center",
              bbox_to_anchor=(0.5, -0.08), ncol=min(len(series), 3))
    fig.text(0.01, 0.005, sl, fontsize=SIZES["source"], color="#404040")
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    fig.savefig(out_path)
    plt.close(fig)
    return out_path

if __name__ == "__main__":
    yrs = list(range(2019, 2036))
    vals = [5.1, 2.2, 3.4, 4.8, 5.3, 5.5]
    while len(vals) < len(yrs): vals.append(round(vals[-1] * 1.04, 3))
    # Output resolves from the data root, or is given on the command line. It previously
    # carried an absolute path from a working session, which wrote nowhere on any other host.
    import os, sys, config
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        config.output_dir(), "style_test.png")
    avia_line_chart({"Passengers (m, two-way)": (yrs, vals)},
        "Passenger Traffic 2019-2035",
        "Illustrative data; AviaSolutions analysis",
        out,
        base_year=2024)
    print("rendered", out)
