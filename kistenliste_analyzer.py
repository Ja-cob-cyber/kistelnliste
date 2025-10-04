"""
Kistenliste Analyzer für Fußballmannschaft
Erstellt automatisch Diagramme und Auswertungen
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import webbrowser
import warnings
import base64
from io import BytesIO
from datetime import datetime

# Konfiguration
EXCEL_FILE = "Kistenliste.xlsx"
AUTO_OPEN_BROWSER = True


def load_data(filepath):
    """Lädt die Excel-Datei"""
    df = pd.read_excel(filepath, sheet_name="Kistenliste")
    df["Name"] = df["Name"].str.strip()

    # Bezahlt-Status korrigieren: nur "J" ist bezahlt, alles andere ist offen
    df["Bezahlt"] = df["Bezahlt"].fillna("").str.strip()
    df["Bezahlt_Status"] = df["Bezahlt"].apply(
        lambda x: "Bezahlt" if x == "J" else "Offen"
    )

    return df


def create_visualizations(df):
    """Erstellt alle Diagramme mit kräftigen Farben"""
    warnings.filterwarnings(
        "ignore", category=UserWarning, message=".*Glyph.*missing from font.*"
    )

    # Style einstellen
    sns.set_style("whitegrid")
    plt.rcParams["figure.facecolor"] = "white"

    # Figure mit Subplots erstellen
    fig = plt.figure(figsize=(16, 12))

    # 1. Kisten pro Person (Gestapeltes Balkendiagramm) - Bezahlt vs. Offen
    ax1 = plt.subplot(2, 2, 1)

    # Daten für gestapeltes Diagramm vorbereiten
    name_bezahlt = df[df["Bezahlt_Status"] == "Bezahlt"].groupby("Name").size()
    name_offen = df[df["Bezahlt_Status"] == "Offen"].groupby("Name").size()

    # Alle Namen mit Kisten
    all_names = df["Name"].value_counts().index
    name_stats = pd.DataFrame(
        {
            "Name": all_names,
            "Bezahlt": [name_bezahlt.get(name, 0) for name in all_names],
            "Offen": [name_offen.get(name, 0) for name in all_names],
        }
    )

    # Nach Gesamt sortieren
    name_stats["Gesamt"] = name_stats["Bezahlt"] + name_stats["Offen"]
    name_stats = name_stats.sort_values("Gesamt", ascending=True)

    # Gestapeltes Balkendiagramm
    y_pos = range(len(name_stats))
    ax1.barh(
        y_pos,
        name_stats["Bezahlt"],
        color="#16a34a",
        label="Bezahlt",
        edgecolor="darkgreen",
        linewidth=1.5,
    )
    ax1.barh(
        y_pos,
        name_stats["Offen"],
        left=name_stats["Bezahlt"],
        color="#dc2626",
        label="Offen",
        edgecolor="darkred",
        linewidth=1.5,
    )

    ax1.set_yticks(y_pos)
    ax1.set_yticklabels(name_stats["Name"])
    ax1.set_title("🏆 Kisten pro Person", fontsize=14, fontweight="bold", pad=20)
    ax1.set_xlabel("Anzahl Kisten", fontsize=11, fontweight="bold")
    ax1.set_ylabel("Name", fontsize=11, fontweight="bold")
    ax1.legend(loc="lower right")
    ax1.grid(axis="x", alpha=0.3)

    # Ganzzahlige X-Achse
    ax1.xaxis.set_major_locator(plt.MaxNLocator(integer=True))

    # Werte anzeigen
    for i, row in enumerate(name_stats.itertuples()):
        total = row.Gesamt
        ax1.text(
            total + 0.1, i, str(total), va="center", fontweight="bold", fontsize=10
        )

    # 2. Bezahlstatus (Tortendiagramm) - KRÄFTIGE FARBEN
    ax2 = plt.subplot(2, 2, 2)
    bezahlt_counts = df["Bezahlt_Status"].value_counts()

    # Kräftige Farben: Grün für bezahlt, Rot für offen
    colors_pie = ["#dc2626", "#16a34a"]  # Kräftiges Grün und Rot
    explode = (0.05, 0.05)  # Leicht auseinander

    wedges, texts, autotexts = ax2.pie(
        bezahlt_counts.values,
        labels=bezahlt_counts.index,
        autopct="%1.1f%%",
        colors=colors_pie,
        startangle=90,
        explode=explode,
        textprops={"fontsize": 11, "fontweight": "bold"},
        wedgeprops={"edgecolor": "white", "linewidth": 2},
    )
    ax2.set_title("💰 Bezahlstatus", fontsize=14, fontweight="bold", pad=20)

    # 3. Häufigste Gründe (Top 10) - KRÄFTIGE GRÜNE TÖNE
    ax3 = plt.subplot(2, 1, 2)
    grund_counts = df["Grund"].value_counts().head(10).sort_values(ascending=True)

    # Kräftiger Grün-Gradient
    n_bars_grund = len(grund_counts)
    colors_green = [
        plt.cm.Greens(0.5 + 0.5 * i / max(n_bars_grund - 1, 1))
        for i in range(n_bars_grund)
    ]

    grund_counts.plot(
        kind="barh", ax=ax3, color=colors_green, edgecolor="darkgreen", linewidth=1.5
    )
    ax3.set_title("📋 Top 10 Häufigste Gründe", fontsize=14, fontweight="bold", pad=20)
    ax3.set_xlabel("Anzahl", fontsize=11, fontweight="bold")
    ax3.set_ylabel("Grund", fontsize=11, fontweight="bold")
    ax3.grid(axis="x", alpha=0.3)

    # Ganzzahlige X-Achse
    ax3.xaxis.set_major_locator(plt.MaxNLocator(integer=True))

    # Werte an den Balken anzeigen
    for i, v in enumerate(grund_counts.values):
        ax3.text(v + 0.05, i, str(v), va="center", fontweight="bold", fontsize=10)

    plt.tight_layout(pad=3.0)

    return fig


def create_statistics_table(df):
    """Erstellt Statistik-Tabelle"""
    bezahlt_count = len(df[df["Bezahlt_Status"] == "Bezahlt"])
    offen_count = len(df[df["Bezahlt_Status"] == "Offen"])

    stats = {
        "Gesamt Einträge": len(df),
        "Bezahlt": bezahlt_count,
        "Offen": offen_count,
        "Verschiedene Personen": df["Name"].nunique(),
        "Verschiedene Gründe": df["Grund"].nunique(),
    }
    return stats


def create_ranking_table(df):
    """Erstellt Rangliste"""
    ranking = df["Name"].value_counts().reset_index()
    ranking.columns = ["Name", "Anzahl Kisten"]
    ranking["Rang"] = range(1, len(ranking) + 1)

    # Medaillen für Top 3
    medals = {1: "🥇", 2: "🥈", 3: "🥉"}
    ranking["Medaille"] = ranking["Rang"].map(lambda x: medals.get(x, ""))

    return ranking[["Rang", "Medaille", "Name", "Anzahl Kisten"]]


def save_dashboard(df, fig, stats, ranking):
    """Speichert Dashboard als HTML mit externem Template"""

    # Diagramm als Base64 enkodieren
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=150, bbox_inches="tight")
    buffer.seek(0)
    image_base64 = base64.b64encode(buffer.read()).decode()

    # Ranking-Tabelle als HTML erstellen
    ranking_html = ""
    for _, row in ranking.iterrows():
        medal = row["Medaille"] if row["Medaille"] else f"{row['Rang']}."
        ranking_html += f"""
                    <tr>
                        <td>{medal}</td>
                        <td><strong>{row['Name']}</strong></td>
                        <td style="text-align: right; color: #2563eb; font-weight: bold;">{row['Anzahl Kisten']}</td>
                    </tr>
        """

    # HTML-Template laden und ersetzen
    template_path = "index.html"

    if os.path.exists(template_path):
        with open(template_path, "r", encoding="utf-8") as f:
            html_content = f.read()
    else:
        # Falls Template nicht existiert, Fallback verwenden
        print("⚠️  Template nicht gefunden, erstelle es...")
        print("   Erstelle index.html Template...")

    # Platzhalter ersetzen
    html_content = html_content.replace(
        "{{TIMESTAMP}}", datetime.now().strftime("%d.%m.%Y %H:%M")
    )
    html_content = html_content.replace("{{GESAMT}}", str(stats["Gesamt Einträge"]))
    html_content = html_content.replace("{{BEZAHLT}}", str(stats["Bezahlt"]))
    html_content = html_content.replace("{{OFFEN}}", str(stats["Offen"]))
    html_content = html_content.replace(
        "{{PERSONEN}}", str(stats["Verschiedene Personen"])
    )
    html_content = html_content.replace("{{DIAGRAMM_BASE64}}", image_base64)
    html_content = html_content.replace("{{RANKING_ROWS}}", ranking_html)

    # Speichern
    output_file = "index.html"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html_content)

    html_path = os.path.abspath(output_file)

    print(f"✅ HTML-Dashboard erstellt: {output_file}")
    print(f"\n🔗 Öffne im Browser:")
    print(f"   file:///{html_path.replace(chr(92), '/')}")

    return html_path


def main():
    """Hauptfunktion"""

    warnings.filterwarnings(
        "ignore", category=UserWarning, message=".*Glyph.*missing from font.*"
    )

    print("🚀 Starte Kistenliste-Analyse...\n")

    # Daten laden
    print(f"📂 Lade {EXCEL_FILE}...")
    df = load_data(EXCEL_FILE)
    print(f"✓ {len(df)} Einträge geladen\n")

    # Statistiken erstellen
    stats = create_statistics_table(df)
    ranking = create_ranking_table(df)

    # Visualisierungen erstellen
    print("📊 Erstelle Diagramme...")
    fig = create_visualizations(df)

    # Dashboard speichern
    html_path = save_dashboard(df, fig, stats, ranking)

    print("\n✨ Fertig! Analyse abgeschlossen.")
    print(f"\n📈 Schnelle Statistik:")
    for key, value in stats.items():
        print(f"   {key}: {value}")

    # Browser öffnen
    if AUTO_OPEN_BROWSER:
        print(f"\n🌐 Öffne Dashboard im Browser...")
        webbrowser.open(f'file:///{html_path.replace(chr(92), "/")}')
        print("✓ Browser geöffnet!")


if __name__ == "__main__":
    main()
