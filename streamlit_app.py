"""
Kistenliste Dashboard - Streamlit App
Für Deployment auf Streamlit Cloud/Render
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Seitenkonfiguration
st.set_page_config(
    page_title="Fc Münster 05 ",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Custom CSS für besseres Design
st.markdown(
    """
    <style>
    .main {
        background: linear-gradient(135deg, #f0fdf4 0%, #dbeafe 100%);
    }
    .stMetric {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h1 {
        color: #1f2937;
        text-align: center;
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_data():
    """Lädt die Excel-Datei"""
    try:
        df = pd.read_excel("Kistenliste.xlsx", sheet_name="Kistenliste")
        df["Name"] = df["Name"].str.strip()

        # Bezahlt-Status korrigieren
        df["Bezahlt"] = df["Bezahlt"].fillna("").str.strip()
        df["Bezahlt_Status"] = df["Bezahlt"].apply(
            lambda x: "Bezahlt" if x == "J" else "Offen"
        )

        return df
    except Exception as e:
        st.error(f"❌ Fehler beim Laden der Datei: {e}")
        return None


def create_person_chart(df):
    """Erstellt gestapeltes Balkendiagramm für Kisten pro Person"""
    sns.set_style("whitegrid")

    # Daten vorbereiten
    name_bezahlt = df[df["Bezahlt_Status"] == "Bezahlt"].groupby("Name").size()
    name_offen = df[df["Bezahlt_Status"] == "Offen"].groupby("Name").size()

    all_names = df["Name"].value_counts().index
    name_stats = pd.DataFrame(
        {
            "Name": all_names,
            "Bezahlt": [name_bezahlt.get(name, 0) for name in all_names],
            "Offen": [name_offen.get(name, 0) for name in all_names],
        }
    )

    name_stats["Gesamt"] = name_stats["Bezahlt"] + name_stats["Offen"]
    name_stats = name_stats.sort_values("Gesamt", ascending=True)

    # Diagramm für Kisten pro Person
    # Dynamische Höhe basierend auf Anzahl der Namen
    n_people = len(name_stats)
    height = max(6, n_people * 0.4)  # Minimum 6, sonst 0.4 pro Person
    fig, ax = plt.subplots(figsize=(10, height))
    fig.patch.set_facecolor("white")
    y_pos = range(len(name_stats))

    ax.barh(
        y_pos,
        name_stats["Bezahlt"],
        color="#16a34a",
        label="Bezahlt",
        edgecolor="darkgreen",
        linewidth=1.5,
    )
    ax.barh(
        y_pos,
        name_stats["Offen"],
        left=name_stats["Bezahlt"],
        color="#dc2626",
        label="Offen",
        edgecolor="darkred",
        linewidth=1.5,
    )

    ax.set_yticks(y_pos)
    ax.set_yticklabels(name_stats["Name"])
    ax.set_xlabel("Anzahl Kisten", fontweight="bold", fontsize=11)
    ax.set_ylabel("Name", fontweight="bold", fontsize=11)
    ax.legend(loc="lower right")
    ax.grid(axis="x", alpha=0.3)
    ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))

    # Werte anzeigen
    for i, row in enumerate(name_stats.itertuples()):
        ax.text(
            row.Gesamt + 0.1,
            i,
            str(row.Gesamt),
            va="center",
            fontweight="bold",
            fontsize=10,
        )

    plt.tight_layout()
    return fig


def create_payment_chart(df):
    """Erstellt Tortendiagramm für Bezahlstatus"""
    bezahlt_counts = df["Bezahlt_Status"].value_counts()

    fig, ax = plt.subplots(figsize=(5, 5))
    fig.patch.set_facecolor("white")

    colors_pie = ["#16a34a", "#dc2626"]
    explode = (0.05, 0.05)

    wedges, texts, autotexts = ax.pie(
        bezahlt_counts.values,
        labels=bezahlt_counts.index,
        autopct="%1.1f%%",
        colors=colors_pie,
        startangle=90,
        explode=explode,
        textprops={"fontsize": 11, "fontweight": "bold"},
        wedgeprops={"edgecolor": "white", "linewidth": 2},
    )

    plt.tight_layout()
    return fig


def create_reasons_chart(df):
    """Erstellt Balkendiagramm für häufigste Gründe"""
    sns.set_style("whitegrid")

    grund_counts = df["Grund"].value_counts().head(10).sort_values(ascending=True)

    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor("white")

    n_bars = len(grund_counts)
    colors_green = [
        plt.cm.Greens(0.5 + 0.5 * i / max(n_bars - 1, 1)) for i in range(n_bars)
    ]

    grund_counts.plot(
        kind="barh", ax=ax, color=colors_green, edgecolor="darkgreen", linewidth=1.5
    )

    ax.set_xlabel("Anzahl", fontweight="bold", fontsize=11)
    ax.set_ylabel("Grund", fontweight="bold", fontsize=11)
    ax.grid(axis="x", alpha=0.3)
    ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))

    # Werte anzeigen
    for i, v in enumerate(grund_counts.values):
        ax.text(v + 0.05, i, str(v), va="center", fontweight="bold", fontsize=10)

    plt.tight_layout()
    return fig


def create_ranking_table(df):
    """Erstellt Rangliste"""
    ranking = df["Name"].value_counts().reset_index()
    ranking.columns = ["Name", "Anzahl Kisten"]
    ranking["Rang"] = range(1, len(ranking) + 1)

    medals = {1: "🥇", 2: "🥈", 3: "🥉"}
    ranking["Medaille"] = ranking["Rang"].map(lambda x: medals.get(x, ""))

    return ranking[["Rang", "Medaille", "Name", "Anzahl Kisten"]]


# Hauptapp
def main():
    # Header
    st.title("⚽ Fc Münster 05 1. Mannschaft")
    st.markdown(
        '<p style="text-align: center; color: #6b7280; font-size: 18px;">Aktuelle Liste von offenen Bierkisten</p>',
        unsafe_allow_html=True,
    )

    # Timestamp
    st.markdown(
        f'<p style="text-align: center; color: #9ca3af; font-size: 12px;">Letzte Aktualisierung: {datetime.now().strftime("%d.%m.%Y %H:%M")}</p>',
        unsafe_allow_html=True,
    )

    st.markdown("---")

    # Daten laden
    df = load_data()

    if df is None:
        st.stop()

    # Statistiken berechnen
    bezahlt_count = len(df[df["Bezahlt_Status"] == "Bezahlt"])
    offen_count = len(df[df["Bezahlt_Status"] == "Offen"])
    personen_count = df["Name"].nunique()

    # Metriken anzeigen
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Gesamt Einträge", len(df))
    with col2:
        st.metric("Bezahlt", bezahlt_count, delta=None, delta_color="normal")
    with col3:
        st.metric("Offen", offen_count, delta=None, delta_color="inverse")
    with col4:
        st.metric("Personen", personen_count)

    st.markdown("---")

    # Diagramme
    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("🏆 Kisten pro Person")
        fig1 = create_person_chart(df)
        st.pyplot(fig1)

    with col_right:
        st.subheader("💰 Bezahlstatus")
        fig2 = create_payment_chart(df)
        st.pyplot(fig2)

    st.markdown("---")

    # Gründe
    st.subheader("📋 Top 10 Häufigste Gründe")
    fig3 = create_reasons_chart(df)
    st.pyplot(fig3)

    st.markdown("---")

    # Rangliste
    st.subheader("🏆 Rangliste")
    ranking = create_ranking_table(df)

    # Styling für Tabelle
    st.dataframe(
        ranking,
        hide_index=True,
        width="stretch",
        column_config={
            "Rang": st.column_config.NumberColumn("Rang", width="small"),
            "Medaille": st.column_config.TextColumn("", width="small"),
            "Name": st.column_config.TextColumn("Name", width="medium"),
            "Anzahl Kisten": st.column_config.NumberColumn(
                "Anzahl Kisten", width="small"
            ),
        },
    )

    # Footer
    st.markdown("---")
    st.markdown(
        '<p style="text-align: center; color: #6b7280; font-size: 14px;">💡 Die Seite aktualisiert sich automatisch bei Änderungen der Excel-Datei</p>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
