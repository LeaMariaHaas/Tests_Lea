"""
Dashboard-Modul des Controlling-Tools.

Erstellt ein interaktives Web-Dashboard mit Plotly Dash:
  - KPI-Cards mit Ampelfarben (YTD Budget, Ist, Abweichungen, Forecast)
  - Filter nach Kostenstelle, Kategorie, Monat
  - Balkendiagramm: monatlicher Plan/Ist-Vergleich
  - Wasserfall-Chart: kumulierte Abweichung YTD
  - Heatmap: Abweichung % nach Kostenstelle × Monat
  - Sortier-/filterbare Detailtabelle
  - Excel-Export-Button

Starten mit: python main.py --mode dashboard
"""

import datetime
import io
import logging
from typing import Optional

import dash
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Input, Output, State, callback, dash_table, dcc, html

import config
from src.transformer import (
    aggregate_by_month,
    calculate_full_year_forecast,
    calculate_plan_ist_comparison,
    calculate_ytd,
    get_traffic_light,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Farben & Stil
# ---------------------------------------------------------------------------
FARBE_BUDGET = "#2E75B6"    # Blau für Budget
FARBE_IST_OK = "#70AD47"    # Grün für Ist (innerhalb Budget)
FARBE_IST_KRITISCH = "#FF0000"  # Rot für Ist (Überschreitung)
FARBE_HEADER = "#1F3864"
AMPEL_FARBEN = {"🟢": "#C6EFCE", "🟡": "#FFEB9C", "🔴": "#FFC7CE", "⚪": "#F2F2F2"}


def _ampel_zu_farbe(ampel: str) -> str:
    return AMPEL_FARBEN.get(ampel, "#F2F2F2")


def _kpi_card(titel: str, wert: str, ampel: str = "⚪", untertitel: str = "") -> dbc.Card:
    """Erstellt eine KPI-Card-Komponente."""
    hintergrund = _ampel_zu_farbe(ampel)
    return dbc.Card(
        dbc.CardBody([
            html.P(titel, className="card-title text-muted", style={"fontSize": "0.8rem", "marginBottom": "4px"}),
            html.H4(wert, className="card-text fw-bold", style={"fontSize": "1.3rem"}),
            html.Small(untertitel, className="text-muted") if untertitel else html.Span(),
            html.Div(ampel, style={"fontSize": "1.2rem", "marginTop": "4px"}),
        ]),
        style={"backgroundColor": hintergrund, "border": "1px solid #dee2e6", "borderRadius": "8px"},
        className="shadow-sm h-100",
    )


def _formatiere_euro(betrag: float) -> str:
    """Formatiert einen Betrag als Euro-String."""
    return f"{betrag:+,.0f} {config.CURRENCY_SYMBOL}".replace(",", ".")


def _formatiere_euro_positiv(betrag: float) -> str:
    return f"{betrag:,.0f} {config.CURRENCY_SYMBOL}".replace(",", ".")


def _formatiere_prozent(pct: float) -> str:
    return f"{pct:+.1f} %"

# ---------------------------------------------------------------------------
# App-Erstellung
# ---------------------------------------------------------------------------

def create_app(budget_df: pd.DataFrame, actuals_df: pd.DataFrame) -> dash.Dash:
    """ 
    Erstellt die Dash-Applikation.

    Args:
        budget_df:   Budget-DataFrame.
        actuals_df:  Ist-DataFrame.

    Returns:
        Konfigurierte Dash-App (noch nicht gestartet).
    """
    app = dash.Dash(
        __name__,
        external_stylesheets=[dbc.themes.BOOTSTRAP],
        title="Controlling Dashboard",
    )

    # Optionen für Dropdown-Filter
    kostenstellen = sorted(budget_df["Kostenstelle"].unique().tolist())
    kategorien = sorted(budget_df["Kategorie"].unique().tolist())
    monate = sorted(budget_df["Monat"].unique().tolist())
    monatsnamen = {1: "Jan", 2: "Feb", 3: "Mär", 4: "Apr", 5: "Mai", 6: "Jun",
                   7: "Jul", 8: "Aug", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Dez"}

    # -----------------------------------------------------------------------
    # Layout
    # -----------------------------------------------------------------------
    app.layout = dbc.Container(fluid=True, children=[

        # Header
        dbc.Row(dbc.Col(html.Div([
            html.H2("📊 Controlling Dashboard", className="text-white mb-0"),
            html.P(f"Plan/Ist-Vergleich | Stichtag: {config.DEFAULT_AS_OF_DATE.strftime('%d.%m.%Y')}",
                   className="text-white-50 mb-0"),
        ], className="p-3"), style={"backgroundColor": FARBE_HEADER}), className="mb-3"),

        # Filter-Zeile
        dbc.Row([
            dbc.Col([
                html.Label("Kostenstelle", className="fw-bold"),
                dcc.Dropdown(
                    id="filter-kostenstelle",
                    options=[{"label": k, "value": k} for k in kostenstellen],
                    value=kostenstellen,
                    multi=True,
                    placeholder="Alle Kostenstellen...",
                ),
            ], md=4),
            dbc.Col([
                html.Label("Kategorie", className="fw-bold"),
                dcc.Dropdown(
                    id="filter-kategorie",
                    options=[{"label": k, "value": k} for k in kategorien],
                    value=kategorien,
                    multi=True,
                    placeholder="Alle Kategorien...",
                ),
            ], md=4),
            dbc.Col([
                html.Label("Monate (von – bis)", className="fw-bold"),
                dcc.RangeSlider(
                    id="filter-monat",
                    min=1, max=12, step=1,
                    value=[1, config.DEFAULT_AS_OF_DATE.month],
                    marks={m: monatsnamen[m] for m in range(1, 13)},
                    tooltip={"placement": "bottom"},
                ),
            ], md=4),
        ], className="mb-3 p-3 bg-light rounded"),

        # KPI-Cards
        dbc.Row(id="kpi-cards", className="mb-3 g-2"),

        # Charts Zeile 1
        dbc.Row([
            dbc.Col(dcc.Graph(id="chart-balken"), md=8),
            dbc.Col(dcc.Graph(id="chart-wasserfall"), md=4),
        ], className="mb-3"),

        # Charts Zeile 2
        dbc.Row([
            dbc.Col(dcc.Graph(id="chart-heatmap"), md=12),
        ], className="mb-3"),

        # Detailtabelle + Export
        dbc.Row([
            dbc.Col([
                dbc.Button(
                    "📥 Excel-Bericht exportieren",
                    id="btn-export",
                    color="success",
                    className="mb-2",
                ),
                dcc.Download(id="download-excel"),
                html.Div(id="export-status", className="text-muted small"),
                dash_table.DataTable(
                    id="detail-tabelle",
                    sort_action="native",
                    filter_action="native",
                    page_size=20,
                    style_table={"overflowX": "auto"},
                    style_header={
                        "backgroundColor": FARBE_HEADER,
                        "color": "white",
                        "fontWeight": "bold",
                        "textAlign": "center",
                    },
                    style_cell={"textAlign": "right", "fontSize": "12px", "padding": "6px"},
                    style_data_conditional=[
                        {"if": {"row_index": "odd"}, "backgroundColor": "#f8f9ff"},
                    ],
                ),
            ]),
        ]),

        # Versteckter Datenspeicher
        dcc.Store(id="store-budget", data=budget_df.to_json(orient="split")),
        dcc.Store(id="store-actuals", data=actuals_df.to_json(orient="split")),
    ], style={"maxWidth": "1600px", "margin": "0 auto"})

    # -----------------------------------------------------------------------
    # Callbacks
    # -----------------------------------------------------------------------

    @app.callback(
        Output("kpi-cards", "children"),
        Output("chart-balken", "figure"),
        Output("chart-wasserfall", "figure"),
        Output("chart-heatmap", "figure"),
        Output("detail-tabelle", "data"),
        Output("detail-tabelle", "columns"),
        Input("filter-kostenstelle", "value"),
        Input("filter-kategorie", "value"),
        Input("filter-monat", "value"),
        State("store-budget", "data"),
        State("store-actuals", "data"),
    )
    def aktualisiere_dashboard(
        kostenstelle_filter, kategorie_filter, monat_filter,
        budget_json, actuals_json
    ): 
        """Aktualisiert alle Dashboard-Komponenten basierend auf den Filtern."""
        # Daten aus Store lesen
        bdf = pd.read_json(io.StringIO(budget_json), orient="split")
        adf = pd.read_json(io.StringIO(actuals_json), orient="split")

        # Filter anwenden
        monat_von, monat_bis = monat_filter[0], monat_filter[1]
        stichtag = datetime.date(config.DEFAULT_AS_OF_DATE.year, monat_bis,
                                 config.DEFAULT_AS_OF_DATE.day
                                 if monat_bis == config.DEFAULT_AS_OF_DATE.month
                                 else 28)

        bdf_f = bdf[
            bdf["Kostenstelle"].isin(kostenstelle_filter) &
            bdf["Kategorie"].isin(kategorie_filter) &
            bdf["Monat"].between(monat_von, monat_bis)
        ]
        adf_f = adf[
            adf["Kostenstelle"].isin(kostenstelle_filter) &
            adf["Kategorie"].isin(kategorie_filter) &
            adf["Monat"].between(monat_von, monat_bis)
        ]

        if bdf_f.empty:
            leer = go.Figure()
            return [], leer, leer, leer, [], []

        comparison = calculate_plan_ist_comparison(bdf_f, adf_f, stichtag)
        ytd = calculate_ytd(bdf_f, adf_f, stichtag)
        monthly = aggregate_by_month(comparison)

        # --- KPI-Cards ---
        ytd_budget = ytd["YTD_Budget"].sum()
        ytd_ist = ytd["YTD_Ist"].sum()
        abw_abs = ytd_ist - ytd_budget
        abw_pct = (abw_abs / ytd_budget * 100) if ytd_budget != 0 else 0
        ampel = get_traffic_light(abw_pct)

        try:
            forecast_df = calculate_full_year_forecast(bdf_f, adf_f, stichtag)
            fc_gesamt = forecast_df["Forecast_Gesamtjahr"].sum()
            fc_budget = forecast_df["Budget_Gesamtjahr"].sum()
            fc_abw_abs = fc_gesamt - fc_budget
            fc_abw_pct = (fc_abw_abs / fc_budget * 100) if fc_budget != 0 else 0
            fc_ampel = get_traffic_light(fc_abw_pct)
        except Exception:
            fc_gesamt = fc_abw_abs = fc_abw_pct = 0
            fc_ampel = "⚪"

        kpi_cards = dbc.Row([
            dbc.Col(_kpi_card("YTD Budget", _formatiere_euro_positiv(ytd_budget)), md=2),
            dbc.Col(_kpi_card("YTD Ist", _formatiere_euro_positiv(ytd_ist)), md=2),
            dbc.Col(_kpi_card("Abweichung (€)", _formatiere_euro(abw_abs), ampel), md=2),
            dbc.Col(_kpi_card("Abweichung (%)", _formatiere_prozent(abw_pct), ampel), md=2),
            dbc.Col(_kpi_card("Forecast Gesamtjahr", _formatiere_euro_positiv(fc_gesamt), fc_ampel), md=2),
            dbc.Col(_kpi_card("Forecast Abweichung", _formatiere_euro(fc_abw_abs), fc_ampel), md=2),
        ], className="g-2").children

        # --- Balkendiagramm ---
        monatsnamen_de = {1: "Jan", 2: "Feb", 3: "Mär", 4: "Apr", 5: "Mai", 6: "Jun",
                          7: "Jul", 8: "Aug", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Dez"}
        monthly["MonatName"] = monthly["Monat"].map(monatsnamen_de)
        fig_balken = go.Figure()
        fig_balken.add_trace(go.Bar(
            name="Budget", x=monthly["MonatName"], y=monthly["Budget"],
            marker_color=FARBE_BUDGET, opacity=0.85,
        ))
        ist_farben = [FARBE_IST_OK if a >= 0 else FARBE_IST_KRITISCH
                      for a in monthly["Abweichung_absolut"]]
        fig_balken.add_trace(go.Bar(
            name="Ist", x=monthly["MonatName"], y=monthly["Ist"],
            marker_color=ist_farben, opacity=0.85,
        ))
        fig_balken.update_layout(
            title="Monatlicher Plan/Ist-Vergleich",
            barmode="group",
            xaxis_title="Monat",
            yaxis_title=f"Betrag ({config.CURRENCY_SYMBOL})",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            plot_bgcolor="white",
            paper_bgcolor="white",
        )

        # --- Wasserfall-Chart (kumulierte Abweichung) ---
        monthly_ytd = monthly[monthly["Ist"] > 0].copy()
        monthly_ytd["Kum_Abweichung"] = monthly_ytd["Abweichung_absolut"].cumsum()
        fig_wf = go.Figure(go.Waterfall(
            name="Kum. Abweichung",
            orientation="v",
            x=monthly_ytd["MonatName"],
            y=monthly_ytd["Abweichung_absolut"],
            connector={"line": {"color": "rgb(63, 63, 63)"}},
            increasing={"marker": {"color": FARBE_IST_OK}},
            decreasing={"marker": {"color": FARBE_IST_KRITISCH}},
        ))
        fig_wf.update_layout(
            title="Kumulierte Abweichung YTD",
            xaxis_title="Monat",
            yaxis_title=f"Abweichung ({config.CURRENCY_SYMBOL})",
            plot_bgcolor="white",
            paper_bgcolor="white",
            showlegend=False,
        )

        # --- Heatmap ---
        pivot = comparison[comparison["Ist"] > 0].pivot_table(
            values="Abweichung_prozent",
            index="Kostenstelle",
            columns="Monat",
            aggfunc="mean",
        )
        pivot.columns = [monatsnamen_de.get(m, str(m)) for m in pivot.columns]
        fig_heatmap = px.imshow(
            pivot,
            color_continuous_scale=["#C6EFCE", "#FFEB9C", "#FFC7CE"],
            color_continuous_midpoint=0,
            title="Abweichung (%) nach Kostenstelle × Monat",
            labels={"color": "Abw. (%)"},
            text_auto=".1f",
        )
        fig_heatmap.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            coloraxis_colorbar=dict(title="Abw. (%)"),
        )

        # --- Detailtabelle ---
        tabelle_df = comparison[comparison["Ist"] > 0][[
            "Jahr", "Monat", "Kostenstelle", "Kategorie",
            "Budget", "Ist", "Abweichung_absolut", "Abweichung_prozent", "Ampel"
        ]].copy()
        tabelle_df["Budget"] = tabelle_df["Budget"].map(lambda x: f"{x:,.0f} €")
        tabelle_df["Ist"] = tabelle_df["Ist"].map(lambda x: f"{x:,.0f} €")
        tabelle_df["Abweichung_absolut"] = tabelle_df["Abweichung_absolut"].map(lambda x: f"{x:+,.0f} €")
        tabelle_df["Abweichung_prozent"] = tabelle_df["Abweichung_prozent"].map(lambda x: f"{x:+.1f} %")

        spalten = [{"name": c, "id": c} for c in tabelle_df.columns]
        daten = tabelle_df.to_dict("records")

        return kpi_cards, fig_balken, fig_wf, fig_heatmap, daten, spalten

    @app.callback(
        Output("download-excel", "data"),
        Output("export-status", "children"),
        Input("btn-export", "n_clicks"),
        State("store-budget", "data"),
        State("store-actuals", "data"),
        prevent_initial_call=True,
    )
    def exportiere_excel(n_clicks, budget_json, actuals_json):
        """Exportiert den Excel-Bericht und stellt ihn zum Download bereit."""
        if not n_clicks:
            return dash.no_update, ""
        try:
            from src.excel_exporter import export_to_excel
            bdf = pd.read_json(io.StringIO(budget_json), orient="split")
            adf = pd.read_json(io.StringIO(actuals_json), orient="split")
            ausgabe = export_to_excel(bdf, adf)
            return dcc.send_file(str(ausgabe)), f"✅ Exportiert: {ausgabe.name}"
        except Exception as e:
            logger.error(f"Excel-Export fehlgeschlagen: {e}")
            return dash.no_update, f"❌ Fehler beim Export: {e}"

    return app


def run_dashboard(
    budget_df: pd.DataFrame,
    actuals_df: pd.DataFrame,
    port: Optional[int] = None,
) -> None:
    """Startet das Dash-Dashboard.

    Args:
        budget_df:   Budget-DataFrame.
        actuals_df:  Ist-DataFrame.
        port:        Port; Standard: config.DASHBOARD_PORT.
    """
    app_port = port or config.DASHBOARD_PORT
    app = create_app(budget_df, actuals_df)
    print(f"\n🚀 Dashboard gestartet: http://localhost:{app_port}\n")
    app.run(
        debug=False,
        port=app_port,
        open_browser=config.DASHBOARD_OPEN_BROWSER,
    )