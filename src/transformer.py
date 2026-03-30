"""
Transformations-Modul – Kernlogik des Controlling-Tools.

Enthält alle Berechnungen:
  - Plan/Ist-Vergleich mit zeitlicher Filterung (YTD)
  - Abweichungen (absolut und prozentual)
  - Jahres-Forecast (Hochrechnung)
  - Aggregationen nach Kostenstelle, Kategorie, Monat
  - Ampel-Bewertung der Abweichungen
"""

import datetime
import logging
from typing import Optional

import numpy as np
import pandas as pd

import config

logger = logging.getLogger(__name__)

# Spaltennamen als Konstanten (verhindert Tippfehler)
COL_JAHR = "Jahr"
COL_MONAT = "Monat"
COL_KOSTENSTELLE = "Kostenstelle"
COL_KATEGORIE = "Kategorie"
COL_BETRAG = "Betrag"
COL_BUDGET = "Budget"
COL_IST = "Ist"
COL_ABWEICHUNG_ABS = "Abweichung_absolut"
COL_ABWEICHUNG_PCT = "Abweichung_prozent"
COL_AMPEL = "Ampel"
COL_FORECAST = "Forecast_Gesamtjahr"
COL_BUDGET_GESAMT = "Budget_Gesamtjahr"
COL_FORECAST_ABW_ABS = "Forecast_Abweichung_absolut"
COL_FORECAST_ABW_PCT = "Forecast_Abweichung_prozent"


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def _get_stichtag(as_of_date: Optional[datetime.date]) -> datetime.date:
    """Gibt den Stichtag zurück; Standard ist config.DEFAULT_AS_OF_DATE."""
    return as_of_date if as_of_date is not None else config.DEFAULT_AS_OF_DATE


def _ytd_monate(as_of_date: datetime.date, jahr: int) -> list:
    """
    Gibt die Liste der Monate zurück, die bis zum Stichtag im gegebenen Jahr
    bereits abgeschlossen sind (d.h. vollständig vergangen).

    Beispiel: Stichtag 2026-03-30 → Monate [1, 2, 3] für Jahr 2026
              (März gilt als abgeschlossen, weil der Stichtag am 30.3. liegt)

    Args:
        as_of_date: Stichtag.
        jahr:       Das Geschäftsjahr.

    Returns:
        Liste der YTD-Monate (1-basiert).
    """
    if as_of_date.year > jahr:
        # Gesamtes Jahr ist vergangen
        return list(range(1, 13))
    elif as_of_date.year < jahr:
        # Zukünftiges Jahr – noch keine Ist-Daten
        return []
    else:
        # Laufendes Jahr: Monate 1 bis einschließlich aktuellem Monat
        return list(range(1, as_of_date.month + 1))


# ---------------------------------------------------------------------------
# Kernfunktionen
# ---------------------------------------------------------------------------

def calculate_plan_ist_comparison(
    budget_df: pd.DataFrame,
    actuals_df: pd.DataFrame,
    as_of_date: Optional[datetime.date] = None,
) -> pd.DataFrame:
    """
    Berechnet den Plan/Ist-Vergleich auf Monats-/Kostenstellen-/Kategorie-Ebene.

    Zeitliche Logik:
      - Nur Monate bis zum Stichtag werden für den Ist-Vergleich herangezogen.
      - Budget-Zeilen außerhalb des YTD-Zeitraums erhalten Ist = 0.

    Args:
        budget_df:   Budget-DataFrame (Jahres-Budget).
        actuals_df:  Ist-DataFrame (nur bereits gebuchte Monate).
        as_of_date:  Stichtag; Standard: heutiges Datum.

    Returns:
        DataFrame mit Spalten:
        Jahr, Monat, Kostenstelle, Kategorie, Budget, Ist,
        Abweichung_absolut, Abweichung_prozent, Ampel
    """
    stichtag = _get_stichtag(as_of_date)
    logger.info(f"Plan/Ist-Vergleich mit Stichtag: {stichtag}")

    if budget_df.empty:
        raise ValueError("Das Budget-DataFrame ist leer.")

    # Umbenennen, damit Merge eindeutig ist
    bdf = budget_df.rename(columns={COL_BETRAG: COL_BUDGET})
    adf = actuals_df.rename(columns={COL_BETRAG: COL_IST}) if not actuals_df.empty else pd.DataFrame(
        columns=[COL_JAHR, COL_MONAT, COL_KOSTENSTELLE, COL_KATEGORIE, COL_IST]
    )

    # Merge: alle Budget-Zeilen behalten, Ist von links einjoinen
    merge_keys = [COL_JAHR, COL_MONAT, COL_KOSTENSTELLE, COL_KATEGORIE]
    df = bdf.merge(adf[merge_keys + [COL_IST]], on=merge_keys, how="left")

    # Zeitliche Filterung: Ist-Werte nur für YTD-Monate gültig
    jahre = df[COL_JAHR].unique()
    ytd_mask = pd.Series(False, index=df.index)
    for jahr in jahre:
        monate = _ytd_monate(stichtag, int(jahr))
        ytd_mask |= (df[COL_JAHR] == jahr) & (df[COL_MONAT].isin(monate))

    # Außerhalb YTD: Ist auf NaN lassen (kein Vergleich möglich)
    df.loc[~ytd_mask, COL_IST] = np.nan
    df[COL_IST] = df[COL_IST].fillna(0.0)

    # Abweichungen berechnen
    df[COL_ABWEICHUNG_ABS] = df[COL_IST] - df[COL_BUDGET]

    # Prozentuale Abweichung (sicher gegen Division durch 0)
    df[COL_ABWEICHUNG_PCT] = np.where(
        df[COL_BUDGET] != 0,
        (df[COL_IST] - df[COL_BUDGET]) / df[COL_BUDGET].abs() * 100,
        np.where(df[COL_IST] != 0, 100.0, 0.0),
    )

    # Ampel hinzufügen
    df[COL_AMPEL] = df[COL_ABWEICHUNG_PCT].apply(get_traffic_light)

    logger.info(f"Plan/Ist-Vergleich erstellt: {len(df)} Zeilen.")
    return df.reset_index(drop=True)


def calculate_ytd(
    budget_df: pd.DataFrame,
    actuals_df: pd.DataFrame,
    as_of_date: Optional[datetime.date] = None,
) -> pd.DataFrame:
    """
    Berechnet die YTD-Aggregation (Year-to-Date) nach Kostenstelle und Kategorie.

    Nur Monate vom Jahresbeginn bis zum Stichtag fließen ein.

    Args:
        budget_df:   Budget-DataFrame.
        actuals_df:  Ist-DataFrame.
        as_of_date:  Stichtag.

    Returns:
        DataFrame mit Kostenstelle, Kategorie,
        YTD_Budget, YTD_Ist, Abweichung_absolut, Abweichung_prozent, Ampel
    """
    stichtag = _get_stichtag(as_of_date)
    df = calculate_plan_ist_comparison(budget_df, actuals_df, stichtag)

    # Nur YTD-Monate (Ist != 0 oder Budget vorhanden mit abgeschlossenem Monat)
    jahre = df[COL_JAHR].unique()
    ytd_mask = pd.Series(False, index=df.index)
    for jahr in jahre:
        monate = _ytd_monate(stichtag, int(jahr))
        ytd_mask |= (df[COL_JAHR] == jahr) & (df[COL_MONAT].isin(monate))

    ytd_df = df[ytd_mask].copy()

    agg = ytd_df.groupby([COL_KOSTENSTELLE, COL_KATEGORIE], as_index=False).agg(
        YTD_Budget=(COL_BUDGET, "sum"),
        YTD_Ist=(COL_IST, "sum"),
    )
    agg[COL_ABWEICHUNG_ABS] = agg["YTD_Ist"] - agg["YTD_Budget"]
    agg[COL_ABWEICHUNG_PCT] = np.where(
        agg["YTD_Budget"] != 0,
        (agg["YTD_Ist"] - agg["YTD_Budget"]) / agg["YTD_Budget"].abs() * 100,
        np.where(agg["YTD_Ist"] != 0, 100.0, 0.0),
    )
    agg[COL_AMPEL] = agg[COL_ABWEICHUNG_PCT].apply(get_traffic_light)

    logger.info(f"YTD berechnet bis Stichtag {stichtag}.")
    return agg.reset_index(drop=True)


def calculate_full_year_forecast(
    budget_df: pd.DataFrame,
    actuals_df: pd.DataFrame,
    as_of_date: Optional[datetime.date] = None,
) -> pd.DataFrame:
    """
    Berechnet die Jahreshochrechnung (Forecast):
        Forecast = Ist YTD + verbleibendes Budget (Monate nach Stichtag)

    Args:
        budget_df:   Budget-DataFrame (Jahres-Budget).
        actuals_df:  Ist-DataFrame.
        as_of_date:  Stichtag.

    Returns:
        DataFrame mit Kostenstelle, Kategorie,
        Budget_Gesamtjahr, YTD_Ist, Restbudget,
        Forecast_Gesamtjahr, Forecast_Abweichung_absolut, Forecast_Abweichung_prozent
    """
    stichtag = _get_stichtag(as_of_date)

    # Jahres-Budget gesamt (alle 12 Monate)
    budget_gesamt = budget_df.groupby([COL_KOSTENSTELLE, COL_KATEGORIE], as_index=False).agg(
        Budget_Gesamtjahr=(COL_BETRAG, "sum")
    )

    # YTD Ist
    ytd = calculate_ytd(budget_df, actuals_df, stichtag)
    ytd = ytd[[COL_KOSTENSTELLE, COL_KATEGORIE, "YTD_Ist"]].copy()

    # Restbudget (Monate nach Stichtag)
    jahre = budget_df[COL_JAHR].unique()
    rest_mask = pd.Series(False, index=budget_df.index)
    for jahr in jahre:
        ytd_monate = _ytd_monate(stichtag, int(jahr))
        rest_mask |= (budget_df[COL_JAHR] == jahr) & (~budget_df[COL_MONAT].isin(ytd_monate))

    rest_budget = budget_df[rest_mask].groupby(
        [COL_KOSTENSTELLE, COL_KATEGORIE], as_index=False
    ).agg(Restbudget=(COL_BETRAG, "sum"))

    # Zusammenführen
    fc = budget_gesamt.merge(ytd, on=[COL_KOSTENSTELLE, COL_KATEGORIE], how="left")
    fc = fc.merge(rest_budget, on=[COL_KOSTENSTELLE, COL_KATEGORIE], how="left")
    fc["YTD_Ist"] = fc["YTD_Ist"].fillna(0.0)
    fc["Restbudget"] = fc["Restbudget"].fillna(0.0)

    fc[COL_FORECAST] = fc["YTD_Ist"] + fc["Restbudget"]
    fc[COL_FORECAST_ABW_ABS] = fc[COL_FORECAST] - fc["Budget_Gesamtjahr"]
    fc[COL_FORECAST_ABW_PCT] = np.where(
        fc["Budget_Gesamtjahr"] != 0,
        (fc[COL_FORECAST] - fc["Budget_Gesamtjahr"]) / fc["Budget_Gesamtjahr"].abs() * 100,
        np.where(fc[COL_FORECAST] != 0, 100.0, 0.0),
    )
    fc[COL_AMPEL] = fc[COL_FORECAST_ABW_PCT].apply(get_traffic_light)

    logger.info("Jahres-Forecast berechnet.")
    return fc.reset_index(drop=True)

# ---------------------------------------------------------------------------
# Aggregationsfunktionen
# ---------------------------------------------------------------------------

def aggregate_by_costcenter(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregiert den Plan/Ist-DataFrame nach Kostenstelle."""
    return df.groupby(COL_KOSTENSTELLE, as_index=False).agg(
        Budget=(COL_BUDGET, "sum"),
        Ist=(COL_IST, "sum"),
        Abweichung_absolut=(COL_ABWEICHUNG_ABS, "sum"),
    ).assign(
        Abweichung_prozent=lambda x: np.where(
            x["Budget"] != 0,
            x["Abweichung_absolut"] / x["Budget"].abs() * 100,
            0.0,
        ),
        Ampel=lambda x: x["Abweichung_prozent"].apply(get_traffic_light),
    )

def aggregate_by_category(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregiert den Plan/Ist-DataFrame nach Kategorie."""
    return df.groupby(COL_KATEGORIE, as_index=False).agg(
        Budget=(COL_BUDGET, "sum"),
        Ist=(COL_IST, "sum"),
        Abweichung_absolut=(COL_ABWEICHUNG_ABS, "sum"),
    ).assign(
        Abweichung_prozent=lambda x: np.where(
            x["Budget"] != 0,
            x["Abweichung_absolut"] / x["Budget"].abs() * 100,
            0.0,
        ),
        Ampel=lambda x: x["Abweichung_prozent"].apply(get_traffic_light),
    )

def aggregate_by_month(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregiert den Plan/Ist-DataFrame nach Monat (Zeitreihe)."""
    return df.groupby([COL_JAHR, COL_MONAT], as_index=False).agg(
        Budget=(COL_BUDGET, "sum"),
        Ist=(COL_IST, "sum"),
        Abweichung_absolut=(COL_ABWEICHUNG_ABS, "sum"),
    ).assign(
        Abweichung_prozent=lambda x: np.where(
            x["Budget"] != 0,
            x["Abweichung_absolut"] / x["Budget"].abs() * 100,
            0.0,
        ),
        Ampel=lambda x: x["Abweichung_prozent"].apply(get_traffic_light),
    ).sort_values([COL_JAHR, COL_MONAT]).reset_index(drop=True)

# ---------------------------------------------------------------------------
# Ampel-Bewertung
# ---------------------------------------------------------------------------

def get_traffic_light(abweichung_prozent: float) -> str:
    """
    Bewertet eine prozentuale Abweichung mit einer Ampelfarbe.

    Schwellenwerte sind in config.py konfigurierbar:
      - |Abw.| < GREEN_THRESHOLD  → 🟢 (unkritisch)
      - |Abw.| < YELLOW_THRESHOLD → 🟡 (Achtung)
      - sonst                      → 🔴 (kritisch)

    Args:
        abweichung_prozent: Prozentuale Abweichung (positiv = Überschreitung).

    Returns:
        Emoji-String: '🟢', '🟡' oder '🔴'
    """
    if pd.isna(abweichung_prozent):
        return "⚪"  # Kein Vergleich möglich (kein Budget)

    abs_abw = abs(abweichung_prozent)
    if abs_abw < config.TRAFFIC_LIGHT_GREEN_THRESHOLD:
        return "🟢"
    elif abs_abw < config.TRAFFIC_LIGHT_YELLOW_THRESHOLD:
        return "🟡"
    else:
        return "🔴"