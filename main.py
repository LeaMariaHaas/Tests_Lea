"""
Einstiegspunkt des Controlling-Tools.

Verwendung:
    python main.py --mode dashboard
    python main.py --mode excel
    python main.py --mode both
    python main.py --budget data/mein_budget.csv --actuals data/meine_ist.csv
    python main.py --as-of 2026-03-31
"""

import argparse
import datetime
import logging
import sys
from pathlib import Path

# Logging konfigurieren
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

def parse_argumente():
    """Parst die Kommandozeilenargumente."""
    parser = argparse.ArgumentParser(
        description="Controlling-Tool: ERP-Zahlen → Plan/Ist-Vergleich mit Excel & Dashboard",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Beispiele:
  python main.py --mode dashboard
  python main.py --mode excel
  python main.py --mode both --as-of 2026-03-31
  python main.py --mode excel --budget data/budget.csv --actuals data/actuals.csv
        """
    )

    parser.add_argument(
        "--mode",
        choices=["dashboard", "excel", "both"],
        default="dashboard",
        help="Ausgabemodus: 'dashboard' (Web-UI), 'excel' (Excel-Bericht) oder 'both' (beides). Standard: dashboard"
    )
    parser.add_argument(
        "--budget",
        type=str,
        default=None,
        help="Pfad zur Budget-Datei (CSV oder Excel). Standard: data/sample_budget.csv"
    )
    parser.add_argument(
        "--actuals",
        type=str,
        default=None,
        help="Pfad zur Ist-Daten-Datei (CSV oder Excel). Standard: data/sample_actuals.csv"
    )
    parser.add_argument(
        "--as-of",
        type=str,
        default=None,
        dest="as_of",
        help="Stichtag im Format YYYY-MM-DD. Standard: heutiges Datum"
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Ausgabepfad für den Excel-Bericht. Standard: output/controlling_bericht.xlsx"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=None,
        help="Port für das Dashboard. Standard: 8050"
    )
    parser.add_argument(
        "--erp",
        action="store_true",
        default=False,
        help="Daten aus ERP-Stub laden (statt lokaler Dateien)"
    )

    return parser.parse_args()

def main():
    """Hauptfunktion des Controlling-Tools."""
    args = parse_argumente()

    # Stichtag verarbeiten
    stichtag = None
    if args.as_of:
        try:
            stichtag = datetime.date.fromisoformat(args.as_of)
            logger.info(f"Stichtag aus Argument: {stichtag}")
        except ValueError:
            logger.error(f"Ungültiges Datum: '{args.as_of}'. Bitte Format YYYY-MM-DD verwenden.")
            sys.exit(1)

    # Daten laden
    logger.info("Lade Daten...")
    try:
        from src.data_loader import load_budget, load_actuals, load_from_erp_stub

        if args.erp:
            logger.info("Lade Daten aus ERP-System (Stub)...")
            budget_df, actuals_df = load_from_erp_stub()
        else:
            budget_df = load_budget(args.budget)
            actuals_df = load_actuals(args.actuals)

        logger.info(
            f"Daten geladen: {len(budget_df)} Budget-Zeilen, {len(actuals_df)} Ist-Zeilen"
        )

    except FileNotFoundError as e:
        logger.error(f"Datei nicht gefunden: {e}")
        sys.exit(1)
    except ValueError as e:
        logger.error(f"Fehler beim Laden der Daten: {e}")
        sys.exit(1)

    # Ausgabe je nach Modus
    if args.mode in ("excel", "both"):
        logger.info("Erstelle Excel-Bericht...")
        try:
            from src.excel_exporter import export_to_excel
            ausgabepfad = export_to_excel(
                budget_df=budget_df,
                actuals_df=actuals_df,
                as_of_date=stichtag,
                output_path=args.output,
            )
            print(f"\n✅ Excel-Bericht erstellt: {ausgabepfad.resolve()}\n")
        except Exception as e:
            logger.error(f"Fehler beim Excel-Export: {e}")
            if args.mode == "excel":
                sys.exit(1)

    if args.mode in ("dashboard", "both"):
        logger.info("Starte Dashboard...")
        try:
            from src.dashboard import run_dashboard
            run_dashboard(
                budget_df=budget_df,
                actuals_df=actuals_df,
                port=args.port,
            )
        except ImportError as e:
            logger.error(
                f"Fehler beim Starten des Dashboards: {e}\n"
                f"Bitte prüfen Sie, ob alle Abhängigkeiten installiert sind: pip install -r requirements.txt"
            )
            sys.exit(1)
        except Exception as e:
            logger.error(f"Unerwarteter Fehler beim Dashboard: {e}")
            sys.exit(1)


if __name__ == "__main__":
    main()