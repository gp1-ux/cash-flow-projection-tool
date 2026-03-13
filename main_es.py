"""Herramienta de Proyección de Flujo de Caja — Interfaz de línea de comandos (Español)."""

from calculator import compute_financials
from excel_generator_es import generate_excel


# ---------------------------------------------------------------------------
# Ayudantes de entrada
# ---------------------------------------------------------------------------

def prompt_str(prompt, allow_empty=False):
    """Solicita una cadena no vacía (salvo que allow_empty=True)."""
    while True:
        value = input(prompt).strip()
        if value or allow_empty:
            return value
        print("  Por favor, ingrese un valor.")


def prompt_int(prompt, min_val=None, max_val=None):
    """Solicita un número entero dentro de los límites opcionales."""
    while True:
        raw = input(prompt).strip()
        try:
            value = int(raw)
        except ValueError:
            print("  Por favor, ingrese un número entero.")
            continue
        if min_val is not None and value < min_val:
            print(f"  El valor debe ser al menos {min_val}.")
            continue
        if max_val is not None and value > max_val:
            print(f"  El valor debe ser como máximo {max_val}.")
            continue
        return value


def prompt_float(prompt, min_val=None, max_val=None, allow_negative=False, default=None):
    """
    Solicita un número decimal dentro de los límites opcionales.
    Si se proporciona `default` y el usuario presiona Enter, se usa el valor por defecto.
    """
    while True:
        raw = input(prompt).strip()
        if raw == "" and default is not None:
            return default
        try:
            value = float(raw.replace(",", "."))
        except ValueError:
            print("  Por favor, ingrese un valor numérico.")
            continue
        if not allow_negative and value < 0:
            print("  El valor no puede ser negativo.")
            continue
        if min_val is not None and value < min_val:
            print(f"  El valor debe ser al menos {min_val}.")
            continue
        if max_val is not None and value > max_val:
            print(f"  El valor debe ser como máximo {max_val}.")
            continue
        return value


# ---------------------------------------------------------------------------
# Recopilación de datos
# ---------------------------------------------------------------------------

def collect_project_metadata():
    """Solicita información general del proyecto. Devuelve un dict."""
    print()
    company_name = prompt_str("Nombre de la empresa / proyecto: ")
    num_years = prompt_int("Número de años de proyección (1–50): ", min_val=1, max_val=50)
    wacc_pct = prompt_float("WACC / Tasa de descuento (%): ", min_val=0.01, max_val=100)
    initial_investment = prompt_float("Inversión inicial / CapEx en Año 0 ($): ", min_val=0)
    filename_raw = input("Nombre del archivo de salida [flujo_caja.xlsx]: ").strip()
    filename = filename_raw if filename_raw else "flujo_caja.xlsx"
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"

    return {
        "company_name": company_name,
        "num_years": num_years,
        "wacc": wacc_pct / 100.0,
        "initial_investment": initial_investment,
        "filename": filename,
    }


def collect_single_year(year_num):
    """Solicita los datos financieros de un año. Devuelve un dict."""
    print(f"\n--- Año {year_num} ---")
    revenue  = prompt_float("  Ingresos ($): ", min_val=0)
    cogs     = prompt_float("  Costo de ventas – COGS ($): ", min_val=0)
    opex     = prompt_float("  Gastos operativos – SG&A ($): ", min_val=0)
    da       = prompt_float("  Depreciación y Amortización ($): ", min_val=0)
    interest = prompt_float("  Gasto por intereses ($): ", min_val=0)
    tax_rate = prompt_float("  Tasa impositiva (%): ", min_val=0, max_val=100)
    capex    = prompt_float("  CapEx del año ($) [0]: ", min_val=0, default=0)
    delta_wc = prompt_float(
        "  Variación en capital de trabajo ($) [0]: ", allow_negative=True, default=0
    )

    return {
        "revenue": revenue,
        "cogs": cogs,
        "opex": opex,
        "da": da,
        "interest": interest,
        "tax_rate": tax_rate / 100.0,
        "capex": capex,
        "delta_wc": delta_wc,
    }


def collect_yearly_data(num_years):
    """Recopila datos para todos los años. Devuelve una lista de dicts."""
    return [collect_single_year(year) for year in range(1, num_years + 1)]


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def main():
    print("=== Herramienta de Proyección de Flujo de Caja ===")

    metadata = collect_project_metadata()
    yearly_data = collect_yearly_data(metadata["num_years"])

    project_data = {
        "company_name": metadata["company_name"],
        "num_years": metadata["num_years"],
        "wacc": metadata["wacc"],
        "initial_investment": metadata["initial_investment"],
        "yearly_data": yearly_data,
    }

    print("\nCalculando métricas financieras…")
    results = compute_financials(project_data)

    filename = metadata["filename"]
    while True:
        try:
            generate_excel(project_data, results, filename)
            break
        except PermissionError:
            print(
                f"\n  No se pudo guardar '{filename}' — el archivo puede estar abierto en Excel.\n"
                "  Por favor, ciérrelo e intente de nuevo, o ingrese un nombre de archivo diferente."
            )
            new_name = input("  Nombre del archivo de salida: ").strip()
            if new_name:
                filename = new_name if new_name.endswith(".xlsx") else new_name + ".xlsx"

    print(f"\nListo. Archivo guardado como: {filename}")


if __name__ == "__main__":
    main()
