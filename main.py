"""Cash Flow Projection Tool — CLI entry point."""

from calculator import compute_financials
from excel_generator import generate_excel


# ---------------------------------------------------------------------------
# Input helpers
# ---------------------------------------------------------------------------

def prompt_str(prompt, allow_empty=False):
    """Prompt for a non-empty string (unless allow_empty=True)."""
    while True:
        value = input(prompt).strip()
        if value or allow_empty:
            return value
        print("  Please enter a value.")


def prompt_int(prompt, min_val=None, max_val=None):
    """Prompt for an integer within optional bounds."""
    while True:
        raw = input(prompt).strip()
        try:
            value = int(raw)
        except ValueError:
            print(f"  Please enter a whole number.")
            continue
        if min_val is not None and value < min_val:
            print(f"  Value must be at least {min_val}.")
            continue
        if max_val is not None and value > max_val:
            print(f"  Value must be at most {max_val}.")
            continue
        return value


def prompt_float(prompt, min_val=None, max_val=None, allow_negative=False, default=None):
    """
    Prompt for a float within optional bounds.
    If `default` is provided and the user presses Enter, the default is used.
    """
    while True:
        raw = input(prompt).strip()
        if raw == "" and default is not None:
            return default
        try:
            value = float(raw)
        except ValueError:
            print("  Please enter a numeric value.")
            continue
        if not allow_negative and value < 0:
            print("  Value cannot be negative.")
            continue
        if min_val is not None and value < min_val:
            print(f"  Value must be at least {min_val}.")
            continue
        if max_val is not None and value > max_val:
            print(f"  Value must be at most {max_val}.")
            continue
        return value


# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------

def collect_project_metadata():
    """Prompt for high-level project info. Returns a dict."""
    print()
    company_name = prompt_str("Company / Project Name: ")
    num_years = prompt_int("Number of projection years (1–50): ", min_val=1, max_val=50)
    wacc_pct = prompt_float("WACC / Discount rate (%): ", min_val=0.01, max_val=100)
    initial_investment = prompt_float("Initial Investment / CapEx at Year 0 ($): ", min_val=0)
    filename_raw = input("Output filename [cashflow.xlsx]: ").strip()
    filename = filename_raw if filename_raw else "cashflow.xlsx"
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
    """Prompt for one year's financial inputs. Returns a dict."""
    print(f"\n--- Year {year_num} ---")
    revenue   = prompt_float("  Revenue ($): ", min_val=0)
    cogs      = prompt_float("  COGS ($): ", min_val=0)
    opex      = prompt_float("  Operating Expenses – SG&A ($): ", min_val=0)
    da        = prompt_float("  Depreciation & Amortization ($): ", min_val=0)
    interest  = prompt_float("  Interest Expense ($): ", min_val=0)
    tax_rate  = prompt_float("  Tax Rate (%): ", min_val=0, max_val=100)
    capex     = prompt_float("  CapEx this year ($) [0]: ", min_val=0, default=0)
    delta_wc  = prompt_float(
        "  Change in Working Capital ($) [0]: ", allow_negative=True, default=0
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
    """Collect data for all years. Returns list of dicts."""
    return [collect_single_year(year) for year in range(1, num_years + 1)]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=== Cash Flow Projection Tool ===")

    metadata = collect_project_metadata()

    yearly_data = collect_yearly_data(metadata["num_years"])

    project_data = {
        "company_name": metadata["company_name"],
        "num_years": metadata["num_years"],
        "wacc": metadata["wacc"],
        "initial_investment": metadata["initial_investment"],
        "yearly_data": yearly_data,
    }

    print("\nCalculating financials…")
    results = compute_financials(project_data)

    # Save with re-prompt on PermissionError (e.g. file open in Excel)
    filename = metadata["filename"]
    while True:
        try:
            generate_excel(project_data, results, filename)
            break
        except PermissionError:
            print(
                f"\n  Could not write '{filename}' — the file may be open in Excel.\n"
                "  Please close it and try again, or enter a different filename."
            )
            new_name = input("  Output filename: ").strip()
            if new_name:
                filename = new_name if new_name.endswith(".xlsx") else new_name + ".xlsx"

    print(f"\nDone. File saved as: {filename}")


if __name__ == "__main__":
    main()
