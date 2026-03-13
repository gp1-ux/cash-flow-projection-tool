"""Financial calculations for the Cash Flow Projection Tool."""

try:
    import numpy_financial as npf
    _NPF_AVAILABLE = True
except ImportError:
    _NPF_AVAILABLE = False


def safe_divide(numerator, denominator, default=0.0):
    """Avoids ZeroDivisionError."""
    if denominator == 0:
        return default
    return numerator / denominator


def compute_year_metrics(year_data, wacc, year_num):
    """Compute all financial metrics for a single year."""
    revenue = year_data["revenue"]
    cogs = year_data["cogs"]
    opex = year_data["opex"]
    da = year_data["da"]
    interest = year_data["interest"]
    tax_rate = year_data["tax_rate"]
    capex = year_data["capex"]
    delta_wc = year_data["delta_wc"]

    gross_profit = revenue - cogs
    gross_margin_pct = safe_divide(gross_profit, revenue)

    ebitda = gross_profit - opex
    ebitda_margin_pct = safe_divide(ebitda, revenue)

    ebit = ebitda - da
    ebt = ebit - interest

    tax = ebt * tax_rate if ebt > 0 else 0.0
    net_income = ebt - tax
    net_margin_pct = safe_divide(net_income, revenue)

    fcf = net_income + da - capex - delta_wc
    discounted_fcf = fcf / (1 + wacc) ** year_num

    return {
        "year": year_num,
        "revenue": revenue,
        "cogs": cogs,
        "gross_profit": gross_profit,
        "gross_margin_pct": gross_margin_pct,
        "opex": opex,
        "ebitda": ebitda,
        "ebitda_margin_pct": ebitda_margin_pct,
        "da": da,
        "ebit": ebit,
        "interest": interest,
        "ebt": ebt,
        "tax": tax,
        "net_income": net_income,
        "net_margin_pct": net_margin_pct,
        "da_addback": da,
        "capex": capex,
        "delta_wc": delta_wc,
        "fcf": fcf,
        "discounted_fcf": discounted_fcf,
        # cumulative fields are filled in by compute_financials
        "cumulative_fcf": None,
        "cumulative_discounted_fcf": None,
    }


def compute_npv(initial_investment, fcf_list, wacc):
    """Compute NPV. Uses numpy_financial if available, else manual sum."""
    if _NPF_AVAILABLE:
        cashflows = [-initial_investment] + fcf_list
        return float(npf.npv(wacc, cashflows))
    # Manual fallback
    pv_sum = sum(fcf / (1 + wacc) ** i for i, fcf in enumerate(fcf_list, start=1))
    return pv_sum - initial_investment


def compute_irr(initial_investment, fcf_list):
    """Compute IRR. Uses numpy_financial if available, else Newton-Raphson."""
    cashflows = [-initial_investment] + fcf_list
    if _NPF_AVAILABLE:
        try:
            result = npf.irr(cashflows)
            if result is None or result != result:  # NaN check
                return None
            return float(result)
        except Exception:
            return None
    return _irr_newton_raphson(cashflows)


def _irr_newton_raphson(cashflows, guess=0.1, tol=1e-6, max_iter=1000):
    """Manual IRR solver using Newton-Raphson method."""
    rate = guess
    for _ in range(max_iter):
        npv_val = sum(cf / (1 + rate) ** i for i, cf in enumerate(cashflows))
        dnpv = sum(-i * cf / (1 + rate) ** (i + 1) for i, cf in enumerate(cashflows))
        if dnpv == 0:
            return None
        new_rate = rate - npv_val / dnpv
        if abs(new_rate - rate) < tol:
            return new_rate
        rate = new_rate
    return None


def compute_simple_payback(initial_investment, fcf_list):
    """Compute simple payback period with linear interpolation between years."""
    cumulative = 0.0
    for i, fcf in enumerate(fcf_list):
        prev_cumulative = cumulative
        cumulative += fcf
        if cumulative >= initial_investment:
            fraction = (initial_investment - prev_cumulative) / fcf if fcf != 0 else 0
            return i + fraction  # years (0-based index + fraction = year number)
    return None


def compute_discounted_payback(initial_investment, discounted_fcf_list):
    """Compute discounted payback period with linear interpolation."""
    cumulative = 0.0
    for i, dfcf in enumerate(discounted_fcf_list):
        prev_cumulative = cumulative
        cumulative += dfcf
        if cumulative >= initial_investment:
            fraction = (initial_investment - prev_cumulative) / dfcf if dfcf != 0 else 0
            return i + fraction
    return None


def compute_financials(project_data):
    """Master orchestrator. Returns full results dict."""
    wacc = project_data["wacc"]
    initial_investment = project_data["initial_investment"]
    yearly_data = project_data["yearly_data"]

    year_metrics = []
    cumulative_fcf = 0.0
    cumulative_discounted_fcf = 0.0

    for i, year_data in enumerate(yearly_data):
        year_num = i + 1
        metrics = compute_year_metrics(year_data, wacc, year_num)

        cumulative_fcf += metrics["fcf"]
        cumulative_discounted_fcf += metrics["discounted_fcf"]
        metrics["cumulative_fcf"] = cumulative_fcf
        metrics["cumulative_discounted_fcf"] = cumulative_discounted_fcf

        year_metrics.append(metrics)

    fcf_list = [m["fcf"] for m in year_metrics]
    discounted_fcf_list = [m["discounted_fcf"] for m in year_metrics]

    npv = compute_npv(initial_investment, fcf_list, wacc)
    irr = compute_irr(initial_investment, fcf_list)
    simple_payback = compute_simple_payback(initial_investment, fcf_list)
    discounted_payback = compute_discounted_payback(initial_investment, discounted_fcf_list)

    # Average margins (exclude years with zero revenue from margin averages)
    gross_margins = [m["gross_margin_pct"] for m in year_metrics if m["revenue"] != 0]
    ebitda_margins = [m["ebitda_margin_pct"] for m in year_metrics if m["revenue"] != 0]
    net_margins = [m["net_margin_pct"] for m in year_metrics if m["revenue"] != 0]

    avg_gross_margin = sum(gross_margins) / len(gross_margins) if gross_margins else 0.0
    avg_ebitda_margin = sum(ebitda_margins) / len(ebitda_margins) if ebitda_margins else 0.0
    avg_net_margin = sum(net_margins) / len(net_margins) if net_margins else 0.0

    return {
        "year_metrics": year_metrics,
        "npv": npv,
        "irr": irr,
        "simple_payback": simple_payback,
        "discounted_payback": discounted_payback,
        "avg_gross_margin": avg_gross_margin,
        "avg_ebitda_margin": avg_ebitda_margin,
        "avg_net_margin": avg_net_margin,
    }
