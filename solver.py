#!/usr/bin/env python3
"""
Sales & Ops Linear/Mixed-Integer Optimizer
- Reads a single Excel workbook with sheets: Base, Items, Limits (optional)
- Builds a linear program (PuLP) to maximize profit or minimize cost
- Optional binary "plan / not plan" decisions per row using big-M
- Writes a new workbook with a Solution sheet (never overwrites your template)
"""

import argparse
import pandas as pd
import numpy as np
import pulp
from pathlib import Path
from datetime import datetime

def read_input(xlsx_path: Path):
    # Required sheets
    base_df = pd.read_excel(xlsx_path, sheet_name="Base")
    items_df = pd.read_excel(xlsx_path, sheet_name="Items")
    # Optional
    try:
        limits_df = pd.read_excel(xlsx_path, sheet_name="Limits")
    except Exception:
        limits_df = pd.DataFrame(columns=["GroupType","GroupValue","MaxQty"])

    # Normalize column names
    items_df.columns = [c.strip() for c in items_df.columns]
    limits_df.columns = [c.strip() for c in limits_df.columns]

    # Fill optional numeric columns
    for col in ["MinQty","MaxQty","PPS","UnitCost"]:
        if col not in items_df.columns:
            items_df[col] = 0.0
    items_df["MinQty"] = items_df["MinQty"].fillna(0.0)
    items_df["MaxQty"] = items_df["MaxQty"].fillna(np.inf)
    items_df["PPS"] = items_df["PPS"].fillna(0.0)
    items_df["UnitCost"] = items_df["UnitCost"].fillna(0.0)

    # Pull config
    cfg = {}
    for _, row in base_df.iterrows():
        cfg[str(row.get("Key"))] = row.get("Value")

    # Defaults
    cfg.setdefault("Problem", "Max")            # Max / Min
    cfg.setdefault("UseBinaries", "No")         # Yes / No
    cfg.setdefault("Budget", np.inf)
    cfg.setdefault("TotalQtyLimit", np.inf)
    cfg.setdefault("OutputFile", "")

    # Clean
    if isinstance(cfg["Budget"], str) and cfg["Budget"] == "":
        cfg["Budget"] = np.inf
    if isinstance(cfg["TotalQtyLimit"], str) and cfg["TotalQtyLimit"] == "":
        cfg["TotalQtyLimit"] = np.inf

    return cfg, items_df, limits_df

def build_model(cfg, items_df, limits_df):
    problem_sense = pulp.LpMaximize if str(cfg["Problem"]).lower().startswith("max") else pulp.LpMinimize
    prob = pulp.LpProblem("SalesOpsPlan", problem_sense)

    n = len(items_df)
    idx = list(range(n))

    # Decision vars
    x = {i: pulp.LpVariable(f"x_{i}", lowBound=0) for i in idx}

    use_bin = str(cfg["UseBinaries"]).lower().startswith("y")
    y = None
    if use_bin:
        y = {i: pulp.LpVariable(f"y_{i}", lowBound=0, upBound=1, cat="Binary") for i in idx}

    # Bounds via constraints, support MinQty / MaxQty (big-M when binaries on)
    for i in idx:
        minq = float(items_df.loc[i, "MinQty"] or 0.0)
        maxq = float(items_df.loc[i, "MaxQty"] if np.isfinite(items_df.loc[i, "MaxQty"]) else 1e9)

        if use_bin:
            # enforce activation
            prob += x[i] >= minq * y[i], f"min_qty_{i}"
            prob += x[i] <= maxq * y[i], f"max_qty_{i}"
        else:
            # pure linear (continuous) with min & max
            if minq > 0:
                prob += x[i] >= minq, f"min_qty_{i}"
            if np.isfinite(maxq):
                prob += x[i] <= maxq, f"max_qty_{i}"

    # Global total quantity
    total_qty_limit = float(cfg.get("TotalQtyLimit", np.inf))
    if np.isfinite(total_qty_limit):
        prob += pulp.lpSum([x[i] for i in idx]) <= total_qty_limit, "total_qty_limit"

    # Budget
    budget = float(cfg.get("Budget", np.inf))
    if np.isfinite(budget):
        costs = [float(items_df.loc[i, "UnitCost"]) * x[i] for i in idx]
        prob += pulp.lpSum(costs) <= budget, "budget_limit"

    # Group constraints from Limits sheet
    # Expected columns: GroupType in {"Product","Channel","Customer","Global"}, GroupValue, MaxQty
    if not limits_df.empty:
        limits_df = limits_df.dropna(subset=["GroupType"]).copy()
        for _, r in limits_df.iterrows():
            gtype = str(r["GroupType"]).strip()
            gval = r.get("GroupValue")
            maxq = r.get("MaxQty", np.nan)
            if pd.isna(maxq):
                continue
            maxq = float(maxq)

            if gtype.lower() == "global":
                prob += pulp.lpSum([x[i] for i in idx]) <= maxq, f"limit_global_{_}"
            else:
                mask = items_df[gtype].astype(str) == str(gval)
                ids = [i for i in idx if mask.iloc[i]]
                if ids:
                    prob += pulp.lpSum([x[i] for i in ids]) <= maxq, f"limit_{gtype}_{gval}_{_}"

    # Objective
    # If maximizing, default to profit contribution PPS * x
    # If minimizing, default to total cost UnitCost * x
    if prob.sense == pulp.LpMaximize:
        objective = pulp.lpSum([float(items_df.loc[i, "PPS"]) * x[i] for i in idx])
    else:
        objective = pulp.lpSum([float(items_df.loc[i, "UnitCost"]) * x[i] for i in idx])
    prob += objective, "objective"

    return prob, x, y

def solve_and_report(prob, x, y, cfg, items_df, xlsx_in: Path, out_path: Path = None):
    status = prob.solve(pulp.PULP_CBC_CMD(msg=False))
    status_name = pulp.LpStatus[prob.status]

    qty = [pulp.value(x[i]) for i in x.keys()]
    plan = [int(round(pulp.value(y[i]))) if y is not None and pulp.value(y[i]) is not None else (1 if (qty[i] or 0) > 0 else 0) for i in x.keys()]

    items_df = items_df.copy()
    items_df["SolutionQty"] = qty
    items_df["Plan"] = plan

    items_df["Revenue"] = items_df["PPS"].fillna(0) * items_df["SolutionQty"].fillna(0)
    items_df["Cost"] = items_df["UnitCost"].fillna(0) * items_df["SolutionQty"].fillna(0)
    items_df["Profit"] = items_df["Revenue"] - items_df["Cost"]

    objective_value = pulp.value(prob.objective)

    # Prepare summary
    summary = pd.DataFrame({
        "Metric": ["SolverStatus","ObjectiveValue","TotalQty","TotalRevenue","TotalCost","TotalProfit"],
        "Value": [
            status_name,
            objective_value,
            items_df["SolutionQty"].sum(),
            items_df["Revenue"].sum(),
            items_df["Cost"].sum(),
            items_df["Profit"].sum(),
        ]
    })

    # Output path
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if out_path is None:
        out_name = f"{xlsx_in.stem}_solved_{ts}.xlsx"
        out_path = xlsx_in.parent / out_name

    # Write original sheets + Solution + Summary into a new workbook
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # copy original sheets
        try:
            base_df = pd.read_excel(xlsx_in, sheet_name="Base")
            base_df.to_excel(writer, sheet_name="Base", index=False)
        except Exception:
            pass
        try:
            items_orig = pd.read_excel(xlsx_in, sheet_name="Items")
            items_orig.to_excel(writer, sheet_name="Items", index=False)
        except Exception:
            pass
        try:
            limits_df = pd.read_excel(xlsx_in, sheet_name="Limits")
            limits_df.to_excel(writer, sheet_name="Limits", index=False)
        except Exception:
            pass

        # add solutions
        items_df.to_excel(writer, sheet_name="Solution", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    return out_path, items_df, summary, status_name, objective_value

def main():
    ap = argparse.ArgumentParser(description="Sales & Ops Linear/Mixed-Integer Optimizer")
    ap.add_argument("--input", "-i", required=True, help="Path to input Excel workbook (.xlsx) with sheets Base, Items, (optional) Limits")
    ap.add_argument("--output", "-o", default="", help="Optional explicit output .xlsx path")
    args = ap.parse_args()

    xlsx_in = Path(args.input)
    cfg, items_df, limits_df = read_input(xlsx_in)
    prob, x, y = build_model(cfg, items_df, limits_df)

    out_path = Path(args.output) if args.output else None
    out_path, sol_df, summary_df, status_name, obj_val = solve_and_report(prob, x, y, cfg, items_df, xlsx_in, out_path)

    print(f"Status: {status_name}")
    print(f"Objective value: {obj_val}")
    print(f"Solved workbook: {out_path}")

if __name__ == "__main__":
    main()
