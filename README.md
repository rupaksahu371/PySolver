# PySolver
A dynamic optimization solver for sales and operations planning that reads product, channel, and customer criteria from Excel, applies configurable constraints (e.g., product limits, dimensions, profitability), and generates an optimized and profitable production/ordering plan into the Excel file.

# How it works 

* Input Excel must have these sheets:

  * Base (config):
    Problem = Max or Min
    UseBinaries = Yes (plan/not-plan per row) or No (pure linear)
    Budget = numeric or leave blank
    TotalQtyLimit = numeric or leave blank

  * Items (rows = candidate order lines):
    Columns: Product, Channel, Customer, MinQty, MaxQty, PPS, UnitCost

  * Limits (optional): group caps
    Columns: GroupType in {Product, Channel, Customer, Global}, GroupValue, MaxQty

* Objective

  * If Base.Problem=Max: maximize Σ PPS_i * x_i.
  * If Base.Problem=Min: minimize Σ UnitCost_i * x_i.

* Constraints (built automatically)

  * Per-row MinQty, MaxQty.
  * Budget: Σ UnitCost_i * x_i ≤ Budget (if set).
  * TotalQtyLimit: Σ x_i ≤ TotalQtyLimit (if set).
  * Limits sheet adds caps by Product/Channel/Customer or Global.
  * If UseBinaries=Yes: adds a binary y_i per row with big-M:
    * x_i ≥ MinQty_i * y_i, x_i ≤ MaxQty_i * y_i → solves “plan vs not plan”.

# Run in VS Code (Windows/macOS/Linux)

1. Create a folder, put the three files in it.
2. Open the folder in VS Code → open a terminal.
3. (Recommended) make a venv:
  ```cmd
  python -m venv .venv
  .venv\Scripts\activate    # Windows
  or
  source .venv/bin/activate # macOS/Linux
  ```
4. Install deps:
  ```cmd
  pip install -r requirements.txt
  ```
5. Copy solver_template.xlsx to e.g. my_plan.xlsx and edit the sheets to your data.
6. Run:
  ```cmd
  python solver.py --input my_plan.xlsx
  ```
  The script never overwrites your template. It writes a new file:
  ```cmd
  my_plan_solved_YYYYMMDD_HHMMSS.xlsx
  ```
  with sheets Base, Items, Limits, Solution, Summary.

# Extend to your criteria
You can add more constraints by just populating the Limits sheet:
* Cap a specific product: GroupType=Product, GroupValue=A, MaxQty=220
* Cap a channel: GroupType=Channel, GroupValue=Retail, MaxQty=500
* Per-customer demand cap: GroupType=Customer, GroupValue=C1, MaxQty=250
* Overall cap: GroupType=Global, MaxQty=900
For “which orders to plan vs not plan”:
* Set UseBinaries=Yes. Then each row gets a Plan (0/1) in Solution with SolutionQty. If you need a minimum activation (e.g., if you plan an order, produce at least 50), set MinQty=50 on that row—the model enforces it only when Plan=1.

# Swapping in your current workbook
If you’ve already built “Working Data” and “Base”:
* Map Working Data → Items with these columns (add if missing):
  Product, Channel, Customer, MinQty, MaxQty, PPS, UnitCost
* Map your config Base to the simple key/value pairs shown above.
