How to use the calculator

### Job Information
Basic project details — Job No., Job Title, Designed by, Checked by, Date. These appear on the exported report but don't affect any calculations.

### Roof Angle
The α (alpha) value — your building's roof slope angle in degrees.

### Wind Load Information
- **Structure Category** and **Terrain Category** — pick from the dropdowns
- **Basic Wind Velocity (Vbasic)** — in m/s
- **Cdyn** — whether dynamic response is "Considered" or "Not Considered" (click to toggle)
- **Natural Frequency** (na-y, na-x) and **Damping Factor** (ζ-SLS, ζ-ULS)

### Aerodynamic Shape Factor
- **Cp,i** — two values, Positive and Negative
- **Ka, Kc, Kl, Kp** — the four pressure-factor inputs (Area Reduction, Combination, Local Pressure, Porous Cladding Reduction)

### Shielding Parameter
Choose one of three methods for **Ms**:
- **Conservative** (Ms = 1.0) — no extra inputs needed
- **Normal Suburban Housing** (Ms = 0.85) — no extra inputs needed
- **Calculate** — this reveals three more fields (average height, average breadth, and number of shielding buildings) and works out Ms from those

### Hill Shape Parameter
Choose **None**, **Hills and Ridges**, or **Escarpment**.
- Selecting **None** hides the extra fields and diagram — no topography effect is applied
- Selecting either of the other two reveals H, Lu, and x fields, plus the matching reference diagram

### Storey Data & Results
This is the main table — one row per floor/storey of your building.

- The first four columns (**Floor Height, Ly, Lx, Dead Load**) are ones you type into directly, for each storey
- Click **+ Add Row** to add another storey (up to 99 total)
- Click the **✕** at the end of any row to delete that storey
- Click **▶ Run All** to calculate every row — this fills in Reduced Level, Vdes, Wy, Fy, Vdes, Wx, and Fx for every storey based on everything you've entered above
- You can keep editing values and clicking Run All again as many times as you like — it recalculates fresh every time

### Generate Excel Report
Once you've clicked **Run All** at least once, this button builds a downloadable Excel file with your results, formatted like your original calculation report — including the correct hill-shape diagram if you selected one. Click it, and the file downloads straight to your computer.

### Notes
A read-only list of the standard notes and limitations that apply to this calculation method (carried over from your original spreadsheet).

