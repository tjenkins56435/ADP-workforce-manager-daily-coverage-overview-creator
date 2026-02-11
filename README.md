# ADP Daily Coverage Overview Creator

A desktop tool that imports ADP Workforce Manager weekly schedule reports and generates color-coded daily playbook spreadsheets — ready to print and post for your team.

## What It Does

Retail and shift-based teams using ADP Workforce Manager can export an "Employee Schedule - Weekly" report, but it's not ideal for day-of use. This tool takes that export and lets you:

1. **Import** the `.xlsx` weekly schedule report from ADP
2. **Select a day** from the week to build a daily view
3. **Assign zones** (Register, Floor, Fitting Room, etc.) to each employee with custom colors
4. **Generate** a formatted, print-ready Excel playbook with a color-coded timeline grid

The output spreadsheet includes employee names, shift times, break times, and a visual half-hour timeline bar colored by zone assignment — making it easy to see coverage at a glance.

## Screenshot

<!-- Add a screenshot of the app here if desired -->

## Requirements

- Python 3.8+
- tkinter (included with most Python installations)

## Installation

```bash
git clone https://github.com/tjenkins56435/ADP-workforce-manager-daily-coverage-overview-creator.git
cd ADP-workforce-manager-daily-coverage-overview-creator
pip install -r requirements.txt
```

## Usage

```bash
python dco_creator.py
```

### Workflow

1. In ADP Workforce Manager, export the **Employee Schedule - Weekly** report as `.xlsx`
2. Open the app and click **Open ADP Report (.xlsx)** to import the file
3. Select a day from the dropdown and click **Load Day**
4. Assign zones to employees by selecting a row and clicking **Set Zone** (or double-click a row)
   - Use **Set All Zones** to quickly assign every employee the same zone
5. Reorder employees with the **Up/Down** buttons as needed
6. Click **Generate Excel** to save the daily playbook to your chosen output folder

### Zone Configuration

Zones are fully customizable — add, edit, or delete zones with any name and color. Zone settings persist across sessions in `config.json`.

Default zones:
| Zone | Color |
|------|-------|
| Adults | Red |
| Kids/Footwear | Yellow |
| Cashiers | Green |
| Replenishment/Refill | Purple |
| Shipment | Blue |
| Operation | Orange |
| Fitting Rooms | Coral |

### Manual Entries

Need to add someone not in the ADP report? Click **Add Manual** to enter a name, shift, break, and zone by hand.

## Output

The generated Excel file includes:
- **Header** with the day and date
- **Employee rows** with name, shift time, and break time
- **Timeline grid** with half-hour columns color-coded by zone assignment
- Print-optimized landscape layout that fits to page width
