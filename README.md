# CCS Curve Construction

A comprehensive Python application for constructing multi-currency interest rate curves and pricing Cross-Currency Swaps (CCS) for USD, THB (Thai Baht), and JPY (Japanese Yen) markets.

## Overview

This project provides a clean, modular framework for:

- **Multi-currency curve construction** (USD, THB, JPY)
- **Cross-currency swap (CCS) pricing**
- **FX forward curve construction**
- **NPV calculations** for booking positions
- **Sensitivity analysis** (FX delta and IR DV01)
- **Cash flow generation** and reporting

### Key Features

- **Multi-Currency Support**: Handles USD, THB, and JPY with proper curve bootstrapping
- **Synthetic Curve Construction**: Creates JPY-THB CCS curves using proper triangulation methodology
- **USD-CSA Consistent Pricing**: Implements USD-CSA (Credit Support Annex) consistent pricing framework
- **Comprehensive Booking System**: Processes booking positions with detailed NPV calculations
- **Excel Integration**: Reads market data and writes results to Excel workbooks
- **Sensitivity Analysis**: Calculates FX deltas and interest rate DV01 buckets
- **Cash Flow Reporting**: Generates detailed cash flow schedules for all positions

## Architecture

The project follows a modular design with the following main components:

### Core Classes

1. **`MarketData`**: Data class holding market data parameters (spot rates, fixings, curve dates)
2. **`BookingPosition`**: Data class representing a CCS booking position with all leg parameters
3. **`CalendarManager`**: Manages holiday calendars for different currencies and their combinations
4. **`MarketDataLoader`**: Loads and processes market data from Excel files
5. **`CurveBuilder`**: Constructs interest rate curves using bootstrapping methodology
6. **`CCSBookingCalculator`**: Handles CCS booking calculations and NPV computations
7. **`LHBCCSProcessor`**: Main orchestrator class that coordinates the entire process

### Key Methodologies

#### JPY-THB Triangulation

The system implements proper triangulation for JPY-THB CCS:

- Basis spread: `b_JPY-THB(t) = b_JPY-USD(t) - b_THB-USD(t)`
- FX forward: `F_JPYTHB(t) = F_USDTHB(t) / F_USDJPY(t)`
- Uses monotone interpolation and aligned tenor grids

#### Curve Construction

- **SOFR Curve**: USD overnight rate curve
- **THOR Curve**: Thai Baht overnight rate curve
- **TONR Curve**: Japanese Yen overnight rate curve
- **CCS Curves**: Cross-currency swap curves for USD/THB, USD/JPY, and synthetic THB/JPY

## Installation

### Prerequisites

- Python 3.8 or higher
- pip package manager

### Dependencies

Install required packages using:

```bash
pip install -r requirements.txt
```

### Required Python Packages

- `pandas` - Data manipulation and analysis
- `numpy` - Numerical computing
- `rateslib` - Interest rate curve construction and pricing
- `openpyxl` - Excel file reading and writing

## Usage

### Basic Usage

1. **Prepare your Excel template** with the following sheets:

   - `Market Data`: Market rates and spot FX rates
   - `Fixing`: Historical fixings for SOFR, THOR, and TONR
   - `Holidays`: Holiday calendar data
   - `Booking`: Booking positions to price
2. **Run the main script**:

```python
from CCS_Curve_Construction import LHBCCSProcessor

# Initialize processor with Excel file
processor = LHBCCSProcessor('your_template.xlsx')

# Process booking positions
booking_results, cashflow_data = processor.process_booking()

# Update Excel with results
processor.update_booking_npv_to_excel(booking_results)
processor.add_curves_to_excel()
processor.populate_cash_flow_settled_sheet(cashflow_data)

# Calculate and add sensitivities
sensitivities = processor.calculate_all_sensitivities()
processor.add_sensitivities_to_excel(sensitivities)
```

### Command Line Usage

Run with default Excel file:

```bash
python CCS_Curve_Construction.py
```

Or specify a custom Excel file:

```bash
python CCS_Curve_Construction.py your_template.xlsx
```

The script will:

1. Load market data from the Excel template
2. Construct all interest rate curves
3. Process booking positions and calculate NPVs
4. Update the Excel file with results
5. Generate cash flow schedules
6. Calculate and report sensitivities

### Excel Template Structure

The Excel template should contain the following sheets:

#### Market Data Sheet

- Curve date
- Spot FX rates (USD/THB, USD/JPY)
- SOFR, THOR, TONR rates with tenors
- FX swap rates for USD/THB and USD/JPY
- CCS rates for USD/THB and USD/JPY

#### Fixing Sheet

- Historical fixings for SOFR, TONR, and THOR
- Organized in side-by-side sections

#### Holidays Sheet

- Holiday dates with CDR codes (US, TH, EN, JP)

#### Booking Sheet

- Booking positions with:
  - Trade details (ID, dates, position type)
  - Leg 1 and Leg 2 parameters (currency, notional, rates, spreads)
  - Day count conventions, frequencies, calendars
  - FX fixing and payment lag
  - P&L currency

## Project Structure

```
CCS Pricing/
├── CCS_Curve_Construction.py  # Main application file
├── CCS_Template.xlsx          # Excel template (example)
├── README.md                  # This file
├── requirements.txt           # Python dependencies
└── .gitignore                 # Git ignore rules
```

## Key Features Explained

### Multi-Currency Curve Construction

The system builds curves for three currencies:

- **USD**: Uses SOFR (Secured Overnight Financing Rate)
- **THB**: Uses THOR (Thai Overnight Repurchase Rate)
- **JPY**: Uses TONR (Tokyo Overnight Average Rate)

### Cross-Currency Swap Pricing

Supports pricing of:

- USD/THB CCS
- USD/JPY CCS
- THB/JPY CCS (synthetic, via triangulation)

### NPV Calculation

For each booking position, the system calculates:

- **Net NPV**: Total present value of the swap
- **Leg 1 NPV**: Present value of the first leg
- **Leg 2 NPV**: Present value of the second leg
- **Break-even Spread**: Spread that makes NPV zero

All NPVs can be reported in USD, THB, or JPY as specified.

### Sensitivity Analysis

The system calculates:

- **FX Delta**: Sensitivity to FX rate movements
- **IR DV01**: Interest rate sensitivity by tenor bucket

## Technical Details

### Curve Bootstrapping

Curves are bootstrapped using:

- Overnight rate fixings for short end
- Interest rate swaps (IRS) for intermediate tenors
- Cross-currency swaps (CCS) for longer tenors

### Interpolation

- Uses linear zero rate interpolation
- Monotone interpolation for synthetic curves
- Aligned tenor grids for triangulation

### Day Count Conventions

- USD: Act/360
- THB: Act/365f
- JPY: Act/365f

### Calendars

- **USD**: NYC (New York) - standalone calendar
- **JPY**: TYO (Tokyo) - standalone calendar
- **THB**: No standalone calendar - THB holidays are always combined with the other currency's calendar:
  - **USD/THB**: NYC calendar (US holidays) + THB holidays from Holidays sheet
  - **JPY/THB**: TYO calendar (JP holidays) + THB holidays from Holidays sheet
- **Combined calendars**: Created dynamically from currency pairs and loaded from the Holidays sheet in Excel

## Error Handling

The system includes comprehensive error handling:

- Validates market data completeness
- Checks for missing fixings
- Validates booking position parameters
- Provides informative error messages

## Output

The system generates:

1. **Updated Excel file** with:

   - NPV values in Booking sheet
   - Constructed curves in Curves sheet
   - Cash flows in CF sheet
   - Sensitivities in Sensitivities sheet
2. **Console output** with:

   - Processing status
   - Booking results summary
   - Cash flow summaries
   - Sensitivity summaries

## Limitations

- Currently supports USD, THB, and JPY only
- Maximum tenor: 10 years
- Requires Excel template in specific format

## Contributing

This is a proprietary project. For questions or issues, please contact the maintainer.

## Version

Current version: Clean Structure with Improved JPY-THB Triangulation

## Notes

- The system uses `rateslib` for curve construction and pricing
- Excel files are read/written using `openpyxl`
- All dates are handled as Python datetime objects
- The system supports both fixed and floating rate legs
