# RT Reconciler

Compares Simple Tire workbook quantities against RT data and organizes results into tabs.

## Download

Go to **Actions** → Click latest build → Download **RT_Reconciler** artifact

## Usage

1. Double-click `RT_Reconciler.exe`
2. Browse to select the **Simple Workbook** file
3. Browse to select the **RT Comparison** file  
4. Click **Reconcile**
5. Output saves to the same folder as the Simple Workbook

## Output Tabs

| Tab | Description |
|-----|-------------|
| Summary | Quick stats - totals, matches, discrepancies |
| Discrepancies | Items that don't match (sorted by biggest variance) |
| Reconciled | Items that match perfectly (DIFF = 0) |
| Not in RT | Items in Simple that RT doesn't have |
| IE Tire Detail | Original transaction data preserved |
| Full Comparison | Complete merged view |

## Color Coding

- 🟢 **Green** = Match
- 🟡 **Yellow** = SIMPLE has more than RT  
- 🔴 **Red** = RT has more than SIMPLE

## Building Locally

```bash
pip install pandas openpyxl pyinstaller
pyinstaller --onefile --windowed --name "RT_Reconciler" rt_reconciler_app.py
```

EXE will be in the `dist` folder.
