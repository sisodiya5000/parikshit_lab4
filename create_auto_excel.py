#!/usr/bin/env python3
"""
Auto-Updating Excel Sheet Generator
Creates an Excel file with interconnected formulas that automatically update when values change.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_auto_update_excel():
    """Create an Excel workbook with auto-updating formulas"""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Auto-Update Dashboard"
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    input_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    formula_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    total_fill = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
    
    # Border style
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # ============ SECTION 1: INPUT VALUES ============
    ws['A1'] = "📊 AUTO-UPDATE EXCEL DASHBOARD"
    ws['A1'].font = Font(bold=True, size=16)
    ws.merge_cells('A1:F1')
    
    # Input section headers
    ws['A3'] = "🔧 INPUT VALUES (Change these to update everything)"
    ws['A3'].font = Font(bold=True, size=12)
    ws.merge_cells('A3:F3')
    
    # Base input values
    inputs = [
        ("Base Price per Unit ($)", "B5", 50),
        ("Quantity", "B6", 100),
        ("Tax Rate (%)", "B7", 8.5),
        ("Discount Rate (%)", "B8", 15),
        ("Exchange Rate (USD to EUR)", "B9", 0.85)
    ]
    
    for i, (label, cell, value) in enumerate(inputs, 5):
        ws[f'A{i}'] = label
        ws[cell] = value
        ws[cell].fill = input_fill
        ws[cell].border = thin_border
        ws[f'A{i}'].border = thin_border
    
    # ============ SECTION 2: CALCULATED VALUES ============
    ws['A11'] = "⚡ AUTO-CALCULATED VALUES (Updated automatically)"
    ws['A11'].font = Font(bold=True, size=12)
    ws.merge_cells('A11:F11')
    
    calculations = [
        ("Subtotal (Before Discount)", "B13", "=B5*B6"),
        ("Discount Amount", "B14", "=B13*(B8/100)"),
        ("Subtotal (After Discount)", "B15", "=B13-B14"),
        ("Tax Amount", "B16", "=B15*(B7/100)"),
        ("Total (USD)", "B17", "=B15+B16"),
        ("Total (EUR)", "B18", "=B17*B9"),
        ("Price per Unit (After All)", "B19", "=B17/B6"),
        ("Savings from Discount", "B20", "=B14")
    ]
    
    for i, (label, cell, formula) in enumerate(calculations, 13):
        ws[f'A{i}'] = label
        ws[cell] = formula
        ws[cell].fill = formula_fill
        ws[cell].border = thin_border
        ws[f'A{i}'].border = thin_border
    
    # ============ SECTION 3: CONDITIONAL FORMATTING & ANALYSIS ============
    ws['A22'] = "📈 ANALYSIS & INSIGHTS"
    ws['A22'].font = Font(bold=True, size=12)
    ws.merge_cells('A22:F22')
    
    analysis = [
        ("Profit Margin (%)", "B24", "=(B17-B5*B6*0.6)/B17*100"),
        ("Volume Category", "B25", '=IF(B6<50,"Low",IF(B6<200,"Medium","High"))'),
        ("Price Category", "B26", '=IF(B5<30,"Budget",IF(B5<100,"Standard","Premium"))'),
        ("Total Margin ($)", "B27", "=B17-B5*B6*0.6"),
        ("Break-even Quantity", "B28", "=100/(B5-30)")
    ]
    
    for i, (label, cell, formula) in enumerate(analysis, 24):
        ws[f'A{i}'] = label
        ws[cell] = formula
        ws[cell].fill = total_fill
        ws[cell].border = thin_border
        ws[f'A{i}'].border = thin_border
    
    # ============ SECTION 4: DETAILED BREAKDOWN TABLE ============
    ws['D4'] = "📋 DETAILED BREAKDOWN"
    ws['D4'].font = Font(bold=True, size=12)
    ws.merge_cells('D4:G4')
    
    # Create a breakdown table
    headers = ["Item", "Quantity", "Unit Price", "Total"]
    for i, header in enumerate(headers, 1):
        cell = ws.cell(row=5, column=3+i)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')
    
    # Breakdown data with formulas
    breakdown_data = [
        ("Product A", "=B6*0.6", "=B5", "=E6*F6"),
        ("Product B", "=B6*0.25", "=B5*1.2", "=E7*F7"),
        ("Product C", "=B6*0.15", "=B5*0.8", "=E8*F8"),
        ("TOTAL", "=SUM(E6:E8)", "", "=SUM(G6:G8)")
    ]
    
    for i, (item, qty, price, total) in enumerate(breakdown_data, 6):
        ws[f'D{i}'] = item
        if qty.startswith('='):
            ws[f'E{i}'] = qty
        else:
            ws[f'E{i}'] = qty
        
        if price.startswith('='):
            ws[f'F{i}'] = price
        else:
            ws[f'F{i}'] = price
            
        ws[f'G{i}'] = total
        
        # Apply formatting
        for col in ['D', 'E', 'F', 'G']:
            ws[f'{col}{i}'].border = thin_border
            if i == 9:  # Total row
                ws[f'{col}{i}'].fill = total_fill
                ws[f'{col}{i}'].font = Font(bold=True)
    
    # ============ SECTION 5: CHARTS DATA ============
    ws['D12'] = "📊 CHART DATA"
    ws['D12'].font = Font(bold=True, size=12)
    ws.merge_cells('D12:E12')
    
    chart_headers = ["Scenario", "Total Value"]
    for i, header in enumerate(chart_headers, 1):
        cell = ws.cell(row=14, column=3+i)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    scenarios = [
        ("Current", "=B17"),
        ("No Discount", "=B13+B13*(B7/100)"),
        ("Double Quantity", "=(B5*B6*2)*(1-B8/100)*(1+B7/100)"),
        ("Half Price", "=(B5/2*B6)*(1-B8/100)*(1+B7/100)")
    ]
    
    for i, (scenario, formula) in enumerate(scenarios, 15):
        ws[f'D{i}'] = scenario
        ws[f'E{i}'] = formula
        ws[f'D{i}'].border = thin_border
        ws[f'E{i}'].border = thin_border
        ws[f'E{i}'].fill = formula_fill
    
    # ============ FORMATTING ============
    # Set column widths
    column_widths = {
        'A': 25, 'B': 15, 'C': 3, 'D': 20, 'E': 12, 'F': 12, 'G': 12
    }
    
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    # Add instructions
    ws['A30'] = "📝 INSTRUCTIONS:"
    ws['A30'].font = Font(bold=True, size=12)
    
    instructions = [
        "1. Change any GREEN input values (B5:B9) to see automatic updates",
        "2. All YELLOW cells contain formulas that auto-calculate",
        "3. ORANGE cells show totals and analysis",
        "4. The breakdown table and scenarios update automatically",
        "5. Try changing Base Price, Quantity, Tax Rate, or Discount Rate"
    ]
    
    for i, instruction in enumerate(instructions, 31):
        ws[f'A{i}'] = instruction
        ws[f'A{i}'].font = Font(size=10)
    
    return wb

def create_advanced_excel():
    """Create a more advanced Excel with multiple sheets and complex formulas"""
    
    wb = Workbook()
    
    # Remove default sheet and create custom sheets
    wb.remove(wb.active)
    
    # ============ SHEET 1: MAIN DASHBOARD ============
    dashboard = wb.create_sheet("Dashboard")
    
    # Main KPI section
    dashboard['A1'] = "🎯 BUSINESS DASHBOARD"
    dashboard['A1'].font = Font(bold=True, size=16)
    dashboard.merge_cells('A1:E1')
    
    # Key inputs that drive everything
    dashboard['A3'] = "Key Business Inputs"
    dashboard['A3'].font = Font(bold=True)
    
    main_inputs = [
        ("Monthly Revenue Target", "B4", 50000),
        ("Cost of Goods Sold %", "B5", 35),
        ("Operating Expenses", "B6", 15000),
        ("Growth Rate %", "B7", 5),
        ("Team Size", "B8", 10)
    ]
    
    input_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    for label, cell, value in main_inputs:
        row = int(cell[1:])
        dashboard[f'A{row}'] = label
        dashboard[cell] = value
        dashboard[cell].fill = input_fill
    
    # Auto-calculated KPIs
    dashboard['A10'] = "📊 Auto-Calculated KPIs"
    dashboard['A10'].font = Font(bold=True)
    
    kpis = [
        ("Gross Profit", "B11", "=B4*(1-B5/100)"),
        ("Net Profit", "B12", "=B11-B6"),
        ("Profit Margin %", "B13", "=B12/B4*100"),
        ("Revenue per Employee", "B14", "=B4/B8"),
        ("Annual Revenue Projection", "B15", "=B4*12*(1+B7/100)"),
        ("Break-even Revenue", "B16", "=B6/(1-B5/100)")
    ]
    
    formula_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    for label, cell, formula in kpis:
        row = int(cell[1:])
        dashboard[f'A{row}'] = label
        dashboard[cell] = formula
        dashboard[cell].fill = formula_fill
    
    # ============ SHEET 2: FINANCIAL MODEL ============
    financial = wb.create_sheet("Financial_Model")
    
    # 12-month projection
    financial['A1'] = "💰 12-Month Financial Projection"
    financial['A1'].font = Font(bold=True, size=14)
    
    # Headers
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    financial['A3'] = "Month"
    for i, month in enumerate(months, 2):
        financial.cell(row=3, column=i).value = month
        financial.cell(row=3, column=i).font = Font(bold=True)
    
    # Financial rows with formulas
    financial_rows = [
        ("Revenue", 4, "=Dashboard.B4"),
        ("COGS", 5, "=B4*Dashboard.B5/100"),
        ("Gross Profit", 6, "=B4-B5"),
        ("Operating Exp", 7, "=Dashboard.B6"),
        ("Net Profit", 8, "=B6-B7"),
        ("Cumulative Profit", 9, "=B8")
    ]
    
    for label, row, formula in financial_rows:
        financial[f'A{row}'] = label
        financial[f'B{row}'] = formula
        
        # Copy formulas across months with growth
        for col in range(3, 14):  # C to M (months 2-12)
            col_letter = get_column_letter(col)
            prev_col = get_column_letter(col-1)
            
            if row == 4:  # Revenue with growth
                financial[f'{col_letter}{row}'] = f"={prev_col}{row}*(1+Dashboard.B7/100)"
            elif row == 5:  # COGS
                financial[f'{col_letter}{row}'] = f"={col_letter}4*Dashboard.B5/100"
            elif row == 6:  # Gross Profit
                financial[f'{col_letter}{row}'] = f"={col_letter}4-{col_letter}5"
            elif row == 7:  # Operating Expenses
                financial[f'{col_letter}{row}'] = "=Dashboard.B6"
            elif row == 8:  # Net Profit
                financial[f'{col_letter}{row}'] = f"={col_letter}6-{col_letter}7"
            elif row == 9:  # Cumulative
                financial[f'{col_letter}{row}'] = f"={prev_col}{row}+{col_letter}8"
    
    # ============ SHEET 3: SCENARIO ANALYSIS ============
    scenario = wb.create_sheet("Scenarios")
    
    scenario['A1'] = "🔮 Scenario Analysis"
    scenario['A1'].font = Font(bold=True, size=14)
    
    # Scenario inputs
    scenarios_data = [
        ("Scenario", "Revenue", "COGS %", "Op Exp", "Result"),
        ("Current", "=Dashboard.B4", "=Dashboard.B5", "=Dashboard.B6", "=B3*(1-C3/100)-D3"),
        ("Optimistic", "=Dashboard.B4*1.2", "=Dashboard.B5*0.9", "=Dashboard.B6*0.95", "=B4*(1-C4/100)-D4"),
        ("Pessimistic", "=Dashboard.B4*0.8", "=Dashboard.B5*1.1", "=Dashboard.B6*1.05", "=B5*(1-C5/100)-D5"),
        ("Best Case", "=Dashboard.B4*1.5", "=Dashboard.B5*0.8", "=Dashboard.B6*0.9", "=B6*(1-C6/100)-D6")
    ]
    
    for i, row_data in enumerate(scenarios_data, 3):
        for j, value in enumerate(row_data, 1):
            cell = scenario.cell(row=i, column=j)
            cell.value = value
            if i == 3:  # Header row
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            elif j > 1:  # Formula cells
                cell.fill = formula_fill
    
    # Save the workbook
    return wb

def main():
    """Main function to create Excel files"""
    
    print("Creating auto-updating Excel sheets...")
    
    # Create basic auto-update Excel
    print("1. Creating basic auto-update Excel...")
    wb1 = create_auto_update_excel()
    wb1.save("auto_update_excel.xlsx")
    print("   ✅ Saved: auto_update_excel.xlsx")
    
    # Create advanced multi-sheet Excel
    print("2. Creating advanced multi-sheet Excel...")
    wb2 = create_advanced_excel()
    wb2.save("advanced_auto_excel.xlsx")
    print("   ✅ Saved: advanced_auto_excel.xlsx")
    
    print("\n🎉 Excel files created successfully!")
    print("\n📋 What you can do:")
    print("   • Open either Excel file")
    print("   • Change any GREEN input values")
    print("   • Watch all formulas update automatically")
    print("   • Try different scenarios and see instant results")
    
    print("\n💡 Key Features:")
    print("   • All formulas are interconnected")
    print("   • Changing one value updates the entire sheet")
    print("   • Color-coded sections (inputs, formulas, totals)")
    print("   • Multiple sheets with cross-references")
    print("   • Real-time calculations and analysis")

if __name__ == "__main__":
    main()