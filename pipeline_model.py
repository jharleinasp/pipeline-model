import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta

# Page configuration
st.set_page_config(page_title="Financial Pipeline Modelling Tool", layout="wide")

# Initialize session state
if 'probabilities' not in st.session_state:
    st.session_state.probabilities = {
        'Secured income': 100,
        'Proposals out for decision': 75,
        'High likelihood projects in development': 55,
        'Medium likelihood projects in development': 35,
        'Ideas at development stage': 15
    }

if 'scenario' not in st.session_state:
    st.session_state.scenario = 'realistic'

# Header
st.title("Financial Pipeline Modelling Tool")
st.markdown("*18-month scenario planning with staff cost recovery and reserve management*")

# Scenario presets
scenario_presets = {
    'conservative': {
        'Secured income': 100,
        'Proposals out for decision': 75,
        'High likelihood projects in development': 40,
        'Medium likelihood projects in development': 0,
        'Ideas at development stage': 0
    },
    'realistic': {
        'Secured income': 100,
        'Proposals out for decision': 75,
        'High likelihood projects in development': 55,
        'Medium likelihood projects in development': 30,
        'Ideas at development stage': 0
    },
    'optimistic': {
        'Secured income': 100,
        'Proposals out for decision': 90,
        'High likelihood projects in development': 70,
        'Medium likelihood projects in development': 50,
        'Ideas at development stage': 20
    }
}

def parse_excel_pipeline(excel_file):
    """Parse multi-sheet Excel file with opportunities"""
    all_opportunities = []
    
    # Read all sheets
    xl_file = pd.ExcelFile(excel_file)
    
    for sheet_name in xl_file.sheet_names:
        # Read the sheet without any date parsing
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, dtype=str)
        
        # Extract opportunity name (A1) and cluster (A2)
        opportunity_name = df.iloc[0, 0] if len(df) > 0 else f"Opportunity_{sheet_name}"
        cluster = df.iloc[1, 0] if len(df) > 1 else "Unknown"
        
        # Month headers are in row 3 (index 2), starting from column B (index 1)
        months = df.iloc[2, 1:].tolist()
        
        # Income is in row 4 (index 3)
        income_values = df.iloc[3, 1:].tolist()
        
        # Staff is in row 5 (index 4)
        staff_values = df.iloc[4, 1:].tolist()
        
        # Expenses is in row 6 (index 5)
        expenses_values = df.iloc[5, 1:].tolist()
        
        # Create opportunity dictionary
        opp_data = {
            'opportunity_name': opportunity_name,
            'cluster': cluster
        }
        
        # Add monthly data
        for i, month in enumerate(months):
            if pd.notna(month) and str(month).strip() != '' and str(month).strip().lower() != 'nan':
                # Remove leading apostrophe if present (Excel text formatting)
                month_str = str(month).strip().lstrip("'")
                
                # Convert string values to float, handling NaN and empty strings
                income = 0
                staff = 0
                expenses = 0
                
                if i < len(income_values) and pd.notna(income_values[i]) and str(income_values[i]).strip() != '':
                    try:
                        income = float(income_values[i])
                    except ValueError:
                        income = 0
                
                if i < len(staff_values) and pd.notna(staff_values[i]) and str(staff_values[i]).strip() != '':
                    try:
                        staff = float(staff_values[i])
                    except ValueError:
                        staff = 0
                
                if i < len(expenses_values) and pd.notna(expenses_values[i]) and str(expenses_values[i]).strip() != '':
                    try:
                        expenses = float(expenses_values[i])
                    except ValueError:
                        expenses = 0
                
                opp_data[f"{month_str}_income"] = income
                opp_data[f"{month_str}_staff"] = staff
                opp_data[f"{month_str}_expenses"] = expenses
        
        all_opportunities.append(opp_data)
    
    return pd.DataFrame(all_opportunities)

def get_month_label(month_index):
    """Convert month index to label like Jan_2026"""
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # Start from January 2026 (month 0 of year 2026)
    year = 2026 + (month_index - 1) // 12
    month = (month_index - 1) % 12
    
    return f"{month_names[month]}_{year}"

def calculate_forecast(pipeline_data, probabilities, unrestricted, restricted, fixed_staff, fixed_backoffice):
    """Calculate 18-month financial forecast with staff cost recovery"""
    months = 18
    forecast = []
    
    # Month 0 (Current)
    forecast.append({
        'month': 0,
        'monthLabel': 'Current',
        'unrestrictedReserves': unrestricted,
        'restrictedReserves': restricted,
        'totalCash': unrestricted + restricted
    })
    
    # Months 1-18
    for month in range(1, months + 1):
        month_label = get_month_label(month)
        
        # Initialize monthly totals
        total_income = 0
        total_project_staff = 0
        total_project_expenses = 0
        
        # Calculate weighted values from pipeline
        for _, opp in pipeline_data.iterrows():
            cluster = opp.get('cluster', '')
            probability = probabilities.get(cluster, 0) / 100
            
            income_col = f"{month_label}_income"
            staff_col = f"{month_label}_staff"
            expenses_col = f"{month_label}_expenses"
            
            income = float(opp.get(income_col, 0)) if pd.notna(opp.get(income_col, 0)) else 0
            staff = float(opp.get(staff_col, 0)) if pd.notna(opp.get(staff_col, 0)) else 0
            expenses = float(opp.get(expenses_col, 0)) if pd.notna(opp.get(expenses_col, 0)) else 0
            
            total_income += income * probability
            total_project_staff += staff * probability
            total_project_expenses += expenses * probability
        
        # Calculate contribution to overheads (from each project)
        project_contribution = total_income - total_project_staff - total_project_expenses
        
        # Calculate staff cost recovery
        staff_recovery = total_project_staff
        unrecovered_staff_costs = fixed_staff - staff_recovery
        
        # Calculate what needs to be covered from contribution
        costs_to_cover = unrecovered_staff_costs + fixed_backoffice
        
        # Calculate net position (contribution vs costs to cover)
        net_position = project_contribution - costs_to_cover
        
        # Get previous month's reserves
        prev_unrestricted = forecast[-1]['unrestrictedReserves']
        prev_restricted = forecast[-1]['restrictedReserves']
        
        # Apply reserve allocation rules
        if net_position >= 0:
            # Surplus: add to restricted reserves
            new_unrestricted = prev_unrestricted
            new_restricted = prev_restricted + net_position
        else:
            # Deficit: deduct from unrestricted reserves
            deficit = abs(net_position)
            new_unrestricted = prev_unrestricted - deficit
            new_restricted = prev_restricted
        
        forecast.append({
            'month': month,
            'monthLabel': month_label,
            'totalIncome': total_income,
            'projectStaffCosts': total_project_staff,
            'projectExpenses': total_project_expenses,
            'projectContribution': project_contribution,
            'fixedStaffCosts': fixed_staff,
            'staffRecovery': staff_recovery,
            'unrecoveredStaffCosts': unrecovered_staff_costs,
            'fixedBackOfficeCosts': fixed_backoffice,
            'costsFromContribution': costs_to_cover,
            'netPosition': net_position,
            'unrestrictedReserves': new_unrestricted,
            'restrictedReserves': new_restricted,
            'totalCash': new_unrestricted + new_restricted
        })
    
    return pd.DataFrame(forecast)

# Three-column layout
col1, col2, col3 = st.columns(3)

# Column 1: Current Financial Position
with col1:
    st.subheader("Current Financial Position")
    
    unrestricted_reserves = st.number_input(
        "Unrestricted Reserves (£)",
        value=143000,
        step=1000,
        format="%d"
    )
    
    restricted_reserves = st.number_input(
        "Restricted Reserves (£)",
        value=50000,
        step=1000,
        format="%d"
    )
    
    st.markdown(f"**Total Cash Reserves:** £{(unrestricted_reserves + restricted_reserves):,.0f}")
    
    st.markdown("---")
    st.markdown("**Fixed Monthly Costs**")
    
    fixed_staff_costs = st.number_input(
        "Fixed Staff Costs (£/month)",
        value=45000,
        step=1000,
        format="%d",
        help="Total monthly salary bill"
    )
    
    fixed_backoffice_costs = st.number_input(
        "Fixed Back Office Costs (£/month)",
        value=12000,
        step=1000,
        format="%d",
        help="Monthly overhead costs"
    )
    
    st.markdown(f"**Total Fixed Costs:** £{(fixed_staff_costs + fixed_backoffice_costs):,.0f}/month")
    
    st.markdown("---")
    
    threshold = st.number_input(
        "Critical Threshold (£)",
        value=143000,
        step=1000,
        format="%d",
        help="Minimum unrestricted reserves"
    )

# Column 2: Pipeline Data Upload
with col2:
    st.subheader("Pipeline Data Upload")
    
    uploaded_file = st.file_uploader("Upload Pipeline Excel File (.xlsx)", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            pipeline_data = parse_excel_pipeline(uploaded_file)
            st.success(f"✓ {len(pipeline_data)} opportunities loaded")
            
            # Show opportunity names
            with st.expander("View loaded opportunities"):
                for _, opp in pipeline_data.iterrows():
                    st.write(f"• **{opp['opportunity_name']}** ({opp['cluster']})")
            
            # Debug: Show what data was actually read
            with st.expander("🔍 Debug: View raw data structure"):
                st.write("**Column names found:**")
                cols_to_show = [col for col in pipeline_data.columns if '_income' in col or '_staff' in col or '_expenses' in col]
                st.write(cols_to_show[:10])  # Show first 10 columns
                
                st.write("**First opportunity sample data:**")
                if len(pipeline_data) > 0:
                    first_opp = pipeline_data.iloc[0]
                    sample_data = {}
                    for col in cols_to_show[:6]:  # Show first 6 months
                        sample_data[col] = first_opp.get(col, 'NOT FOUND')
                    st.json(sample_data)
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            st.error(f"Full error: {repr(e)}")
            import traceback
            st.code(traceback.format_exc())
            pipeline_data = pd.DataFrame()
    else:
        pipeline_data = pd.DataFrame()
        st.info("Upload an Excel file to begin modelling")
    
    st.markdown("---")
    st.markdown("**Quick Scenarios**")
    
    col2a, col2b, col2c = st.columns(3)
    
    with col2a:
        if st.button("Conservative", use_container_width=True):
            st.session_state.probabilities = scenario_presets['conservative']
            st.session_state.scenario = 'conservative'
            st.rerun()
    
    with col2b:
        if st.button("Realistic", use_container_width=True):
            st.session_state.probabilities = scenario_presets['realistic']
            st.session_state.scenario = 'realistic'
            st.rerun()
    
    with col2c:
        if st.button("Optimistic", use_container_width=True):
            st.session_state.probabilities = scenario_presets['optimistic']
            st.session_state.scenario = 'optimistic'
            st.rerun()

# Column 3: Probability Settings
with col3:
    st.subheader("Probability Settings (%)")
    
    for cluster in st.session_state.probabilities.keys():
        st.session_state.probabilities[cluster] = st.slider(
            cluster,
            min_value=0,
            max_value=100,
            value=st.session_state.probabilities[cluster],
            format="%d%%"
        )

# Only proceed if data is uploaded
if not pipeline_data.empty:
    # Calculate forecast
    forecast_df = calculate_forecast(
        pipeline_data,
        st.session_state.probabilities,
        unrestricted_reserves,
        restricted_reserves,
        fixed_staff_costs,
        fixed_backoffice_costs
    )
    
    # Calculate risk metrics
    min_unrestricted = forecast_df['unrestrictedReserves'].min()
    months_below_threshold = (forecast_df['unrestrictedReserves'] < threshold).sum()
    first_breach = forecast_df[forecast_df['unrestrictedReserves'] < threshold]['monthLabel'].iloc[0] if months_below_threshold > 0 else None
    min_total_cash = forecast_df['totalCash'].min()
    max_total_cash = forecast_df['totalCash'].max()
    is_at_risk = min_unrestricted < threshold
    
    # Calculate average staff recovery rate
    avg_staff_recovery = forecast_df[forecast_df['month'] > 0]['staffRecovery'].mean()
    avg_staff_recovery_pct = (avg_staff_recovery / fixed_staff_costs * 100) if fixed_staff_costs > 0 else 0
    
    # Risk Metrics Dashboard
    st.markdown("---")
    metric_col1, metric_col2, metric_col3, metric_col4, metric_col5, metric_col6 = st.columns(6)
    
    with metric_col1:
        if is_at_risk:
            st.metric("Risk Status", "At Risk ⚠️")
        else:
            st.metric("Risk Status", "Healthy ✓")
    
    with metric_col2:
        st.metric(
            "Min. Unrestricted",
            f"£{min_unrestricted:,.0f}",
            delta="Below threshold" if min_unrestricted < threshold else "Above threshold"
        )
    
    with metric_col3:
        st.metric(
            "Avg Staff Recovery",
            f"{avg_staff_recovery_pct:.0f}%",
            delta=f"£{avg_staff_recovery:,.0f}/month"
        )
    
    with metric_col4:
        st.metric(
            "Total Cash Range",
            f"£{min_total_cash:,.0f}",
            delta=f"to £{max_total_cash:,.0f}"
        )
    
    with metric_col5:
        st.metric(
            "Months Below",
            f"{months_below_threshold}",
            delta="of 18 months"
        )
    
    with metric_col6:
        st.metric(
            "First Breach",
            first_breach if first_breach else "None"
        )
    
    # Reserve Levels Forecast Chart
    st.markdown("---")
    st.subheader("Reserve Levels Forecast (18 Months)")
    
    fig = go.Figure()
    
    # Add unrestricted reserves line
    fig.add_trace(go.Scatter(
        x=forecast_df['monthLabel'],
        y=forecast_df['unrestrictedReserves'],
        mode='lines+markers',
        name='Unrestricted Reserves',
        line=dict(color='#2563eb', width=3),
        marker=dict(size=6)
    ))
    
    # Add total cash line
    fig.add_trace(go.Scatter(
        x=forecast_df['monthLabel'],
        y=forecast_df['totalCash'],
        mode='lines+markers',
        name='Total Cash',
        line=dict(color='#10b981', width=2, dash='dash'),
        marker=dict(size=4)
    ))
    
    # Add threshold line
    fig.add_hline(
        y=threshold,
        line_dash="dash",
        line_color="red",
        annotation_text="Critical Threshold",
        annotation_position="right"
    )
    
    fig.update_layout(
        height=400,
        xaxis_title="Month",
        yaxis_title="Amount (£)",
        hovermode='x unified',
        yaxis=dict(tickformat='£,.0f')
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Staff Cost Recovery Chart
    st.markdown("---")
    st.subheader("Staff Cost Recovery Analysis")
    
    fig2 = go.Figure()
    
    # Filter out month 0
    recovery_df = forecast_df[forecast_df['month'] > 0]
    
    fig2.add_trace(go.Bar(
        x=recovery_df['monthLabel'],
        y=recovery_df['staffRecovery'],
        name='Recovered from Projects',
        marker_color='#10b981'
    ))
    
    fig2.add_trace(go.Bar(
        x=recovery_df['monthLabel'],
        y=recovery_df['unrecoveredStaffCosts'],
        name='Unrecovered Staff Costs',
        marker_color='#ef4444'
    ))
    
    fig2.add_hline(
        y=fixed_staff_costs,
        line_dash="dash",
        line_color="gray",
        annotation_text=f"Total Staff Costs (£{fixed_staff_costs:,.0f})",
        annotation_position="right"
    )
    
    fig2.update_layout(
        barmode='stack',
        height=350,
        xaxis_title="Month",
        yaxis_title="Amount (£)",
        hovermode='x unified',
        yaxis=dict(tickformat='£,.0f')
    )
    
    st.plotly_chart(fig2, use_container_width=True)
    
    # Monthly Breakdown Table
    st.markdown("---")
    st.subheader("Detailed Monthly Breakdown")
    
    # Prepare display dataframe
    display_df = forecast_df[forecast_df['month'] > 0].copy()
    
    # Select and rename columns
    display_df = display_df[[
        'monthLabel', 'totalIncome', 'projectStaffCosts', 'projectExpenses', 
        'projectContribution', 'staffRecovery', 'unrecoveredStaffCosts',
        'fixedBackOfficeCosts', 'costsFromContribution', 'netPosition',
        'unrestrictedReserves', 'restrictedReserves', 'totalCash'
    ]].copy()
    
    display_df.columns = [
        'Month', 'Income', 'Project Staff', 'Project Expenses', 
        'Contribution', 'Staff Recovery', 'Unrecovered Staff',
        'Back Office', 'Costs from Contrib.', 'Net Position',
        'Unrestricted', 'Restricted', 'Total Cash'
    ]
    
    # Format currency columns
    currency_cols = ['Income', 'Project Staff', 'Project Expenses', 'Contribution',
                     'Staff Recovery', 'Unrecovered Staff', 'Back Office', 
                     'Costs from Contrib.', 'Net Position', 'Unrestricted', 
                     'Restricted', 'Total Cash']
    
    for col in currency_cols:
        display_df[col] = display_df[col].apply(lambda x: f"£{x:,.0f}")
    
    # Display table
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        height=400
    )

# Information Box
st.markdown("---")
st.info("""
**Excel File Format:**

Create a multi-sheet Excel (.xlsx) file where each sheet represents one opportunity.

**Each sheet structure:**
- **Cell A1:** Opportunity name (e.g., "Project Alpha")
- **Cell A2:** Cluster name (e.g., "Secured income")
- **Row 3, starting Column B:** Month headers (Jan_2026, Feb_2026, Mar_2026, ... Jun_2027)
- **Row 4, starting Column B:** Income values for each month
- **Row 5, starting Column B:** Staff cost values for each month
- **Row 6, starting Column B:** Expense values for each month

**Example Sheet:**
```
A                                      B         C         D
1  Project Alpha
2  Secured income
3                                      Jan_2026  Feb_2026  Mar_2026
4  Income                              50000     50000     50000
5  Staff                               30000     30000     30000
6  Expenses                            5000      5000      5000
```

**Valid Clusters:** 
- Secured income
- Proposals out for decision
- High likelihood projects in development
- Medium likelihood projects in development
- Ideas at development stage

**How the Model Works:**
1. Income comes in from pipeline opportunities
2. Project staff costs directly recover/offset the fixed staff costs (£45,000/month)
3. Project expenses are deducted from income
4. Contribution = Income - Staff - Expenses
5. Unrecovered Staff + Back Office costs are covered from Contribution
6. **Surplus** → Added to restricted reserves
7. **Deficit** → Deducted from unrestricted reserves
""")
