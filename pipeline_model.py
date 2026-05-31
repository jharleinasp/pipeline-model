import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta

# Password protection
def check_password():
    """Returns `True` if the user had the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.write("*Please enter the password to access the application.*")
        return False
    elif not st.session_state["password_correct"]:
        # Password incorrect, show input + error
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct
        return True

if not check_password():
    st.stop()

# Page configuration
st.set_page_config(page_title="Financial Pipeline Modelling Tool", layout="wide")

# Initialize session state
if 'probabilities' not in st.session_state:
    st.session_state.probabilities = {
        'Secured income': 100,
        'Contracting': 100,
        'Negotiating': 90,
        'Proposals out for decision': 65,
        'High likelihood projects in development': 50,
        'Medium likelihood projects in development': 30,
        'Ideas at development stage': 15
    }

if 'scenario' not in st.session_state:
    st.session_state.scenario = 'realistic'

if 'opportunity_toggles' not in st.session_state:
    st.session_state.opportunity_toggles = {}

# Header
st.title("Financial Pipeline Modelling Tool")
st.markdown("*18-month scenario planning with staff cost recovery and reserve management*")

# Scenario presets
scenario_presets = {
    'conservative': {
        'Secured income': 100,
        'Contracting': 100,
        'Negotiating': 90,
        'Proposals out for decision': 45,
        'High likelihood projects in development': 30,
        'Medium likelihood projects in development': 15,
        'Ideas at development stage': 5
    },
    'realistic': {
        'Secured income': 100,
        'Contracting': 100,
        'Negotiating': 90,
        'Proposals out for decision': 65,
        'High likelihood projects in development': 50,
        'Medium likelihood projects in development': 30,
        'Ideas at development stage': 15
    },
    'optimistic': {
        'Secured income': 100,
        'Contracting': 100,
        'Negotiating': 90,
        'Proposals out for decision': 85,
        'High likelihood projects in development': 70,
        'Medium likelihood projects in development': 45,
        'Ideas at development stage': 25
    }
}

# Available start months for the picker
_MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

START_MONTH_OPTIONS = [
    'Jan_2026', 'Feb_2026', 'Mar_2026', 'Apr_2026', 'May_2026', 'Jun_2026',
    'Jul_2026', 'Aug_2026', 'Sep_2026', 'Oct_2026', 'Nov_2026', 'Dec_2026',
    'Jan_2027', 'Feb_2027', 'Mar_2027', 'Apr_2027', 'May_2027', 'Jun_2027',
    'Jul_2027', 'Aug_2027', 'Sep_2027', 'Oct_2027', 'Nov_2027', 'Dec_2027'
]

def generate_month_list(start_month_str):
    """Generate 18-month list starting from the given month e.g. 'May_2026'"""
    month_name, year = start_month_str.split('_')
    year = int(year)
    start_idx = _MONTH_NAMES.index(month_name)
    months = []
    for i in range(18):
        m = (start_idx + i) % 12
        y = year + (start_idx + i) // 12
        months.append(f"{_MONTH_NAMES[m]}_{y}")
    return months

# Default month list — overridden by user selection at runtime
MONTH_LIST = generate_month_list('Jan_2026')

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

def calculate_pipeline_funnel(pipeline_data, probabilities, active_opportunities, months_filter):
    """Calculate pipeline funnel values for visualization"""
    
    # Determine which months to include based on filter
    if months_filter == 6:
        month_range = MONTH_LIST[:6]
    elif months_filter == 12:
        month_range = MONTH_LIST[:12]
    else:  # 18 months
        month_range = MONTH_LIST
    
    # Define funnel stages and their cluster mappings
    funnel_stages = {
        'All Opportunities': ['Ideas at development stage', 'Medium likelihood projects in development', 
                             'High likelihood projects in development', 'Proposals out for decision',
                             'Negotiating', 'Contracting'],
        'Identified Income': ['Medium likelihood projects in development', 'High likelihood projects in development',
                             'Proposals out for decision', 'Negotiating', 'Contracting'],
        'Proposals': ['Proposals out for decision', 'Negotiating', 'Contracting'],
        'Negotiating': ['Negotiating', 'Contracting'],
        'Contracting': ['Contracting']
    }
    
    funnel_data = []
    
    for stage_name, included_clusters in funnel_stages.items():
        total_value = 0
        weighted_value = 0
        
        for _, opp in pipeline_data.iterrows():
            opp_name = opp['opportunity_name']
            
            # Skip if opportunity is toggled off
            if opp_name not in active_opportunities or not active_opportunities[opp_name]:
                continue
            
            cluster = opp.get('cluster', '')
            
            # Skip if cluster not in this stage
            if cluster not in included_clusters:
                continue
            
            probability = probabilities.get(cluster, 0) / 100
            
            # Sum income across selected months
            for month_label in month_range:
                income_col = f"{month_label}_income"
                income = float(opp.get(income_col, 0)) if pd.notna(opp.get(income_col, 0)) else 0
                total_value += income
                weighted_value += income * probability
        
        funnel_data.append({
            'stage': stage_name,
            'total_value': total_value,
            'weighted_value': weighted_value
        })
    
    return pd.DataFrame(funnel_data)

def get_month_label(month_index):
    """Convert month index (1-18) to label using the current dynamic MONTH_LIST"""
    if 1 <= month_index <= len(MONTH_LIST):
        return MONTH_LIST[month_index - 1]
    return f"Month_{month_index}"

def get_month_index(month_label):
    """Convert month label like Jan_2026 to index (1-18)"""
    try:
        return MONTH_LIST.index(month_label) + 1
    except ValueError:
        return 0

def get_fixed_costs_for_month(month_label, cost_changes):
    """Get the applicable fixed costs for a given month based on cost changes"""
    month_idx = get_month_index(month_label)
    
    # Sort cost changes by month index
    sorted_changes = sorted(cost_changes, key=lambda x: get_month_index(x['month']))
    
    # Find the most recent cost change that applies to this month
    applicable_costs = {'staff': 45000, 'backoffice': 10500}  # defaults
    
    for change in sorted_changes:
        change_idx = get_month_index(change['month'])
        if change_idx > 0 and change_idx <= month_idx:
            applicable_costs['staff'] = change['staff']
            applicable_costs['backoffice'] = change['backoffice']
    
    return applicable_costs

def calculate_forecast(pipeline_data, probabilities, unrestricted_start, total_funds_start, 
                      base_staff, base_backoffice, reserve_deposits, cost_changes, active_opportunities,
                      special_projects_costs):
    """Calculate 18-month financial forecast with staff cost recovery"""
    months = 18
    forecast = []
    
    # Calculate static restricted funds
    restricted_funds = total_funds_start - unrestricted_start
    
    # Month 0 (Current)
    forecast.append({
        'month': 0,
        'monthLabel': 'Current',
        'unrestrictedReserves': unrestricted_start,
        'unrestrictedAfterSpecial': unrestricted_start,
        'restrictedFunds': restricted_funds,
        'totalFunds': total_funds_start
    })
    
    # Months 1-18
    for month in range(1, months + 1):
        month_label = get_month_label(month)
        
        # Get applicable fixed costs for this month
        fixed_costs = get_fixed_costs_for_month(month_label, cost_changes)
        fixed_staff = fixed_costs['staff']
        fixed_backoffice = fixed_costs['backoffice']
        
        # Get special projects cost for this month
        special_cost = 0
        for sp in special_projects_costs:
            if sp['month'] == month_label:
                special_cost = sp['amount']
                break
        
        # Initialize monthly totals
        total_income = 0
        total_project_staff = 0
        total_project_expenses = 0
        
        # Calculate weighted values from pipeline (only for active opportunities)
        for _, opp in pipeline_data.iterrows():
            opp_name = opp['opportunity_name']
            
            # Skip if opportunity is toggled off
            if opp_name not in active_opportunities or not active_opportunities[opp_name]:
                continue
            
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
        
        # Calculate contribution
        project_contribution = total_income - total_project_staff - total_project_expenses
        
        # Calculate staff cost recovery
        staff_recovery = total_project_staff
        unrecovered_staff_costs = max(0, fixed_staff - staff_recovery)
        
        # Use contribution to cover unrecovered staff costs first, then back office
        remaining_after_staff = project_contribution - unrecovered_staff_costs
        net_position = remaining_after_staff - fixed_backoffice
        costs_to_cover = unrecovered_staff_costs + fixed_backoffice
        
        # Get previous month's reserves
        prev_unrestricted = forecast[-1]['unrestrictedReserves']
        
        # Check for reserve deposits this month
        deposit_this_month = 0
        for deposit in reserve_deposits:
            if deposit['month'] == month_label and deposit['amount'] > 0:
                deposit_this_month += deposit['amount']
        
        # Apply simplified reserve rules
        new_unrestricted = prev_unrestricted + net_position + deposit_this_month
        
        # Calculate unrestricted after special projects
        new_unrestricted_after_special = new_unrestricted - special_cost
        
        # Total funds = unrestricted + static restricted funds
        new_total_funds = new_unrestricted + restricted_funds
        
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
            'reserveDeposit': deposit_this_month,
            'specialProjectsCost': special_cost,
            'unrestrictedReserves': new_unrestricted,
            'unrestrictedAfterSpecial': new_unrestricted_after_special,
            'restrictedFunds': restricted_funds,
            'totalFunds': new_total_funds
        })
    
    return pd.DataFrame(forecast)

# Model start month selector
st.markdown("---")
_start_col, _ = st.columns([1, 2])
with _start_col:
    _default_idx = START_MONTH_OPTIONS.index("May_2026") if "May_2026" in START_MONTH_OPTIONS else 0
    selected_start_month = st.selectbox(
        "📅 Model Start Month",
        options=START_MONTH_OPTIONS,
        index=_default_idx,
        help="Data before this month is ignored. The 18-month forecast runs from this month forward."
    )
# Rebuild MONTH_LIST from the selected start month
MONTH_LIST = generate_month_list(selected_start_month)

# Three-column layout
col1, col2, col3 = st.columns(3)

# Column 1: Current Financial Position
with col1:
    st.subheader("Current Financial Position")
    
    unrestricted_reserves = st.number_input(
        "Unrestricted Reserves (£)",
        value=100000,
        step=1000,
        format="%d"
    )
    
    total_funds = st.number_input(
        "Total Funds (£)",
        value=100000,
        step=1000,
        format="%d",
        help="Unrestricted reserves + Restricted funds held"
    )
    
    restricted_funds = total_funds - unrestricted_reserves
    st.markdown(f"**Restricted Funds Held:** £{restricted_funds:,.0f}")
    
    st.markdown("---")
    st.markdown("**Base Fixed Monthly Costs**")
    
    base_fixed_staff_costs = st.number_input(
        "Fixed Staff Costs (£/month)",
        value=45000,
        step=1000,
        format="%d",
        help="Base monthly salary bill"
    )
    
    base_fixed_backoffice_costs = st.number_input(
        "Fixed Back Office Costs (£/month)",
        value=10500,
        step=1000,
        format="%d",
        help="Base monthly overhead costs"
    )
    
    st.markdown(f"**Total Base Fixed Costs:** £{(base_fixed_staff_costs + base_fixed_backoffice_costs):,.0f}/month")
    
    # Cost changes
    with st.expander("💰 Fixed Cost Changes (up to 4)"):
        st.markdown("**Specify changes to fixed costs from specific months:**")
        cost_changes = []
        
        for i in range(4):
            col_month, col_staff, col_office = st.columns(3)
            
            with col_month:
                change_month = st.selectbox(
                    f"Month {i+1}",
                    options=['None'] + MONTH_LIST,
                    key=f"cost_month_{i}"
                )
            
            if change_month != 'None':
                with col_staff:
                    new_staff = st.number_input(
                        "Staff (£)",
                        value=base_fixed_staff_costs,
                        step=1000,
                        format="%d",
                        key=f"cost_staff_{i}"
                    )
                
                with col_office:
                    new_office = st.number_input(
                        "Back Office (£)",
                        value=base_fixed_backoffice_costs,
                        step=1000,
                        format="%d",
                        key=f"cost_office_{i}"
                    )
                
                cost_changes.append({
                    'month': change_month,
                    'staff': new_staff,
                    'backoffice': new_office
                })
    
    # Reserve deposits
    with st.expander("💵 Reserve Deposits (up to 4)"):
        st.markdown("**Add one-time deposits to unrestricted reserves:**")
        reserve_deposits = []
        
        for i in range(4):
            col_month, col_amount = st.columns(2)
            
            with col_month:
                deposit_month = st.selectbox(
                    f"Deposit {i+1} Month",
                    options=['None'] + MONTH_LIST,
                    key=f"deposit_month_{i}"
                )
            
            if deposit_month != 'None':
                with col_amount:
                    deposit_amount = st.number_input(
                        "Amount (£)",
                        value=0,
                        step=1000,
                        format="%d",
                        key=f"deposit_amount_{i}"
                    )
                
                reserve_deposits.append({
                    'month': deposit_month,
                    'amount': deposit_amount
                })
    
    # Special projects costs
    st.markdown("---")
    
    enable_special_projects = st.checkbox(
        "Enable Special Projects Costs",
        value=False,
        help="Add monthly costs for special projects (deducted from unrestricted reserves)"
    )
    
    special_projects_costs = []
    
    if enable_special_projects:
        with st.expander("🔧 Special Projects Costs (monthly)"):
            st.markdown("**Specify additional monthly costs for special projects:**")
            
            for month_label in MONTH_LIST:
                sp_cost = st.number_input(
                    f"{month_label}",
                    value=0,
                    step=1000,
                    format="%d",
                    key=f"special_{month_label}"
                )
                
                if sp_cost > 0:
                    special_projects_costs.append({
                        'month': month_label,
                        'amount': sp_cost
                    })
    
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
            
            # Initialize toggles for new opportunities
            for _, opp in pipeline_data.iterrows():
                opp_name = opp['opportunity_name']
                if opp_name not in st.session_state.opportunity_toggles:
                    st.session_state.opportunity_toggles[opp_name] = True
            
            # Opportunity toggles
            with st.expander("🎯 Toggle Opportunities"):
                st.markdown("**Select which opportunities to include in the model:**")
                for _, opp in pipeline_data.iterrows():
                    opp_name = opp['opportunity_name']
                    cluster = opp['cluster']
                    
                    st.session_state.opportunity_toggles[opp_name] = st.checkbox(
                        f"{opp_name} ({cluster})",
                        value=st.session_state.opportunity_toggles.get(opp_name, True),
                        key=f"toggle_{opp_name}"
                    )
            
            # Show opportunity names
            with st.expander("View loaded opportunities"):
                for _, opp in pipeline_data.iterrows():
                    status = "✓" if st.session_state.opportunity_toggles.get(opp['opportunity_name'], True) else "✗"
                    st.write(f"{status} **{opp['opportunity_name']}** ({opp['cluster']})")
            
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
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
        total_funds,
        base_fixed_staff_costs,
        base_fixed_backoffice_costs,
        reserve_deposits,
        cost_changes,
        st.session_state.opportunity_toggles,
        special_projects_costs
    )
    
    # Calculate risk metrics
    min_unrestricted = forecast_df['unrestrictedReserves'].min()
    months_below_threshold = (forecast_df['unrestrictedReserves'] < threshold).sum()
    first_breach = forecast_df[forecast_df['unrestrictedReserves'] < threshold]['monthLabel'].iloc[0] if months_below_threshold > 0 else None
    min_total_funds = forecast_df['totalFunds'].min()
    max_total_funds = forecast_df['totalFunds'].max()
    is_at_risk = min_unrestricted < threshold
    
    # Calculate average staff recovery rate (excluding month 0)
    avg_staff_recovery = forecast_df[forecast_df['month'] > 0]['staffRecovery'].mean()
    
    # Get average of fixed staff costs across all months
    avg_fixed_staff = forecast_df[forecast_df['month'] > 0]['fixedStaffCosts'].mean()
    avg_staff_recovery_pct = (avg_staff_recovery / avg_fixed_staff * 100) if avg_fixed_staff > 0 else 0
    
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
            "Total Funds Range",
            f"£{min_total_funds:,.0f}",
            delta=f"to £{max_total_funds:,.0f}"
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
    
    # Pipeline Funnel
    st.markdown("---")
    st.subheader("Pipeline Funnel Analysis")
    
    # Funnel filter
    funnel_col1, funnel_col2 = st.columns([1, 3])
    
    with funnel_col1:
        funnel_months = st.selectbox(
            "Time Period",
            options=[6, 12, 18],
            format_func=lambda x: f"Next {x} months",
            index=2  # Default to 18 months
        )
    
    with funnel_col2:
        st.markdown("*Shows total pipeline value and probability-weighted value at each stage (excludes Secured income)*")
    
    # Calculate funnel data
    funnel_df = calculate_pipeline_funnel(
        pipeline_data,
        st.session_state.probabilities,
        st.session_state.opportunity_toggles,
        funnel_months
    )
    
    # Create funnel visualization
    fig_funnel = go.Figure()
    
    # Add bars for total value
    fig_funnel.add_trace(go.Bar(
        y=funnel_df['stage'],
        x=funnel_df['total_value'],
        name='Total Value',
        orientation='h',
        marker=dict(color='#93c5fd'),
        text=funnel_df['total_value'].apply(lambda x: f'£{x:,.0f}'),
        textposition='inside',
        textfont=dict(size=12)
    ))
    
    # Add bars for weighted value
    fig_funnel.add_trace(go.Bar(
        y=funnel_df['stage'],
        x=funnel_df['weighted_value'],
        name='Probability-Weighted Value',
        orientation='h',
        marker=dict(color='#2563eb'),
        text=funnel_df['weighted_value'].apply(lambda x: f'£{x:,.0f}'),
        textposition='inside',
        textfont=dict(size=12, color='white')
    ))
    
    fig_funnel.update_layout(
        barmode='overlay',
        height=350,
        xaxis_title="Value (£)",
        yaxis_title="Pipeline Stage",
        xaxis=dict(tickformat='£,.0f'),
        yaxis=dict(autorange='reversed'),  # Top to bottom
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    st.plotly_chart(fig_funnel, use_container_width=True)
    
    # Funnel summary table
    st.markdown("**Pipeline Funnel Summary**")
    funnel_display = funnel_df.copy()
    funnel_display['Conversion Rate'] = (funnel_display['weighted_value'] / funnel_display['total_value'] * 100).apply(lambda x: f"{x:.1f}%")
    funnel_display['Total Value'] = funnel_display['total_value'].apply(lambda x: f'£{x:,.0f}')
    funnel_display['Weighted Value'] = funnel_display['weighted_value'].apply(lambda x: f'£{x:,.0f}')
    funnel_display = funnel_display[['stage', 'Total Value', 'Weighted Value', 'Conversion Rate']]
    funnel_display.columns = ['Stage', 'Total Value', 'Weighted Value', 'Conversion Rate']
    
    st.dataframe(
        funnel_display,
        use_container_width=True,
        hide_index=True
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
    
    # Add unrestricted after special projects line if enabled
    if enable_special_projects:
        fig.add_trace(go.Scatter(
            x=forecast_df['monthLabel'],
            y=forecast_df['unrestrictedAfterSpecial'],
            mode='lines+markers',
            name='Unrestricted After Special Projects',
            line=dict(color='#f59e0b', width=2, dash='dot'),
            marker=dict(size=4)
        ))
    
    # Add total funds line
    fig.add_trace(go.Scatter(
        x=forecast_df['monthLabel'],
        y=forecast_df['totalFunds'],
        mode='lines+markers',
        name='Total Funds',
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
    recovery_df = forecast_df[forecast_df['month'] > 0].copy()
    
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
    
    # Add annotations showing when fixed staff costs change
    prev_staff_cost = None
    annotations = []
    for idx, row in recovery_df.iterrows():
        current_staff_cost = float(row['fixedStaffCosts'])
        if prev_staff_cost is not None and current_staff_cost != prev_staff_cost:
            annotations.append({
                'x': row['monthLabel'],
                'y': current_staff_cost,
                'text': f"Cost change to £{current_staff_cost:,.0f}",
                'showarrow': True,
                'arrowhead': 2,
                'ax': 0,
                'ay': -40,
                'font': {'color': 'purple', 'size': 10}
            })
        prev_staff_cost = current_staff_cost
    
    fig2.update_layout(
        barmode='stack',
        height=350,
        xaxis_title="Month",
        yaxis_title="Amount (£)",
        hovermode='x unified',
        yaxis=dict(tickformat='£,.0f'),
        annotations=annotations
    )
    
    st.plotly_chart(fig2, use_container_width=True)
    
    # Monthly Breakdown Table
    st.markdown("---")
    st.subheader("Detailed Monthly Breakdown")
    
    # Prepare display dataframe
    display_df = forecast_df[forecast_df['month'] > 0].copy()
    
    # Select columns based on whether special projects is enabled
    if enable_special_projects:
        display_df = display_df[[
            'monthLabel', 'totalIncome', 'projectStaffCosts', 'projectExpenses', 
            'projectContribution', 'staffRecovery', 'unrecoveredStaffCosts',
            'fixedBackOfficeCosts', 'costsFromContribution', 'netPosition', 'reserveDeposit',
            'unrestrictedReserves', 'specialProjectsCost', 'unrestrictedAfterSpecial',
            'restrictedFunds', 'totalFunds'
        ]].copy()
        
        display_df.columns = [
            'Month', 'Income', 'Project Staff', 'Project Expenses', 
            'Contribution', 'Staff Recovery', 'Unrecovered Staff',
            'Back Office', 'Costs from Contrib.', 'Net Position', 'Deposits',
            'Unrestricted', 'Special Projects', 'Unres. After Special',
            'Restricted Funds', 'Total Funds'
        ]
        
        currency_cols = ['Income', 'Project Staff', 'Project Expenses', 'Contribution',
                         'Staff Recovery', 'Unrecovered Staff', 'Back Office', 
                         'Costs from Contrib.', 'Net Position', 'Deposits', 'Unrestricted',
                         'Special Projects', 'Unres. After Special', 'Restricted Funds', 'Total Funds']
    else:
        display_df = display_df[[
            'monthLabel', 'totalIncome', 'projectStaffCosts', 'projectExpenses', 
            'projectContribution', 'staffRecovery', 'unrecoveredStaffCosts',
            'fixedBackOfficeCosts', 'costsFromContribution', 'netPosition', 'reserveDeposit',
            'unrestrictedReserves', 'restrictedFunds', 'totalFunds'
        ]].copy()
        
        display_df.columns = [
            'Month', 'Income', 'Project Staff', 'Project Expenses', 
            'Contribution', 'Staff Recovery', 'Unrecovered Staff',
            'Back Office', 'Costs from Contrib.', 'Net Position', 'Deposits',
            'Unrestricted', 'Restricted Funds', 'Total Funds'
        ]
        
        currency_cols = ['Income', 'Project Staff', 'Project Expenses', 'Contribution',
                         'Staff Recovery', 'Unrecovered Staff', 'Back Office', 
                         'Costs from Contrib.', 'Net Position', 'Deposits', 'Unrestricted', 
                         'Restricted Funds', 'Total Funds']
    
    # Format currency columns
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
- **Row 3, starting Column B:** Month headers matching your pipeline (e.g. Jan_2026, Feb_2026 ... covering 18 months from your selected start month)
- **Row 4, starting Column B:** Income values for each month
- **Row 5, starting Column B:** Staff cost values for each month
- **Row 6, starting Column B:** Expense values for each month

**Valid Clusters:** 
- Secured income (100% - excluded from funnel, already secured)
- Contracting (100%)
- Negotiating (90%)
- Proposals out for decision
- High likelihood projects in development
- Medium likelihood projects in development
- Ideas at development stage

**Pipeline Funnel Stages:**
1. **All Opportunities** - All pipeline items except Secured income
2. **Identified Income** - Medium/High likelihood + Proposals + Negotiating + Contracting
3. **Proposals** - Proposals out for decision + Negotiating + Contracting
4. **Negotiating** - Negotiating + Contracting
5. **Contracting** - Contracting only

**New Features:**
- **Toggle Opportunities:** Turn individual projects on/off in the model
- **Reserve Deposits:** Add up to 4 one-time deposits to unrestricted reserves
- **Cost Changes:** Specify up to 4 changes to fixed costs throughout the forecast period
- **Special Projects:** Enable additional monthly costs that reduce unrestricted reserves
- **Pipeline Funnel:** Visualize pipeline value at each stage with total and weighted values
""")
