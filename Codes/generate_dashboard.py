import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json

def create_dashboard(excel_file="Master_Compliance_Report.xlsx", output_html="Compliance_Dashboard.html"):
    print("Generating Detailed Dashboard...")
    
    profile_name = "Overall"
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
            profile_name = config.get("profile_name", "Overall")
    except:
        pass
        
    try:
        commit_df = pd.read_excel(excel_file, sheet_name='Master Data - All Months')
        prod_df = pd.read_excel(excel_file, sheet_name='Production Data')
    except Exception as e:
        print(f"Error reading {excel_file}: {e}")
        return

    # Find the right columns
    def get_col(df, keywords):
        for col in df.columns:
            if all(k in str(col).lower() for k in keywords):
                return col
        return None

    # Commitment Columns
    c_gr = get_col(commit_df, ['planned gr target'])
    c_cust = get_col(commit_df, ['clean customer name'])
    c_month = get_col(commit_df, ['month'])
    c_org = get_col(commit_df, ['sale'])
    c_pcode = get_col(commit_df, ['p code'])

    # Production Columns
    p_qty = get_col(prod_df, ['quantity'])
    p_cust = get_col(prod_df, ['clean customer name'])
    p_month = get_col(prod_df, ['month'])
    p_org = get_col(prod_df, ['sale'])
    p_pcode = get_col(prod_df, ['p code'])

    if not all([c_gr, c_cust, c_month, p_qty, p_cust, p_month]):
        print("Missing required core columns for dashboard.")
        return
        
    # --- Data Processing ---
    # Monthly
    m_commit = commit_df.groupby(c_month)[c_gr].sum().reset_index().rename(columns={c_month: 'Month', c_gr: 'Planned (MT)'})
    m_prod = prod_df.groupby(p_month)[p_qty].sum().reset_index().rename(columns={p_month: 'Month', p_qty: 'Actual (MT)'})
    monthly = pd.merge(m_commit, m_prod, on='Month', how='outer').fillna(0)
    month_order = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March']
    monthly['Month'] = pd.Categorical(monthly['Month'], categories=[m for m in month_order if m in monthly['Month'].values], ordered=True)
    monthly = monthly.sort_values('Month')
    monthly['Compliance %'] = (monthly['Actual (MT)'] / monthly['Planned (MT)'] * 100).fillna(0)

    # Sales Org
    if c_org and p_org:
        o_commit = commit_df.groupby(c_org)[c_gr].sum().reset_index().rename(columns={c_org: 'Sales Org', c_gr: 'Planned (MT)'})
        o_prod = prod_df.groupby(p_org)[p_qty].sum().reset_index().rename(columns={p_org: 'Sales Org', p_qty: 'Actual (MT)'})
        sales_org = pd.merge(o_commit, o_prod, on='Sales Org', how='outer').fillna(0)
        sales_org['Compliance %'] = (sales_org['Actual (MT)'] / sales_org['Planned (MT)'] * 100).fillna(0)
    else:
        sales_org = pd.DataFrame()

    # Customer
    cu_commit = commit_df.groupby(c_cust)[c_gr].sum().reset_index().rename(columns={c_cust: 'Customer', c_gr: 'Planned (MT)'})
    cu_prod = prod_df.groupby(p_cust)[p_qty].sum().reset_index().rename(columns={p_cust: 'Customer', p_qty: 'Actual (MT)'})
    customers = pd.merge(cu_commit, cu_prod, on='Customer', how='outer').fillna(0)
    customers['Compliance %'] = (customers['Actual (MT)'] / customers['Planned (MT)'] * 100).fillna(0)
    
    # P Code
    if c_pcode and p_pcode:
        pc_commit = commit_df.groupby(c_pcode)[c_gr].sum().reset_index().rename(columns={c_pcode: 'Product Code', c_gr: 'Planned (MT)'})
        pc_prod = prod_df.groupby(p_pcode)[p_qty].sum().reset_index().rename(columns={p_pcode: 'Product Code', p_qty: 'Actual (MT)'})
        pcodes = pd.merge(pc_commit, pc_prod, on='Product Code', how='outer').fillna(0)
        pcodes['Compliance %'] = (pcodes['Actual (MT)'] / pcodes['Planned (MT)'] * 100).fillna(0)
    else:
        pcodes = pd.DataFrame()

    # --- KPIs ---
    total_planned = monthly['Planned (MT)'].sum()
    total_actual = monthly['Actual (MT)'].sum()
    overall_comp = (total_actual / total_planned * 100) if total_planned else 0
    shortfall = total_planned - total_actual
    active_customers = len(customers[customers['Planned (MT)'] > 0])

    # --- Charts ---
    # 1. Gauge
    fig_gauge = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = overall_comp,
        number = {'suffix': "%", 'font': {'size': 40, 'color': '#0033A0'}},
        delta = {'reference': 100, 'position': "bottom", 'suffix': "%"},
        gauge = {
            'axis': {'range': [0, max(150, overall_comp)], 'tickwidth': 1, 'tickcolor': "darkblue"},
            'bar': {'color': "#0033A0"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, 80], 'color': '#ff4d4d'},
                {'range': [80, 95], 'color': '#ffd633'},
                {'range': [95, max(150, overall_comp)], 'color': '#33cc33'}
            ]
        }
    ))
    fig_gauge.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20), paper_bgcolor='rgba(0,0,0,0)', font={'family': "Inter, sans-serif"})

    # 2. Monthly Trend (Combo Chart)
    fig_month = go.Figure()
    fig_month.add_trace(go.Bar(x=monthly['Month'], y=monthly['Planned (MT)'], name='Planned (MT)', marker_color='#b3c6ff'))
    fig_month.add_trace(go.Scatter(x=monthly['Month'], y=monthly['Actual (MT)'], name='Actual (MT)', mode='lines+markers', line=dict(color='#0033A0', width=3), marker=dict(size=8)))
    fig_month.update_layout(xaxis_title='', yaxis_title='Metric Tons', hovermode="x unified", height=350, margin=dict(l=20, r=20, t=20, b=20), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))

    # 3. Sales Org Donut
    if not sales_org.empty:
        fig_org = px.pie(sales_org, values='Planned (MT)', names='Sales Org', hole=0.5, color_discrete_sequence=px.colors.qualitative.Prism)
        fig_org.update_traces(textposition='inside', textinfo='percent+label')
        fig_org.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20), paper_bgcolor='rgba(0,0,0,0)', showlegend=False)
    else:
        fig_org = go.Figure()

    # 4. Top 15 Customers
    top_cust = customers.sort_values('Planned (MT)', ascending=False).head(15).sort_values('Planned (MT)', ascending=True)
    fig_cust = go.Figure()
    fig_cust.add_trace(go.Bar(y=top_cust['Customer'], x=top_cust['Planned (MT)'], name='Planned', orientation='h', marker_color='#d9e3f0'))
    fig_cust.add_trace(go.Bar(y=top_cust['Customer'], x=top_cust['Actual (MT)'], name='Actual', orientation='h', marker_color='#0052cc'))
    fig_cust.update_layout(barmode='group', height=500, margin=dict(l=20, r=20, t=20, b=20), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))

    # 5. Product Treemap
    if not pcodes.empty:
        fig_pcode = px.treemap(pcodes[pcodes['Planned (MT)'] > 0], path=[px.Constant("Products"), 'Product Code'], values='Planned (MT)', color='Compliance %', color_continuous_scale='RdYlGn', color_continuous_midpoint=100)
        fig_pcode.update_layout(height=500, margin=dict(l=0, r=0, t=0, b=0), paper_bgcolor='rgba(0,0,0,0)')
    else:
        fig_pcode = go.Figure()

    # Data Table HTML (using bootstrap classes)
    table_html = customers.sort_values('Planned (MT)', ascending=False).to_html(classes='table table-striped table-hover align-middle', index=False, float_format='%.1f')

    # HTML Assembly
    shortfall_color = '#e74c3c' if shortfall > 0 else '#2ecc71'
    
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{profile_name} | Compliance Analytics</title>
    <!-- Fonts & Icons -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- DataTables CSS -->
    <link href="https://cdn.datatables.net/1.13.5/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        body {{ font-family: 'Inter', sans-serif; background-color: #f4f6f9; color: #333; }}
        .navbar {{ background-color: #0033A0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .navbar-brand {{ font-weight: 700; letter-spacing: 1px; color: #fff !important; }}
        .kpi-card {{ border: none; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); transition: transform 0.2s; background: #fff; }}
        .kpi-card:hover {{ transform: translateY(-3px); }}
        .kpi-icon {{ font-size: 2.5rem; opacity: 0.15; position: absolute; right: 20px; top: 20px; }}
        .kpi-value {{ font-size: 2rem; font-weight: 700; color: #2c3e50; margin: 10px 0 5px 0; }}
        .kpi-label {{ color: #7f8c8d; font-size: 0.85rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }}
        .chart-card {{ border: none; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 25px; background: #fff; overflow: hidden; }}
        .chart-header {{ background: #fff; border-bottom: 1px solid #f0f0f0; padding: 15px 20px; font-weight: 600; color: #0033A0; text-transform: uppercase; font-size: 0.9rem; letter-spacing: 0.5px; }}
        .table-container {{ background: #fff; border-radius: 10px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 50px; }}
        /* Custom scrollbar */
        ::-webkit-scrollbar {{ width: 8px; height: 8px; }}
        ::-webkit-scrollbar-track {{ background: #f1f1f1; }}
        ::-webkit-scrollbar-thumb {{ background: #c1c1c1; border-radius: 4px; }}
        ::-webkit-scrollbar-thumb:hover {{ background: #a8a8a8; }}
        .plotly-graph-div {{ width: 100% !important; }}
    </style>
</head>
<body>

    <!-- Header -->
    <nav class="navbar navbar-expand-lg navbar-dark mb-4 p-3">
        <div class="container-fluid px-4">
            <a class="navbar-brand" href="#"><i class="fa-solid fa-industry me-2"></i> TATA STEEL | {profile_name}</a>
            <span class="text-white opacity-75 d-none d-md-block">Automated Intelligence Report</span>
        </div>
    </nav>

    <div class="container-fluid px-4">
        <!-- KPI Row -->
        <div class="row g-4 mb-4">
            <div class="col-xl-3 col-md-6">
                <div class="card kpi-card h-100 p-3">
                    <i class="fa-solid fa-bullseye kpi-icon text-primary"></i>
                    <div class="kpi-label">Total Committed (MT)</div>
                    <div class="kpi-value">{total_planned:,.0f}</div>
                </div>
            </div>
            <div class="col-xl-3 col-md-6">
                <div class="card kpi-card h-100 p-3">
                    <i class="fa-solid fa-dolly kpi-icon text-success"></i>
                    <div class="kpi-label">Total Produced (MT)</div>
                    <div class="kpi-value">{total_actual:,.0f}</div>
                </div>
            </div>
            <div class="col-xl-3 col-md-6">
                <div class="card kpi-card h-100 p-3">
                    <i class="fa-solid fa-chart-line kpi-icon text-info"></i>
                    <div class="kpi-label">Shortfall / Excess</div>
                    <div class="kpi-value" style="color: {shortfall_color};">{-shortfall:,.0f}</div>
                </div>
            </div>
            <div class="col-xl-3 col-md-6">
                <div class="card kpi-card h-100 p-3">
                    <i class="fa-solid fa-users kpi-icon text-warning"></i>
                    <div class="kpi-label">Active Customers</div>
                    <div class="kpi-value">{active_customers}</div>
                </div>
            </div>
        </div>

        <!-- Charts Row 1 -->
        <div class="row g-4">
            <div class="col-xl-4 col-lg-5">
                <div class="card chart-card h-100">
                    <div class="chart-header"><i class="fa-solid fa-gauge-high me-2"></i> Overall Score</div>
                    <div class="card-body d-flex align-items-center justify-content-center">
                        {fig_gauge.to_html(full_html=False, include_plotlyjs='cdn')}
                    </div>
                </div>
            </div>
            <div class="col-xl-8 col-lg-7">
                <div class="card chart-card h-100">
                    <div class="chart-header"><i class="fa-solid fa-calendar-alt me-2"></i> Monthly Production Trend</div>
                    <div class="card-body">
                        {fig_month.to_html(full_html=False, include_plotlyjs='cdn')}
                    </div>
                </div>
            </div>
        </div>

        <!-- Charts Row 2 -->
        <div class="row g-4">
            <div class="col-xl-4 col-lg-5">
                <div class="card chart-card">
                    <div class="chart-header"><i class="fa-solid fa-chart-pie me-2"></i> Sales Organization Share</div>
                    <div class="card-body p-0">
                        {fig_org.to_html(full_html=False, include_plotlyjs='cdn')}
                    </div>
                </div>
                <div class="card chart-card mt-4">
                    <div class="chart-header"><i class="fa-solid fa-boxes-stacked me-2"></i> Product Mix (P Code)</div>
                    <div class="card-body p-0">
                        {fig_pcode.to_html(full_html=False, include_plotlyjs='cdn')}
                    </div>
                </div>
            </div>
            <div class="col-xl-8 col-lg-7">
                <div class="card chart-card h-100">
                    <div class="chart-header"><i class="fa-solid fa-ranking-star me-2"></i> Top 15 Customers (Volume)</div>
                    <div class="card-body p-0">
                        {fig_cust.to_html(full_html=False, include_plotlyjs='cdn')}
                    </div>
                </div>
            </div>
        </div>

        <!-- Data Table Row -->
        <div class="row mt-2">
            <div class="col-12">
                <div class="table-container">
                    <h5 class="mb-4" style="color: #0033A0; font-weight: 600;"><i class="fa-solid fa-table-list me-2"></i> Customer Compliance Details</h5>
                    <div class="table-responsive">
                        {table_html}
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Scripts -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.5/js/dataTables.bootstrap5.min.js"></script>
    <script>
        $(document).ready(function() {{
            $('.table').DataTable({{
                pageLength: 10,
                order: [[1, 'desc']], // Sort by Planned MT descending
                language: {{ search: "", searchPlaceholder: "Search records..." }}
            }});
        }});
    </script>
</body>
</html>'''

    with open(output_html, "w", encoding="utf-8-sig") as f:
        f.write(html)
    
    print(f"Detailed dashboard successfully generated: {output_html}")

if __name__ == "__main__":
    create_dashboard()
