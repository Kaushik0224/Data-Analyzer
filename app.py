import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
import numpy as np
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet

# Page configuration
st.set_page_config(
    page_title="Data Analyzer Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Fixed theme: Dark Ocean only (high-contrast)
st.session_state["theme_choice"] = "Dark Ocean"
grad1 = "linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%)"
grad2 = "linear-gradient(135deg, #1f1c2c 0%, #928dab 100%)"
btn_grad = "linear-gradient(135deg, #5ac8fa 0%, #4c9df0 100%)"
text, muted = "#e6edf3", "#9aa4b2"
surface = "rgba(11,14,18,0.7)"
surface_border = "#2a2f36"

st.markdown(f"""
    <style>
    :root {{ --grad-1:{grad1}; --grad-2:{grad2}; --btn-grad:{btn_grad}; --card-br:{surface_border}; --text:{text}; --muted:{muted}; --surface:{surface}; }}
    .stApp {{ background: var(--grad-1); }}
    .main {{ padding: 1.25rem 2rem; }}
    /* Global +1 step text size */
    html, body {{ font-size: 17px; }}
    .stButton>button {{ width:100%; border-radius:10px; border:0; color:#0b0e12; font-weight:700; background:var(--btn-grad); box-shadow:0 6px 18px rgba(90,200,250,0.25); }}
    .stButton>button:hover {{ filter:brightness(1.05); }}
    .card {{ background: var(--grad-2); border:1px solid var(--card-br); border-radius:14px; padding:16px; }}
    .table-container {{ max-height:480px; overflow:auto; border:1px solid var(--card-br); border-radius:10px; padding:8px; background:var(--surface); backdrop-filter:blur(6px); }}
    .metric {{ background: var(--surface); border:1px solid var(--card-br); border-radius:14px; padding:12px; text-align:center; color:var(--text); }}
    .metric .label {{ color:var(--muted); font-size:12px; }}
    .metric .value {{ color:var(--text); font-size:21px; font-weight:800; }}
    h1, h2, h3 {{ color:var(--text) !important; text-shadow: 0 1px 0 rgba(0,0,0,0.10); }}
    p, span, label, .stMarkdown, .stSelectbox label, .stTextInput label, .stSlider label {{ color:var(--text) !important; }}
    /* Improve contrast for inputs on light theme */
    .stTextInput>div>div>input, .stSelectbox>div>div, .stNumberInput input {{ color: var(--text) !important; }}
    /* File uploader contrast */
    [data-testid="stFileUploaderDropzone"] {{ background: var(--surface); border:1px solid var(--card-br); }}
    [data-testid="stFileUploaderDropzone"] * {{ color: var(--text) !important; }}
    /* Sticky sidebar content */
    [data-testid="stSidebar"] > div:first-child {{ position: sticky; top: 0; }}
    </style>
""", unsafe_allow_html=True)

def main():
    st.title("📊 Data Analyzer Pro")
    st.markdown("Upload your Excel file to get started with data analysis and visualization.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls', 'csv'])
    
    if uploaded_file is not None:
        try:
            # Read the uploaded file
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
                sheet_name = None
            else:  # Excel file with sheet selector
                xls = pd.ExcelFile(uploaded_file)
                sheet_name = st.sidebar.selectbox("Sheet", options=xls.sheet_names, index=0)
                df = xls.parse(sheet_name)
            # Ensure full column visibility in HTML tables
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', None)
            
            # Sidebar filters
            st.sidebar.header("Filters")
            all_columns = list(df.columns)
            show_columns = st.sidebar.multiselect("Columns to display", options=all_columns, default=all_columns)
            
            # Text search across string/object columns
            text_query = st.sidebar.text_input("Text contains (any text column)")
            
            # Numeric range filter
            numeric_cols_all = df.select_dtypes(include=['number']).columns.tolist()
            num_col = st.sidebar.selectbox("Numeric column filter (optional)", options=["<None>"] + numeric_cols_all)
            min_val = max_val = None
            if num_col != "<None>":
                col_series = pd.to_numeric(df[num_col], errors='coerce')
                min_default = float(col_series.min()) if not col_series.dropna().empty else 0.0
                max_default = float(col_series.max()) if not col_series.dropna().empty else 0.0
                min_val, max_val = st.sidebar.slider(
                    f"Range for {num_col}", min_value=min_default, max_value=max_default, value=(min_default, max_default)
                )
            
            # Advanced options
            st.sidebar.header("Advanced")
            auto_cast = st.sidebar.checkbox("Auto type-cast numerics", value=True, help="Try to turn number-like text into numeric")
            na_mode = st.sidebar.selectbox("Missing values", ["None", "Drop rows with NA", "Fill with value"])
            fill_value = None
            if na_mode == "Fill with value":
                fill_value = st.sidebar.text_input("Fill value", value="")
            # Derived column builder
            with st.sidebar.expander("Derived column", expanded=False):
                derived_name = st.text_input("New column name", value="")
                deriv_col_a = st.selectbox("Column A", options=["<None>"] + all_columns)
                deriv_op = st.selectbox("Operation", options=["+", "-", "*", "/"]) 
                deriv_col_b = st.selectbox("Column B", options=["<None>"] + all_columns)

            # Column Type Editor + Rename (instant)
            with st.sidebar.expander("Column Type Editor / Rename", expanded=False):
                edit_col = st.selectbox("Select column", options=["<None>"] + all_columns, help="Choose a column to edit")
                if 'col_renames' not in st.session_state:
                    st.session_state['col_renames'] = {}
                if 'col_types' not in st.session_state:
                    st.session_state['col_types'] = {}
                if edit_col != "<None>":
                    new_name = st.text_input("New name", value=edit_col)
                    dtype_choice = st.selectbox("Target type", ["auto","numeric","categorical","datetime"], index=0)
                    if st.button("Apply", use_container_width=True):
                        if new_name and new_name != edit_col:
                            st.session_state['col_renames'][edit_col] = new_name
                            # also update local list so subsequent selects see new name immediately
                            all_columns = [st.session_state['col_renames'].get(c, c) for c in all_columns]
                        st.session_state['col_types'][new_name if new_name else edit_col] = dtype_choice

            # Apply advanced transforms
            df_work = df.copy()
            # Rename columns first
            if st.session_state.get('col_renames'):
                df_work = df_work.rename(columns=st.session_state['col_renames'])
                all_columns = list(df_work.columns)
            if auto_cast:
                for c in df_work.select_dtypes(include=['object', 'string']).columns:
                    converted = pd.to_numeric(df_work[c], errors='ignore')
                    df_work[c] = converted
            if na_mode == "Drop rows with NA":
                df_work = df_work.dropna()
            elif na_mode == "Fill with value":
                df_work = df_work.fillna(fill_value)
            # Cast columns based on editor selections
            if st.session_state.get('col_types'):
                for col, t in st.session_state['col_types'].items():
                    if col in df_work.columns and t != "auto":
                        if t == "numeric":
                            df_work[col] = pd.to_numeric(df_work[col], errors='coerce')
                        elif t == "categorical":
                            df_work[col] = df_work[col].astype('string')
                        elif t == "datetime":
                            df_work[col] = pd.to_datetime(df_work[col], errors='coerce')
            if derived_name and deriv_col_a != "<None>" and deriv_col_b != "<None>":
                a = pd.to_numeric(df_work[deriv_col_a], errors='coerce')
                b = pd.to_numeric(df_work[deriv_col_b], errors='coerce')
                with pd.option_context('mode.use_inf_as_na', True):
                    if deriv_op == "+":
                        df_work[derived_name] = a + b
                    elif deriv_op == "-":
                        df_work[derived_name] = a - b
                    elif deriv_op == "*":
                        df_work[derived_name] = a * b
                    elif deriv_op == "/":
                        df_work[derived_name] = a / b

            # Preview control
            preview_rows = st.sidebar.number_input("Preview rows", min_value=5, max_value=500, value=50, step=5)

            # Apply filters
            df_filtered = df_work
            if text_query.strip():
                mask = pd.Series(False, index=df_filtered.index)
                for c in df_filtered.select_dtypes(include=['object', 'string']).columns:
                    mask = mask | df_filtered[c].astype(str).str.contains(text_query, case=False, na=False)
                df_filtered = df_filtered[mask]
            if num_col != "<None>" and min_val is not None and max_val is not None:
                numeric_series = pd.to_numeric(df_filtered[num_col], errors='coerce')
                df_filtered = df_filtered[(numeric_series >= min_val) & (numeric_series <= max_val)]
            if show_columns:
                df_filtered = df_filtered[show_columns]
            
            # Presets (filters + sections)
            st.sidebar.header("Presets")
            if 'presets' not in st.session_state:
                st.session_state['presets'] = {}
            preset_names = list(st.session_state['presets'].keys())
            load_choice = st.sidebar.selectbox("Load preset", options=["<None>"] + preset_names)
            new_preset_name = st.sidebar.text_input("Preset name")
            c1, c2, c3 = st.sidebar.columns(3)
            with c1:
                if st.button("Save") and new_preset_name:
                    st.session_state['presets'][new_preset_name] = {
                        'show_columns': show_columns,
                        'text_query': text_query,
                        'num_col': num_col,
                        'num_range': (min_val, max_val),
                        'auto_cast': auto_cast,
                        'na_mode': na_mode,
                        'fill_value': fill_value,
                    }
            with c2:
                if st.button("Load") and load_choice != "<None>":
                    p = st.session_state['presets'][load_choice]
                    show_columns = p.get('show_columns', show_columns)
                    text_query = p.get('text_query', text_query)
                    num_col = p.get('num_col', num_col)
                    if isinstance(p.get('num_range'), tuple):
                        min_val, max_val = p['num_range']
                    auto_cast = p.get('auto_cast', auto_cast)
                    na_mode = p.get('na_mode', na_mode)
                    fill_value = p.get('fill_value', fill_value)
            with c3:
                if st.button("Delete") and load_choice != "<None>":
                    st.session_state['presets'].pop(load_choice, None)

            # Sidebar output controls
            st.sidebar.header("Output")
            section_choices = st.sidebar.multiselect(
                "Sections to show",
                options=["Summary", "Preview", "Statistics", "Analysis", "Visualization", "Profile", "Smart Insights", "Exports", "PDF"],
                default=["Summary"],
            )
            show_summary = "Summary" in section_choices
            show_preview = "Preview" in section_choices
            show_stats = "Statistics" in section_choices
            show_analysis = "Analysis" in section_choices
            show_vis = "Visualization" in section_choices
            show_profile = "Profile" in section_choices
            show_insights = "Smart Insights" in section_choices
            show_exports = "Exports" in section_choices
            show_pdf = "PDF" in section_choices

            # Focus Mode (guided) – runs only the selected goal
            st.sidebar.header("Focus Mode")
            focus_mode = st.sidebar.checkbox("Enable Focus Mode (guided)", value=False)
            focus_goal = None
            if focus_mode:
                focus_goal = st.sidebar.selectbox("Goal", ["Value Counts", "Group By", "Pivot Table", "Correlation"])
                # hide all regular sections; we'll render a focused result below
                show_summary = show_preview = show_stats = show_analysis = show_vis = show_profile = show_insights = show_exports = show_pdf = False

            # SUMMARY (minimal structured)
            if show_summary:
                st.subheader("Summary")
                n_rows, n_cols = df_filtered.shape
                missing = int(df_filtered.isna().sum().sum())
                mem_mb = round(df_filtered.memory_usage(deep=True).sum()/1_048_576, 2)
                c1, c2, c3, c4 = st.columns(4)
                with c1: st.markdown(f"<div class='metric'><div class='label'>Rows</div><div class='value'>{n_rows}</div></div>", unsafe_allow_html=True)
                with c2: st.markdown(f"<div class='metric'><div class='label'>Columns</div><div class='value'>{n_cols}</div></div>", unsafe_allow_html=True)
                with c3: st.markdown(f"<div class='metric'><div class='label'>Missing</div><div class='value'>{missing}</div></div>", unsafe_allow_html=True)
                with c4: st.markdown(f"<div class='metric'><div class='label'>Memory (MB)</div><div class='value'>{mem_mb}</div></div>", unsafe_allow_html=True)
                st.markdown("**Columns:** " + ", ".join([str(c) for c in df_filtered.columns]))
                # quick numeric overview
                num_cols_quick = df_filtered.select_dtypes(include=['number']).columns
                if len(num_cols_quick) > 0:
                    quick = df_filtered[num_cols_quick].agg(['min','max','mean']).round(3)
                    st.markdown('<div class="table-container">' + quick.to_html() + '</div>', unsafe_allow_html=True)

            # PREVIEW
            if show_preview:
                st.subheader("Data Preview")
                st.markdown('<div class="table-container">' + df_filtered.head(int(preview_rows)).to_html(index=False) + '</div>', unsafe_allow_html=True)

            # STATISTICS
            stats = None
            if show_stats:
                st.subheader("Basic Statistics")
                stats = df_filtered.describe(include='all')
                st.markdown('<div class="table-container">' + stats.to_html() + '</div>', unsafe_allow_html=True)

            # PROFILE (comprehensive per-column overview)
            if show_profile:
                st.subheader("Profile")
                prof_rows = []
                for col in df_filtered.columns:
                    s = df_filtered[col]
                    dtype = str(s.dtype)
                    non_null = int(s.notna().sum())
                    missing = int(s.isna().sum())
                    missing_pct = round(float(missing) * 100.0 / max(1, len(s)), 2)
                    unique = int(s.nunique(dropna=True))
                    sample_val = s.dropna().astype(str).head(1).iloc[0] if non_null > 0 else ""
                    min_val = max_val = mean_val = std_val = Q1 = Q3 = iqr_outliers = ""
                    dt_min = dt_max = ""
                    if pd.api.types.is_numeric_dtype(s):
                        ss = pd.to_numeric(s, errors='coerce')
                        min_val = ss.min()
                        max_val = ss.max()
                        mean_val = round(float(ss.mean()), 4) if ss.notna().any() else ""
                        std_val = round(float(ss.std()), 4) if ss.notna().any() else ""
                        q1 = ss.quantile(0.25)
                        q3 = ss.quantile(0.75)
                        iqr = q3 - q1
                        Q1, Q3 = q1, q3
                        if pd.notna(iqr) and iqr != 0:
                            iqr_outliers = int(((ss < (q1 - 1.5*iqr)) | (ss > (q3 + 1.5*iqr))).sum())
                        else:
                            iqr_outliers = 0
                    elif pd.api.types.is_datetime64_any_dtype(s):
                        ds = pd.to_datetime(s, errors='coerce')
                        dt_min = str(ds.min())
                        dt_max = str(ds.max())
                    # top category
                    top_cat = ""
                    if not pd.api.types.is_numeric_dtype(s):
                        vc = s.astype(str).value_counts(dropna=True)
                        if len(vc) > 0:
                            top_cat = f"{vc.index[0]} ({int(vc.iloc[0])})"
                    prof_rows.append([
                        col, dtype, non_null, missing, missing_pct, unique, sample_val,
                        min_val, max_val, mean_val, std_val, Q1, Q3, iqr_outliers, dt_min, dt_max, top_cat
                    ])
                prof_cols = [
                    "Column", "Dtype", "Non-Null", "Missing", "Missing%", "Unique",
                    "Sample", "Min", "Max", "Mean", "Std", "Q1", "Q3", "IQR Outliers",
                    "Date Min", "Date Max", "Top Category"
                ]
                prof_df = pd.DataFrame(prof_rows, columns=prof_cols)
                st.markdown('<div class="table-container">' + prof_df.to_html(index=False) + '</div>', unsafe_allow_html=True)
                # Missing values chart
                miss_counts = df_filtered.isna().sum().sort_values(ascending=False)
                if (miss_counts > 0).any():
                    fig_miss = px.bar(miss_counts[miss_counts>0], x=miss_counts.index[miss_counts>0], y=miss_counts[miss_counts>0].values, title="Missing Values by Column")
                    st.plotly_chart(fig_miss, use_container_width=True)
                # Duplicates preview
                dups = df_filtered[df_filtered.duplicated(keep=False)]
                if not dups.empty:
                    st.markdown(f"Duplicates: <b>{len(dups)}</b>", unsafe_allow_html=True)
                    st.markdown('<div class="table-container">' + dups.head(20).to_html(index=False) + '</div>', unsafe_allow_html=True)

            # SMART INSIGHTS (automatic quick wins)
            if show_insights:
                st.subheader("Smart Insights")
                bullets = []
                # 1) Top category in first categorical column
                cat_cols = df_filtered.select_dtypes(exclude=['number']).columns
                if len(cat_cols) > 0:
                    vc = df_filtered[cat_cols[0]].astype(str).value_counts()
                    if len(vc) > 0:
                        topk = vc.head(5)
                        bullets.append(f"Top categories in {cat_cols[0]}: {', '.join([f'{a} ({b})' for a,b in topk.items()])}")
                        # Build a DataFrame with explicit column names to avoid 'index' name issues
                        top_df = topk.reset_index()
                        top_df.columns = [cat_cols[0], 'count']
                        fig = px.pie(top_df, names=cat_cols[0], values='count', title=f"{cat_cols[0]} distribution (Top 5)")
                        st.plotly_chart(fig, use_container_width=True)
                # 2) Strongest correlation pair
                num = df_filtered.select_dtypes(include=['number'])
                if num.shape[1] >= 2:
                    corr = num.corr(numeric_only=True).abs()
                    np.fill_diagonal(corr.values, 0)
                    i,j = divmod(corr.values.argmax(), corr.shape[1])
                    bullets.append(f"Strongest correlation: {corr.index[i]} vs {corr.columns[j]} = {corr.values[i,j]:.3f}")
                    figc = px.imshow(corr, text_auto=True, aspect="auto", title="Correlation (abs)")
                    st.plotly_chart(figc, use_container_width=True)
                    # Outliers via Z-score
                    z = (num - num.mean())/num.std(ddof=0)
                    out_counts = (np.abs(z) > 3).sum()
                    if (out_counts > 0).any():
                        bullets.append("Potential outliers detected (|z|>3): " + ", ".join([f"{c}:{int(v)}" for c,v in out_counts[out_counts>0].items()]))
                # 3) Simple trend if datetime present
                dt_cols = [c for c in df_filtered.columns if pd.api.types.is_datetime64_any_dtype(df_filtered[c])]
                if len(dt_cols) and num.shape[1] > 0:
                    dtc = dt_cols[0]
                    metric = num.columns[0]
                    ts = df_filtered[[dtc, metric]].copy()
                    ts[dtc] = pd.to_datetime(ts[dtc], errors='coerce')
                    ts = ts.dropna().set_index(dtc).resample('W')[metric].sum()
                    if not ts.empty:
                        bullets.append(f"Weekly trend of {metric} over {dtc} shown below")
                        figt = px.line(ts, y=ts.name, title=f"Weekly {metric}")
                        st.plotly_chart(figt, use_container_width=True)
                if bullets:
                    st.markdown("\n".join([f"- {b}" for b in bullets]))

            # ANALYSIS (user-driven)
            if show_analysis or (focus_mode and focus_goal is not None):
                st.subheader("Analysis")
                analysis_type = focus_goal if focus_mode else st.selectbox(
                    "Choose analysis",
                    ["Value Counts", "Group By", "Pivot Table", "Correlation"],
                )
                if analysis_type == "Value Counts":
                    cat_cols = df_filtered.select_dtypes(exclude=['number']).columns.tolist() or list(df_filtered.columns)
                    target_col = st.selectbox("Column", cat_cols)
                    top_n = st.number_input("Top N", min_value=1, value=10, step=1)
                    counts = df_filtered[target_col].astype(str).value_counts().head(int(top_n))
                    st.markdown('<div class="table-container">' + counts.rename("count").to_frame().to_html() + '</div>', unsafe_allow_html=True)
                    chart_type = st.radio("Chart", ["Bar", "Pie"], horizontal=True)
                    if chart_type == "Bar":
                        fig = px.bar(counts, x=counts.index, y=counts.values, labels={"x": target_col, "y": "count"}, title=f"Top {top_n} {target_col}")
                    else:
                        fig = px.pie(values=counts.values, names=counts.index, title=f"{target_col} distribution")
                    st.plotly_chart(fig, use_container_width=True)
                elif analysis_type == "Group By":
                    group_cols = st.multiselect("Group by columns", options=list(df_filtered.columns))
                    agg_col = st.selectbox("Aggregate column", options=df_filtered.select_dtypes(include=['number']).columns)
                    agg_func = st.selectbox("Aggregation", ["sum", "mean", "min", "max", "count"])
                    if group_cols:
                        grouped = getattr(df_filtered.groupby(group_cols)[agg_col], agg_func)().reset_index()
                        st.markdown('<div class="table-container">' + grouped.to_html(index=False) + '</div>', unsafe_allow_html=True)
                        fig = px.bar(grouped, x=group_cols[0], y=agg_col, color=group_cols[1] if len(group_cols) > 1 else None, title=f"{agg_func} of {agg_col} by {', '.join(group_cols)}")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("Select at least one group-by column.")
                elif analysis_type == "Pivot Table":
                    rows = st.multiselect("Rows", options=list(df_filtered.columns))
                    cols = st.multiselect("Columns", options=list(df_filtered.columns))
                    values = st.selectbox("Values", options=df_filtered.select_dtypes(include=['number']).columns)
                    aggfunc = st.selectbox("Agg", ["sum", "mean", "min", "max", "count"])
                    if rows or cols:
                        pivot = pd.pivot_table(df_filtered, index=rows if rows else None, columns=cols if cols else None, values=values, aggfunc=aggfunc, fill_value=0)
                        st.markdown('<div class="table-container">' + pivot.to_html() + '</div>', unsafe_allow_html=True)
                        fig = px.imshow(pivot, aspect="auto", title="Pivot Heatmap")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("Choose at least rows or columns for pivot.")
                elif analysis_type == "Correlation":
                    num = df_filtered.select_dtypes(include=['number'])
                    if num.shape[1] >= 2:
                        corr = num.corr(numeric_only=True)
                        st.markdown('<div class="table-container">' + corr.to_html() + '</div>', unsafe_allow_html=True)
                        fig = px.imshow(corr, text_auto=True, aspect="auto", title="Correlation Heatmap")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("Need at least 2 numeric columns for correlation.")

            # VISUALIZATION
            if show_vis:
                st.subheader("Data Visualization")
                chart_kind = st.selectbox("Chart type", ["Scatter", "Bar", "Stacked Bar", "Line", "Area", "Histogram", "Box", "Pie"])
                numeric_cols = df_filtered.select_dtypes(include=['number']).columns
                all_cols = df_filtered.columns
                if chart_kind == "Pie":
                    name_col = st.selectbox("Names (category)", options=all_cols)
                    value_mode = st.radio("Pie values", ["Counts", "Numeric column"], horizontal=True)
                    if value_mode == "Numeric column" and len(numeric_cols) > 0:
                        value_col = st.selectbox("Values (numeric)", options=numeric_cols)
                        data = df_filtered.groupby(name_col)[value_col].sum().reset_index()
                        fig = px.pie(data, names=name_col, values=value_col, title=f"{value_col} by {name_col}")
                        st.plotly_chart(fig, use_container_width=True)
                        st.session_state['last_chart'] = fig
                    else:
                        counts = df_filtered[name_col].astype(str).value_counts().reset_index()
                        counts.columns = [name_col, "count"]
                        fig = px.pie(counts, names=name_col, values="count", title=f"{name_col} distribution")
                        st.plotly_chart(fig, use_container_width=True)
                        st.session_state['last_chart'] = fig
                else:
                    if len(numeric_cols) == 0:
                        st.info("No numeric columns found for this chart.")
                    else:
                        if chart_kind == "Scatter" and len(numeric_cols) >= 2:
                            c1, c2 = st.columns(2)
                            with c1:
                                x_axis = st.selectbox('X-axis', numeric_cols)
                            with c2:
                                y_axis = st.selectbox('Y-axis', numeric_cols)
                            fig = px.scatter(df_filtered, x=x_axis, y=y_axis, title=f"{y_axis} vs {x_axis}")
                            st.plotly_chart(fig, use_container_width=True)
                            st.session_state['last_chart'] = fig
                        elif chart_kind == "Bar":
                            x_col = st.selectbox('Category (X)', options=all_cols)
                            y_col = st.selectbox('Value (Y)', options=numeric_cols)
                            data = df_filtered.groupby(x_col)[y_col].sum().reset_index()
                            fig = px.bar(data, x=x_col, y=y_col, title=f"{y_col} by {x_col}")
                            st.plotly_chart(fig, use_container_width=True)
                            st.session_state['last_chart'] = fig
                        elif chart_kind == "Stacked Bar":
                            x_col = st.selectbox('Category (X)', options=all_cols)
                            y_col = st.selectbox('Value (Y)', options=numeric_cols)
                            color_col = st.selectbox('Stack by', options=all_cols)
                            data = df_filtered.groupby([x_col, color_col])[y_col].sum().reset_index()
                            fig = px.bar(data, x=x_col, y=y_col, color=color_col, title=f"{y_col} by {x_col} (stacked)")
                            st.plotly_chart(fig, use_container_width=True)
                            st.session_state['last_chart'] = fig
                        elif chart_kind == "Line":
                            x_col = st.selectbox('X (category/date)', options=all_cols)
                            y_col = st.selectbox('Y (numeric)', options=numeric_cols)
                            color_col = st.selectbox('Color (optional)', options=["<None>"] + list(all_cols))
                            data = df_filtered
                            fig = px.line(data, x=x_col, y=y_col, color=None if color_col=="<None>" else color_col, title=f"{y_col} over {x_col}")
                            st.plotly_chart(fig, use_container_width=True)
                            st.session_state['last_chart'] = fig
                        elif chart_kind == "Area":
                            x_col = st.selectbox('X (category/date)', options=all_cols)
                            y_col = st.selectbox('Y (numeric)', options=numeric_cols)
                            fig = px.area(df_filtered, x=x_col, y=y_col, title=f"{y_col} over {x_col}")
                            st.plotly_chart(fig, use_container_width=True)
                            st.session_state['last_chart'] = fig
                        elif chart_kind == "Histogram":
                            x_col = st.selectbox('Numeric column', options=numeric_cols)
                            fig = px.histogram(df_filtered, x=x_col, nbins=30, title=f"Distribution of {x_col}")
                            st.plotly_chart(fig, use_container_width=True)
                            st.session_state['last_chart'] = fig
                        elif chart_kind == "Box":
                            y_col = st.selectbox('Numeric column', options=numeric_cols)
                            x_col = st.selectbox('Group by (optional)', options=["<None>"] + list(all_cols))
                            if x_col == "<None>":
                                fig = px.box(df_filtered, y=y_col, points='outliers', title=f"Box plot of {y_col}")
                            else:
                                fig = px.box(df_filtered, x=x_col, y=y_col, points='outliers', title=f"{y_col} by {x_col}")
                            st.plotly_chart(fig, use_container_width=True)
                            st.session_state['last_chart'] = fig

            # EXPORTS
            if show_exports:
                st.subheader("Export Data")
                csv_bytes = df_filtered.to_csv(index=False).encode('utf-8')
                st.download_button("Download CSV", data=csv_bytes, file_name="filtered_data.csv", mime="text/csv")
                excel_buf = BytesIO()
                with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                    df_filtered.to_excel(writer, index=False, sheet_name='Data')
                st.download_button("Download Excel", data=excel_buf.getvalue(), file_name="filtered_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # PDF report (landscape) – include filters summary and optional charts
            def _add_page_number(canvas, doc):
                canvas.saveState()
                canvas.setFont('Helvetica', 9)
                canvas.drawRightString(landscape(A4)[0]-doc.rightMargin, 10, f"Page {doc.page}")
                canvas.restoreState()

            def build_pdf(buffer: BytesIO, data: pd.DataFrame, stats_df: pd.DataFrame, include_stats: bool, filters_info: dict, sample_rows: int, include_charts: bool):
                doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
                styles = getSampleStyleSheet()
                elements = []
                elements.append(Paragraph("Data Analyzer Pro - Summary Report", styles['Title']))
                elements.append(Spacer(1, 12))
                elements.append(Paragraph(f"Rows: {len(data)} | Columns: {len(data.columns)}", styles['Normal']))
                elements.append(Spacer(1, 12))
                # Filters summary
                if filters_info:
                    elements.append(Paragraph("Filters", styles['Heading2']))
                    filt_lines = []
                    if filters_info.get('show_columns'):
                        filt_lines.append(f"Columns: {', '.join(map(str, filters_info['show_columns']))}")
                    if filters_info.get('text_query'):
                        filt_lines.append(f"Text contains: {filters_info['text_query']}")
                    if filters_info.get('num_col') and filters_info.get('num_range'):
                        rng = filters_info['num_range']
                        filt_lines.append(f"Numeric filter: {filters_info['num_col']} in [{rng[0]}, {rng[1]}]")
                    if 'auto_cast' in filters_info:
                        filt_lines.append(f"Auto type-cast numerics: {filters_info['auto_cast']}")
                    if 'na_mode' in filters_info and filters_info['na_mode'] != 'None':
                        fv = filters_info.get('fill_value', '')
                        filt_lines.append(f"Missing values: {filters_info['na_mode']} {('('+str(fv)+')' if filters_info['na_mode']=='Fill with value' else '')}")
                    if filters_info.get('derived_desc'):
                        filt_lines.append(f"Derived column: {filters_info['derived_desc']}")
                    for line in filt_lines:
                        elements.append(Paragraph(line, styles['Normal']))
                    elements.append(Spacer(1, 12))
                if include_stats and stats_df is not None:
                    stats_tbl = [["Statistic"] + list(stats_df.columns)] + [[idx] + [str(v) for v in row] for idx, row in stats_df.iterrows()]
                    t1 = Table(stats_tbl, repeatRows=1)
                    t1.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.grey),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                        ('GRID', (0,0), (-1,-1), 0.25, colors.darkgrey),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ]))
                    elements.append(Paragraph("Basic Statistics", styles['Heading2']))
                    elements.append(t1)
                    elements.append(Spacer(1, 12))
                # Sample rows
                sample = data.head(sample_rows).astype(str)
                # Wrap cells for long text
                header = [Paragraph(str(h), styles['BodyText']) for h in list(sample.columns)]
                body = [[Paragraph(str(v), styles['BodyText']) for v in row] for row in sample.values.tolist()]
                tbl = [ header ] + body
                # Fit to page width by equal column widths
                usable_width = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
                col_count = max(1, len(header))
                col_widths = [usable_width/col_count]*col_count
                t2 = Table(tbl, repeatRows=1, colWidths=col_widths)
                t2.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.grey),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('GRID', (0,0), (-1,-1), 0.25, colors.darkgrey),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.lightgrey]),
                    ]))
                elements.append(Paragraph("Sample Rows", styles['Heading2']))
                elements.append(t2)
                # Optionally include last chart image if available and kaleido present
                if include_charts and 'last_chart' in st.session_state and st.session_state['last_chart'] is not None:
                    try:
                        import plotly.io as pio
                        img_bytes = st.session_state['last_chart'].to_image(format='png', scale=2)
                        from reportlab.lib.utils import ImageReader
                        elements.append(Spacer(1, 12))
                        elements.append(Paragraph("Selected Chart", styles['Heading2']))
                        elements.append(Spacer(1, 6))
                        elements.append(RLImage(ImageReader(BytesIO(img_bytes)), width=landscape(A4)[0]-doc.leftMargin-doc.rightMargin, height=300))
                    except Exception:
                        elements.append(Spacer(1, 12))
                        elements.append(Paragraph("Chart image could not be embedded (requires kaleido).", styles['Italic']))
                doc.build(elements, onFirstPage=_add_page_number, onLaterPages=_add_page_number)
            
            if show_pdf:
                # Controls for PDF generation
                include_filters = st.sidebar.checkbox("Include filters in PDF", value=True)
                include_charts = st.sidebar.checkbox("Include selected chart in PDF", value=False, help="Embeds the last rendered chart if possible")
                sample_rows = st.sidebar.number_input("PDF sample rows", min_value=5, max_value=100, value=20, step=5)
                pdf_buf = BytesIO()
                derived_desc = None
                if 'derived_name' in locals() and derived_name:
                    derived_desc = f"{derived_name} = {deriv_col_a} {deriv_op} {deriv_col_b}"
                filters_payload = {
                    'show_columns': show_columns,
                    'text_query': text_query,
                    'num_col': num_col,
                    'num_range': (min_val, max_val) if num_col != "<None>" else None,
                    'auto_cast': auto_cast,
                    'na_mode': na_mode,
                    'fill_value': fill_value,
                    'derived_desc': derived_desc,
                } if include_filters else {}
                build_pdf(pdf_buf, df_filtered, stats, include_stats=show_stats, filters_info=filters_payload, sample_rows=int(sample_rows), include_charts=include_charts)
                st.download_button("Download PDF Report", data=pdf_buf.getvalue(), file_name="report.pdf", mime="application/pdf")
                
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    else:
        st.info("Please upload a file to begin analysis.")

if __name__ == "__main__":
    main()
