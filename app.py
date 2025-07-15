import streamlit as st
import pandas as pd
import plotly.express as px
import io
from pandas import Timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import numpy as np

# Function to process time formats (handle strings, Timedelta, and numbers)
def clean_time_format(value):
    try:
        if isinstance(value, Timedelta):
            return value.total_seconds() / 3600.0
        elif isinstance(value, str) and ":" in value:
            return float(value.replace(":", "."))
        elif pd.isna(value):
            return 0.0
        else:
            return float(value)
    except (ValueError, TypeError):
        return 0.0

# Function to process the uploaded Excel file
def process_excel_file(file):
    try:
        # Read Excel file
        df = pd.read_excel(file, sheet_name=0)
        
        # Clean column names
        df.columns = [col.strip() for col in df.columns]
        
        # Apply time format cleaning to relevant columns (D and H)
        if len(df.columns) >= 8:  # Ensure columns D (index 3) and H (index 7) exist
            df.iloc[:, 3] = df.iloc[:, 3].apply(clean_time_format)  # Heures Normales
            df.iloc[:, 7] = df.iloc[:, 7].apply(clean_time_format)  # Congé annuel
        
        # Initialize Presence column
        df['Presence'] = 0.0
        
        # Process workers based on Jours de Présences (column L, index 11)
        for idx, row in df.iterrows():
            jours_presence = row.iloc[11] if len(df.columns) > 11 else None
            if pd.isna(jours_presence) or jours_presence == '':
                # Hourly worker: Apply formula D - ((H * 8) + 14)
                heures_normales = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0.0
                conge_annuel = float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0
                df.at[idx, 'Presence'] = heures_normales - ((conge_annuel * 8) + 14)
            else:
                # Daily worker: Use Jours de Présences directly
                df.at[idx, 'Presence'] = float(jours_presence) if pd.notna(jours_presence) else 0.0
        
        # Convert Service en cours to string and handle NaN
        if 'Service en cours' in df.columns:
            df['Service en cours'] = df['Service en cours'].fillna('Unknown').astype(str)
        
        # Create output DataFrame with required columns
        output_df = pd.DataFrame({
            'Matr': df['Matricule'],
            'Nom & Prénom': df['Nom'] + ' ' + df['Prénom'],
            'Présence': df['Presence'],
            'H Supp 75%': df.get('H SUP 75% Hebdomadaire 24/5-23/6', ''),
            'Congé': df.get('Congé annuel 24/5-23/6', ''),
            'Congé spécial': '',
            'Feries': df.get('Nombre de jours fériés 24/5-23/6', ''),
            'Chom Tech': df.get('Chomage technique 24/5-23/6', ''),
            'Prime de Rendement': '',
            'Avance sur salaire': '',
            'Prêt /Salaire': '',
            'Observations': 'Prime anciennete:30 dt',
            'Service en cours': df.get('Service en cours', 'Unknown')
        })
        
        return df, output_df
    except Exception as e:
        st.error(f"Failed to process Excel file: {str(e)}")
        return None, None

# Function to create styled Excel file
def create_styled_excel(df, output_buffer):
    wb = Workbook()
    ws = wb.active
    ws.title = "Feuil1"
    
    # Define font
    calibri_font = Font(name='Calibri', size=11, bold=False)
    calibri_bold = Font(name='Calibri', size=11, bold=True)
    
    # Add title, subtitle, and reference line
    ws['A1'] = "PRINCE MEDICAL INDUSTRY PMI SARL"
    ws.merge_cells('A1:L1')
    ws['A1'].font = calibri_bold
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    
    ws['A2'] = "ETAT DE POINTAGE Juin 2025"
    ws.merge_cells('A2:L2')
    ws['A2'].font = calibri_bold
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
    
    ws['A3'] = "Présence 192 heures"
    ws.merge_cells('A3:L3')
    ws['A3'].font = calibri_font
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')
    
    ws.append([])  # Blank row
    
    # Define header row
    headers = ["Matr", "Nom & Prénom", "Présence", "H Supp 75%", "Congé", "Congé spécial",
               "Feries", "Chom Tech", "Prime de Rendement", "Avance sur salaire", "Prêt /Salaire", "Observations"]
    
    # Group employees into chunks of 20
    chunk_size = 20
    row_idx = 5
    table_count = 0
    
    for i in range(0, len(df), chunk_size):
        # Add header
        ws.append(headers)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col)
            cell.font = calibri_bold
            cell.alignment = Alignment(horizontal='center')
        
        # Add data
        chunk = df.iloc[i:i + chunk_size]
        for _, row in chunk.iterrows():
            row_idx += 1
            ws.append([row[col] for col in headers])
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).font = calibri_font
        
        # Create table
        table_ref = f"A{row_idx - len(chunk)}:L{row_idx}"
        table = Table(displayName=f"Table{table_count}", ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws.add_table(table)
        table_count += 1
        
        # Add visa lines
        row_idx += 1
        ws.append([])  # Blank row
        row_idx += 1
        ws.append(["Visa RAP"] + [""] * 8 + ["Visa Direction Général"])
        ws.cell(row=row_idx, column=1).font = calibri_font
        ws.cell(row=row_idx, column=10).font = calibri_font
        row_idx += 1
        ws.append([])  # Blank row
        row_idx += 1
    
    # Adjust column widths based on data rows (starting from row 5)
    for col_idx in range(1, len(headers) + 1):
        column_letter = get_column_letter(col_idx)
        max_length = 0
        for row in range(5, ws.max_row + 1):  # Start from header row
            cell = ws.cell(row=row, column=col_idx)
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(output_buffer)

# Streamlit app
st.title("HR Attendance Dashboard")
st.write("Navigate using the sidebar to view processed data or statistics.")

# Initialize session state
if 'output_df' not in st.session_state:
    st.session_state.output_df = None
if 'original_df' not in st.session_state:
    st.session_state.original_df = None
if 'page' not in st.session_state:
    st.session_state.page = "Upload"

# Sidebar for navigation with buttons
st.sidebar.header("Navigation")
if st.session_state.page != "Upload":
    if st.sidebar.button("Data Processing"):
        st.session_state.page = "Data Processing"
        st.rerun()
    if st.sidebar.button("Stats"):
        st.session_state.page = "Stats"
        st.rerun()
else:
    st.sidebar.write("Upload a file to enable navigation.")

# Handle file upload and page navigation
if st.session_state.page == "Upload" and st.session_state.output_df is None:
    # Center the file uploader with wider span
    col1, col2, col3 = st.columns([1, 4, 1])
    with col2:
        uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        original_df, output_df = process_excel_file(uploaded_file)
        if output_df is not None:
            st.session_state.original_df = original_df
            st.session_state.output_df = output_df
            st.session_state.page = "Data Processing"
            st.rerun()  # Redirect to Data Processing page
else:
    # Service filter for Stats page only
    if st.session_state.output_df is not None:
        try:
            services = sorted([str(x) for x in st.session_state.output_df['Service en cours'].unique() if pd.notna(x)])
            services = ["All Services"] + services  # Add All Services option
        except Exception as e:
            st.error(f"Error sorting services: {str(e)}")
            services = []
        
        if st.session_state.page == "Data Processing":
            filtered_df = st.session_state.output_df  # No filtering
        else:  # Stats page
            selected_service = st.selectbox(
                "Filter by Service", 
                services,
                index=0  # Default to All Services
            )
            if selected_service == "All Services":
                filtered_df = st.session_state.output_df
            else:
                filtered_df = st.session_state.output_df[
                    st.session_state.output_df['Service en cours'] == selected_service
                ]
    else:
        filtered_df = None

    # Page logic
    if st.session_state.page == "Data Processing" and st.session_state.output_df is not None:
        # Display processed data
        st.subheader("Processed Data")
        st.dataframe(filtered_df[['Matr', 'Nom & Prénom', 'Présence', 'H Supp 75%', 
                                 'Congé', 'Feries', 'Chom Tech', 'Observations']])
        
        # Download processed data
        output = io.BytesIO()
        create_styled_excel(filtered_df, output)
        st.download_button(
            label="Download Processed Data",
            data=output.getvalue(),
            file_name="etat_de_pointage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    elif st.session_state.page == "Stats" and st.session_state.output_df is not None:
        st.subheader("Attendance Statistics")
        
        # Summary statistics
        st.write("**Summary Statistics for Presence**")
        stats = filtered_df['Présence'].describe()
        st.write(stats)
        
        # Anomaly detection (using z-scores)
        st.write("**Anomaly Detection**")
        worker_type = st.session_state.original_df['Jours de Présences 24/5-23/6'].apply(
            lambda x: 'Hourly' if pd.isna(x) or x == '' else 'Daily'
        )
        filtered_worker_type = worker_type[st.session_state.output_df['Service en cours'].isin(
            [selected_service] if selected_service != "All Services" else services[1:]
        )]
        hourly_workers = filtered_df[filtered_worker_type == 'Hourly']
        daily_workers = filtered_df[filtered_worker_type == 'Daily']
        
        # Top 5 and Worst 5 for Hourly Workers
        if not hourly_workers.empty:
            hourly_workers = hourly_workers.copy()
            hourly_workers['Z-Score'] = np.abs((hourly_workers['Présence'] - hourly_workers['Présence'].mean()) / 
                                              hourly_workers['Présence'].std())
            st.write("**Hourly Workers**")
            
            # Top 5 (highest presence)
            top_5_hourly = hourly_workers.nlargest(5, 'Présence')[['Matr', 'Nom & Prénom', 'Présence', 'Service en cours']]
            st.write("Top 5 Hourly Workers (Highest Presence)")
            st.dataframe(top_5_hourly)
            
            # Worst 5 (lowest presence)
            worst_5_hourly = hourly_workers.nsmallest(5, 'Présence')[['Matr', 'Nom & Prénom', 'Présence', 'Service en cours']]
            st.write("Worst 5 Hourly Workers (Lowest Presence)")
            st.dataframe(worst_5_hourly)
        
        # Top 5 and Worst 5 for Daily Workers
        if not daily_workers.empty:
            daily_workers = daily_workers.copy()
            daily_workers['Z-Score'] = np.abs((daily_workers['Présence'] - daily_workers['Présence'].mean()) / 
                                             daily_workers['Présence'].std())
            st.write("**Daily Workers**")
            
            # Top 5 (highest presence)
            top_5_daily = daily_workers.nlargest(5, 'Présence')[['Matr', 'Nom & Prénom', 'Présence', 'Service en cours']]
            st.write("Top 5 Daily Workers (Highest Presence)")
            st.dataframe(top_5_daily)
            
            # Worst 5 (lowest presence)
            worst_5_daily = daily_workers.nsmallest(5, 'Présence')[['Matr', 'Nom & Prénom', 'Présence', 'Service en cours']]
            st.write("Worst 5 Daily Workers (Lowest Presence)")
            st.dataframe(worst_5_daily)
        
        # Visualizations
        st.subheader("Visualizations")
        
        # Bar chart: Average Presence by Service
        worker_type = st.session_state.original_df['Jours de Présences 24/5-23/6'].apply(
            lambda x: 'Hourly' if pd.isna(x) or x == '' else 'Daily'
        )
        df_with_worker_type = filtered_df.copy()
        df_with_worker_type['Worker Type'] = worker_type
        avg_presence_by_service = df_with_worker_type.groupby(['Service en cours', 'Worker Type'])['Présence'].mean().reset_index()
        fig_bar = px.bar(
            avg_presence_by_service,
            x='Service en cours',
            y='Présence',
            color='Worker Type',
            title="Average Presence by Service (Hours/Days)",
            labels={'Présence': 'Average Presence (Hours/Days)', 'Service en cours': 'Service'},
            barmode='group',
            color_discrete_sequence=['#1f77b4', '#ff7f0e'],
            range_y=[0, 300]  # Cap y-axis at 300
        )
        st.plotly_chart(fig_bar)
        
        # Box plot: Presence Distribution by Service
        fig_box = px.box(
            filtered_df,
            x='Service en cours',
            y='Présence',
            title="Presence Distribution by Service",
            labels={'Présence': 'Presence (Hours/Days)', 'Service en cours': 'Service'},
            hover_data=['Nom & Prénom'],
            range_y=[0, 300]  # Cap y-axis at 300
        )
        st.plotly_chart(fig_box)
        
        # Histogram of Presence
        st.write("**Presence Distribution**")
        fig_hist = px.histogram(
            filtered_df,
            x='Présence',
            color=filtered_worker_type,
            title="Distribution of Presence (Hourly vs Daily Workers)",
            labels={'Présence': 'Presence (Hours/Days)', 'color': 'Worker Type'},
            nbins=20
        )
        st.plotly_chart(fig_hist)
    
    else:
        st.info("Please upload an Excel file to begin.")