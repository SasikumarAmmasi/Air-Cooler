import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import io
import xlsxwriter 

# ==============================================================================
# STREAMLIT APP CONFIGURATION
# ==============================================================================
st.set_page_config(page_title="Air Cooler Performance Analysis", layout="wide")
st.title("Air Cooler Performance Curve Analysis (Multi-Sheet Report Generation)")

# ==============================================================================
# 1. INPUT DATA AND CONSTANTS
# ==============================================================================

# File upload
uploaded_file = st.file_uploader("Upload the Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # Get input constants from user
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        RATED_POWER = st.number_input("Rated Power (kW)", value=30.0, step=0.1)
    
    with col2:
        DESIGN_DUTY = st.number_input("Design Air Cooler Duty (kcal/hr)", value=3350000.0, step=1000.0)
    
    with col3:
        DESIGN_UA = st.number_input("Design UA (kcal/hr.m².°C)", value=100000.0, step=100.0)
    
    with col4:
        DESIGN_TS_OUTLET_TEMP = st.number_input("Design TS Outlet Temp (°C)", value=37.0, step=0.1)
    
    if st.button("Generate Combined Excel Report with Plots"):
        
        # Buffer to hold the final Excel file
        output = io.BytesIO()
        
        # Use ExcelWriter with xlsxwriter engine for image support
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            # Get the workbook object for image insertion
            workbook = writer.book
            
            with st.spinner("Processing data, generating plots, and compiling Excel report..."):
                try:
                    # Load ALL sheets into a dictionary
                    all_sheets_data = pd.read_excel(uploaded_file, sheet_name=None)

                    # --- Color/Style Maps (Defined once) ---
                    temp_colors = {
                        50.0: '#1f77b4',  # Blue
                        55.0: '#ff7f0e',  # Orange
                        60.0: '#2ca02c',  # Green
                        65.0: '#d62728',  # Red
                        70.0: '#9467bd',  # Purple
                        75.0: '#8c564b'   # Brown
                    }
                    
                    # --------------------------------------------------------------------------
                    # 🔁 Iterate over each sheet
                    # --------------------------------------------------------------------------
                    for sheet_name, df in all_sheets_data.items():
                        st.header(f"Processing Sheet: **{sheet_name}**")
                        
                        # --- Data Cleaning and Standardization (same as prior correct code) ---
                        df.columns = df.columns.str.strip()

                        df = df.rename(columns={
                            'TS Gas Mass Flow (kg/h)': 'Mass Flow Rate (kg/hr)',
                            'TS Inlet Temperature (Deg C)': 'TS Inlet Temp (Deg C)',
                            'UA (kcal/hr.m².°C)': 'Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)',
                            'HE Duty (kcal/h)': 'Heat Exchanger Duty (kcal/hr)',
                            'Brake Power/Fan, Summer (kW)': 'Break Power/Fan Summer (kW)',
                            'Brake Power/Fan, Winter (kW)': 'Break Power/Fan Winter (kW)',
                            'TS Outlet Temperature (Deg C)': 'TS Outlet Temperature (Deg C)',
                            'Air Mass Flow (kg/h)': 'Air Mass Flow (kg/h)'
                        })

                        cols_to_convert = [
                            'Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)',
                            'Heat Exchanger Duty (kcal/hr)',
                            'Break Power/Fan Summer (kW)',
                            'Break Power/Fan Winter (kW)',
                            'TS Outlet Temperature (Deg C)'
                        ]
                        for col in cols_to_convert:
                            if col in df.columns:
                                df[col] = pd.to_numeric(df[col], errors='coerce')

                        df.dropna(subset=['Mass Flow Rate (kg/hr)', 'TS Inlet Temp (Deg C)', 'Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)', 'Heat Exchanger Duty (kcal/hr)', 'TS Outlet Temperature (Deg C)'], inplace=True)
                        
                        if df.empty:
                            st.warning(f"Sheet '{sheet_name}' is empty or contains no valid data after cleaning. Skipping plots for this sheet.")
                            st.markdown("---")
                            continue

                        df['Heat Exchanger Duty (kcal/hr)'] = df['Heat Exchanger Duty (kcal/hr)'].abs()
                        grouped = df.groupby('TS Inlet Temp (Deg C)')

                        # --- Write Data to Excel ---
                        # Use a clean sheet name (Excel sheet names are limited to 31 chars)
                        clean_sheet_name = sheet_name.replace('/', '-').replace('\\', '-')[:31]
                        df.to_excel(writer, sheet_name=clean_sheet_name, index=False, startrow=0, startcol=0)
                        
                        # Get the worksheet object for image insertion
                        worksheet = writer.sheets[clean_sheet_name]
                        data_end_row = len(df) + 1 # Last row of data (0-indexed + header row)
                        
                        
                        # ==============================================================================
                        # 2. PLOT 1 GENERATION (UA and Duty)
                        # ==============================================================================
                        fig1, ax1 = plt.subplots(figsize=(14, 8))
                        fig1.suptitle(f'Sheet: {sheet_name} | Performance Curve: UA and Heat Duty vs. Mass Flow Rate', fontsize=16, fontweight='bold')

                        ax1.set_xlabel('Mass Flow Rate (kg/hr)', fontsize=12)
                        ax1.set_ylabel('Service Overall Heat Transfer Coefficient (UA) (kcal/hr.m².°C)', color=temp_colors.get(list(temp_colors.keys())[0], 'k'), fontsize=12)
                        ax1.tick_params(axis='y', labelcolor=temp_colors.get(list(temp_colors.keys())[0], 'k'))

                        ax3 = ax1.twinx()
                        duty_color = 'green'
                        ax3.spines['right'].set_color(duty_color)
                        ax3.set_ylabel('Heat Exchanger Duty (kcal/hr)', color=duty_color, fontsize=12)
                        ax3.tick_params(axis='y', labelcolor=duty_color)

                        for name, group in grouped:
                            # UA Plot (Primary Y) and Duty Plot (Tertiary Y)
                            ax1.plot(group['Mass Flow Rate (kg/hr)'], group['Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)'], color=temp_colors.get(name, 'k'), linestyle='-', marker='', label=f'UA @ {name}°C')
                            ax1.fill_between(group['Mass Flow Rate (kg/hr)'], group['Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)'], DESIGN_UA, where=(group['Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)'] > DESIGN_UA), color='red', alpha=0.3, interpolate=True)
                            ax3.plot(group['Mass Flow Rate (kg/hr)'], group['Heat Exchanger Duty (kcal/hr)'], color=temp_colors.get(name, 'k'), linestyle=':', marker='', linewidth=2, label=f'Duty @ {name}°C')
                            ax3.fill_between(group['Mass Flow Rate (kg/hr)'], group['Heat Exchanger Duty (kcal/hr)'], DESIGN_DUTY, where=(group['Heat Exchanger Duty (kcal/hr)'] > DESIGN_DUTY), color='red', alpha=0.3, interpolate=True)

                        # Design Lines
                        ax3.axhline(y=DESIGN_DUTY, color='purple', linestyle='--', linewidth=2.5, label=f'Design Duty ({DESIGN_DUTY} kcal/hr)')
                        ax1.axhline(y=DESIGN_UA, color='darkorange', linestyle='-.', linewidth=2.5, label=f'Design UA ({DESIGN_UA} kcal/hr.m².°C)')

                        # Legend
                        lines_ax1, labels_ax1 = ax1.get_legend_handles_labels()
                        lines_ax3, labels_ax3 = ax3.get_legend_handles_labels()
                        unique_labels_1 = {l: h for h, l in zip(lines_ax1 + lines_ax3, labels_ax1 + labels_ax3)}
                        fig1.tight_layout(rect=[0, 0.15, 0.95, 0.98])
                        ax1.legend(list(unique_labels_1.values()), list(unique_labels_1.keys()), loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=4, frameon=False, fontsize=9)
                        
                        # --- Insert Plot 1 into Excel ---
                        plot1_buf = io.BytesIO()
                        fig1.savefig(plot1_buf, format='png', bbox_inches='tight')
                        plot1_buf.seek(0)
                        
                        # Start Plot 1 insertion right after the data (e.g., cell B<data_end_row + 2>)
                        worksheet.insert_image(f'B{data_end_row + 2}', 'plot1.png', {'image_data': plot1_buf, 'x_scale': 0.7, 'y_scale': 0.7})
                        
                        # Display and close for Streamlit
                        st.pyplot(fig1)
                        plt.close(fig1)

                        
                        # ==============================================================================
                        # 3. PLOT 2 GENERATION (Fan Power and TS Outlet Temp)
                        # ==============================================================================
                        
                        fig2, ax2 = plt.subplots(figsize=(14, 8))
                        fig2.suptitle(f'Sheet: {sheet_name} | Performance Curve: Fan Power and TS Outlet Temperature vs. Mass Flow Rate', fontsize=16, fontweight='bold')

                        ax2.set_xlabel('Mass Flow Rate (kg/hr)', fontsize=12)
                        ax2.set_ylabel('Break Power/Fan (kW)', color='blue', fontsize=12)
                        ax2.tick_params(axis='y', labelcolor='blue')

                        ax4 = ax2.twinx()
                        temp_out_color = '#d62728'
                        ax4.spines['right'].set_color(temp_out_color)
                        ax4.set_ylabel('TS Outlet Temperature (Deg C)', color=temp_out_color, fontsize=12)
                        ax4.tick_params(axis='y', labelcolor=temp_out_color)

                        # Plotting Fan Power and TS Outlet Temp
                        for name, group in grouped:
                            ax2.plot(group['Mass Flow Rate (kg/hr)'], group['Break Power/Fan Summer (kW)'], color=temp_colors.get(name, 'k'), linestyle='-', marker='', label=f'Summer Power @ {name}°C')
                            ax2.plot(group['Mass Flow Rate (kg/hr)'], group['Break Power/Fan Winter (kW)'], color=temp_colors.get(name, 'k'), linestyle='--', marker='', label=f'Winter Power @ {name}°C')
                            ax4.plot(group['Mass Flow Rate (kg/hr)'], group['TS Outlet Temperature (Deg C)'], color=temp_colors.get(name, 'k'), linestyle='-', marker='.', linewidth=1.5, label=f'TS Outlet Temp @ {name}°C')
                            ax4.fill_between(group['Mass Flow Rate (kg/hr)'], group['TS Outlet Temperature (Deg C)'], DESIGN_TS_OUTLET_TEMP, where=(group['TS Outlet Temperature (Deg C)'] > DESIGN_TS_OUTLET_TEMP), color='red', alpha=0.3, interpolate=True)

                        # Design Lines and Limits
                        ax2.axhline(y=RATED_POWER, color='k', linestyle='-.', linewidth=2.5, label=f'Rated Power ({RATED_POWER} kW)')
                        ax4.axhline(y=DESIGN_TS_OUTLET_TEMP, color=temp_out_color, linestyle='--', linewidth=2.5, label=f'Design TS Outlet Temp ({DESIGN_TS_OUTLET_TEMP} °C)')
                        
                        min_power = df[['Break Power/Fan Summer (kW)', 'Break Power/Fan Winter (kW)']].min().min()
                        max_power = max(df[['Break Power/Fan Summer (kW)', 'Break Power/Fan Winter (kW)']].max().max(), RATED_POWER)
                        ax2.set_ylim(min_power * 0.9, max_power * 1.1)
                        min_temp = df['TS Outlet Temperature (Deg C)'].min()
                        max_temp = max(df['TS Outlet Temperature (Deg C)'].max(), DESIGN_TS_OUTLET_TEMP)
                        ax4.set_ylim(min_temp * 0.9, max_temp * 1.1)


                        # Legend
                        lines_ax2, labels_ax2 = ax2.get_legend_handles_labels()
                        lines_ax4, labels_ax4 = ax4.get_legend_handles_labels()
                        unique_labels_2 = {l: h for h, l in zip(lines_ax2 + lines_ax4, labels_ax2 + labels_ax4)}
                        fig2.tight_layout(rect=[0, 0.15, 0.95, 0.98])
                        ax2.legend(list(unique_labels_2.values()), list(unique_labels_2.keys()), loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=5, frameon=False, fontsize=8)

                        # --- Insert Plot 2 into Excel ---
                        plot2_buf = io.BytesIO()
                        fig2.savefig(plot2_buf, format='png', bbox_inches='tight')
                        plot2_buf.seek(0)
                        
                        # Start Plot 2 insertion below Plot 1 (approx 25 rows down from Plot 1 start)
                        start_row_2 = data_end_row + 27 
                        worksheet.insert_image(f'B{start_row_2}', 'plot2.png', {'image_data': plot2_buf, 'x_scale': 0.7, 'y_scale': 0.7})

                        # Display and close for Streamlit
                        st.pyplot(fig2)
                        plt.close(fig2)
                        
                        st.markdown("---") # Separator between sheet results

                except Exception as e:
                    st.error(f"An error occurred during processing or Excel generation: {str(e)}")
                    st.write("Please check your Excel file structure, sheet contents, and column names.")
                    st.stop() 

            st.success("✅ Excel report compilation complete. Download below.")
            
            # --- FINAL DOWNLOAD BUTTON FOR EXCEL ---
            output.seek(0)
            st.download_button(
                label="⬇️ Download Combined Multi-Sheet Performance Report (Excel)",
                data=output,
                file_name="Air_Cooler_Performance_Report_Combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Please upload an Excel file to begin the analysis.")
