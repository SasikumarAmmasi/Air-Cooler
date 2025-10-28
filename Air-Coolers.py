import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import io

# ==============================================================================
# STREAMLIT APP CONFIGURATION
# ==============================================================================
st.set_page_config(page_title="Air Cooler Performance Analysis", layout="wide")
st.title("Air Cooler Performance Curve Analysis (Multi-Sheet Analysis)")

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
    
    if st.button("Generate Performance Curves for All Sheets"):
        with st.spinner("Processing data and generating plots..."):
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
                
                # Iterate over each sheet
                for sheet_name, df in all_sheets_data.items():
                    st.header(f"Results for Sheet: **{sheet_name}**")
                    
                    # Clean column names by stripping whitespace
                    df.columns = df.columns.str.strip()

                    # --- FIX COLUMN ALIGNMENT ISSUE ---
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

                    # Convert required columns to float, coercing errors
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

                    # Remove rows where key columns are NaN
                    df.dropna(subset=['Mass Flow Rate (kg/hr)', 'TS Inlet Temp (Deg C)', 'Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)', 'Heat Exchanger Duty (kcal/hr)', 'TS Outlet Temperature (Deg C)'], inplace=True)
                    
                    if df.empty:
                        st.warning(f"Sheet '{sheet_name}' is empty or contains no valid data after cleaning. Skipping plots for this sheet.")
                        st.markdown("---")
                        continue

                    # *** Make Heat Exchanger Duty positive (take absolute value) ***
                    df['Heat Exchanger Duty (kcal/hr)'] = df['Heat Exchanger Duty (kcal/hr)'].abs()

                    # Group data by the TS Inlet Temperature for distinct curves
                    grouped = df.groupby('TS Inlet Temp (Deg C)')

                    # ==============================================================================
                    # 2. PLOT GENERATION (Plot 1: UA and Duty)
                    # ==============================================================================

                    fig1, ax1 = plt.subplots(figsize=(14, 8))
                    # DYNAMIC TITLE
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
                        # UA Plot (Primary Y)
                        ax1.plot(
                            group['Mass Flow Rate (kg/hr)'],
                            group['Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)'],
                            color=temp_colors.get(name, 'k'),
                            linestyle='-',
                            marker='',
                            label=f'UA @ {name}°C'
                        )
                        # Shade area above Design UA
                        ax1.fill_between(
                            group['Mass Flow Rate (kg/hr)'],
                            group['Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)'],
                            DESIGN_UA,
                            where=(group['Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)'] > DESIGN_UA),
                            color='red',
                            alpha=0.3,
                            interpolate=True
                        )
                        # Duty Plot (Tertiary Y)
                        ax3.plot(
                            group['Mass Flow Rate (kg/hr)'],
                            group['Heat Exchanger Duty (kcal/hr)'],
                            color=temp_colors.get(name, 'k'),
                            linestyle=':',
                            marker='',
                            linewidth=2,
                            label=f'Duty @ {name}°C'
                        )
                        # Shade area above Design Duty
                        ax3.fill_between(
                            group['Mass Flow Rate (kg/hr)'],
                            group['Heat Exchanger Duty (kcal/hr)'],
                            DESIGN_DUTY,
                            where=(group['Heat Exchanger Duty (kcal/hr)'] > DESIGN_DUTY),
                            color='red',
                            alpha=0.3,
                            interpolate=True
                        )

                    # Add Design Air Cooler Duty line to Plot 1
                    ax3.axhline(
                        y=DESIGN_DUTY,
                        color='purple',
                        linestyle='--',
                        linewidth=2.5,
                        label=f'Design Duty ({DESIGN_DUTY} kcal/hr)'
                    )

                    # Add Design UA line to Plot 1
                    ax1.axhline(
                        y=DESIGN_UA,
                        color='darkorange',
                        linestyle='-.',
                        linewidth=2.5,
                        label=f'Design UA ({DESIGN_UA} kcal/hr.m².°C)'
                    )

                    # Combine legends for Plot 1
                    lines_ax1, labels_ax1 = ax1.get_legend_handles_labels()
                    lines_ax3, labels_ax3 = ax3.get_legend_handles_labels()

                    unique_labels_1 = {}
                    for h, l in zip(lines_ax1 + lines_ax3, labels_ax1 + labels_ax3):
                        unique_labels_1[l] = h

                    combined_lines_1 = list(unique_labels_1.values())
                    combined_labels_1 = list(unique_labels_1.keys())

                    fig1.tight_layout(rect=[0, 0.15, 0.95, 0.98])

                    ax1.legend(
                        combined_lines_1,
                        combined_labels_1,
                        loc='lower center',
                        bbox_to_anchor=(0.5, -0.25),
                        ncol=5,
                        frameon=False,
                        fontsize=9
                    )

                    # Display Plot 1 in Streamlit
                    st.pyplot(fig1)
                    
                    # Provide download button for Plot 1 (Dynamic filename)
                    buf1 = io.BytesIO()
                    fig1.savefig(buf1, format='png', bbox_inches='tight')
                    buf1.seek(0)
                    st.download_button(
                        label=f"Download UA & Duty Plot ({sheet_name})",
                        data=buf1,
                        file_name=f"Air_Cooler_Performance_Curve_UA_Duty_{sheet_name}.png",
                        mime="image/png"
                    )
                    
                    # 📢 CRITICAL FIX: Close the figure after displaying/saving
                    plt.close(fig1)

                    st.markdown("---")
                    
                    # ==============================================================================
                    # 2. PLOT GENERATION (Plot 2: Fan Power and TS Outlet Temp)
                    # ==============================================================================
                    
                    fig2, ax2 = plt.subplots(figsize=(14, 8))
                    # DYNAMIC TITLE
                    fig2.suptitle(f'Sheet: {sheet_name} | Performance Curve: Fan Power and TS Outlet Temperature vs. Mass Flow Rate', fontsize=16, fontweight='bold')

                    ax2.set_xlabel('Mass Flow Rate (kg/hr)', fontsize=12)
                    ax2.set_ylabel('Break Power/Fan (kW)', color='blue', fontsize=12)
                    ax2.tick_params(axis='y', labelcolor='blue')

                    # Create secondary y-axis for TS Outlet Temperature
                    ax4 = ax2.twinx()
                    temp_out_color = '#d62728'
                    ax4.spines['right'].set_color(temp_out_color)
                    ax4.set_ylabel('TS Outlet Temperature (Deg C)', color=temp_out_color, fontsize=12)
                    ax4.tick_params(axis='y', labelcolor=temp_out_color)

                    # Plotting Fan Power and TS Outlet Temp
                    for name, group in grouped:
                        # Summer Power (on ax2)
                        ax2.plot(
                            group['Mass Flow Rate (kg/hr)'],
                            group['Break Power/Fan Summer (kW)'],
                            color=temp_colors.get(name, 'k'),
                            linestyle='-',
                            marker='',
                            label=f'Summer Power @ {name}°C'
                        )
                        # Winter Power (on ax2)
                        ax2.plot(
                            group['Mass Flow Rate (kg/hr)'],
                            group['Break Power/Fan Winter (kW)'],
                            color=temp_colors.get(name, 'k'),
                            linestyle='--',
                            marker='',
                            label=f'Winter Power @ {name}°C'
                        )
                        
                        # TS Outlet Temperature (on ax4)
                        ax4.plot(
                            group['Mass Flow Rate (kg/hr)'],
                            group['TS Outlet Temperature (Deg C)'],
                            color=temp_colors.get(name, 'k'),
                            linestyle='-',
                            marker='',
                            linewidth=1.5,
                            label=f'TS Outlet Temp @ {name}°C'
                        )

                        # Shade area above Design TS Outlet Temp (on ax4)
                        ax4.fill_between(
                            group['Mass Flow Rate (kg/hr)'],
                            group['TS Outlet Temperature (Deg C)'],
                            DESIGN_TS_OUTLET_TEMP,
                            where=(group['TS Outlet Temperature (Deg C)'] > DESIGN_TS_OUTLET_TEMP),
                            color='red',
                            alpha=0.3,
                            interpolate=True
                        )

                    # Plot Fixed Rated Power (on ax2)
                    ax2.axhline(
                        y=RATED_POWER,
                        color='k',
                        linestyle='-.',
                        linewidth=2.5,
                        label=f'Rated Power ({RATED_POWER} kW)'
                    )

                    # Plot Design TS Outlet Temperature (on ax4)
                    ax4.axhline(
                        y=DESIGN_TS_OUTLET_TEMP,
                        color=temp_out_color, 
                        linestyle='--',
                        linewidth=2.5,
                        label=f'Design TS Outlet Temp ({DESIGN_TS_OUTLET_TEMP} °C)'
                    )

                    # Adjust AXIS 2 (Fan Power) limits
                    min_power = df[['Break Power/Fan Summer (kW)', 'Break Power/Fan Winter (kW)']].min().min()
                    max_power = max(df[['Break Power/Fan Summer (kW)', 'Break Power/Fan Winter (kW)']].max().max(), RATED_POWER)
                    ax2.set_ylim(min_power * 0.9, max_power * 1.1)

                    # Adjust AXIS 4 (TS Outlet Temp) limits
                    min_temp = df['TS Outlet Temperature (Deg C)'].min()
                    max_temp = max(df['TS Outlet Temperature (Deg C)'].max(), DESIGN_TS_OUTLET_TEMP)
                    ax4.set_ylim(min_temp * 0.9, max_temp * 1.1)


                    # Combine legends for Plot 2
                    lines_ax2, labels_ax2 = ax2.get_legend_handles_labels()
                    lines_ax4, labels_ax4 = ax4.get_legend_handles_labels()

                    unique_labels_2 = {}
                    for h, l in zip(lines_ax2 + lines_ax4, labels_ax2 + labels_ax4):
                        unique_labels_2[l] = h

                    combined_lines_2 = list(unique_labels_2.values())
                    combined_labels_2 = list(unique_labels_2.keys())
                    
                    fig2.tight_layout(rect=[0, 0.15, 0.95, 0.98])
                    ax2.legend(
                        combined_lines_2,
                        combined_labels_2,
                        loc='lower center',
                        bbox_to_anchor=(0.5, -0.25),
                        ncol=5,
                        frameon=False,
                        fontsize=8
                    )

                    # Display Plot 2 in Streamlit
                    st.pyplot(fig2)
                    
                    # Provide download button for Plot 2 (Dynamic filename)
                    buf2 = io.BytesIO()
                    fig2.savefig(buf2, format='png', bbox_inches='tight')
                    buf2.seek(0)
                    st.download_button(
                        label=f"Download Fan Power & Outlet Temp Plot ({sheet_name})",
                        data=buf2,
                        file_name=f"Air_Cooler_Performance_Curve_Fan_Power_Outlet_Temp_{sheet_name}.png",
                        mime="image/png"
                    )
                    
                    # 📢 CRITICAL FIX: Close the figure after displaying/saving
                    plt.close(fig2)
                    
                    st.markdown("---") # Separator between sheet results

                st.success("Performance curves generated successfully for all valid sheets!")

            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
                st.write("Please check your Excel file structure, sheet contents, and column names.")

else:
    st.info("Please upload an Excel file to begin the analysis.")
