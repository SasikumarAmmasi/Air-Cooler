import pandas as pd
import matplotlib.pyplot as plt
import io
from google.colab import files

# ==============================================================================
# 1. INPUT DATA AND CONSTANTS
# ==============================================================================

# Upload the Excel file
uploaded = files.upload()

# Get the file path of the uploaded file
for fn in uploaded.keys():
  file_path = fn

# Get input data and constants from user
# file_path = input("Please enter the path to the Excel file (e.g., /content/A01-2601 Air Cooler Data - Input.xlsx): ")

# CONSTANT: The fixed Rated Power value (kW).
# ADJUST THIS VALUE IF YOUR RATED POWER IS DIFFERENT
# RATED_POWER = 30.0
RATED_POWER = float(input("Please enter the Rated Power value (kW): "))

# CONSTANT: Design Air Cooler Duty (kcal/hr)
# DESIGN_DUTY = 3350000
DESIGN_DUTY = float(input("Please enter the Design Air Cooler Duty (kcal/hr): "))

# CONSTANT: Design UA (kcal/hr.m².°C)
DESIGN_UA = float(input("Please enter the Design UA value (kcal/hr.m².°C): "))


# Load the data from the uploaded Excel file
df = pd.read_excel(file_path)


# Clean column names by stripping whitespace
df.columns = df.columns.str.strip()

# Print column names to help identify the correct name for 'Heat Exchanger Duty (kcal/hr)'
print("DataFrame columns after reading Excel and stripping whitespace:")
print(df.columns)

# --- FIX COLUMN ALIGNMENT ISSUE ---
# The columns were misaligned due to the comma in "Brake Power/Fan, Summer (kW)".
# We must rename the columns as they were actually parsed by pandas.
# Based on the provided Excel file content, the column names are expected to be:
# 'TS Gas Mass Flow (kg/hr)', 'TS Inlet Temperature (Deg C)', 'TS Outlet Temperature (Deg C)',
# 'Air Mass Flow (kg/h)', 'UA (kJ/C-h)', 'HE Duty (kcal/h)', # Corrected column name here
# 'Brake Power/Fan, Summer (kW)', 'Brake Power/Fan Winter (kW)'
# We need to map these to the desired internal names for consistency with the plotting code.
df = df.rename(columns={
    'TS Gas Mass Flow (kg/h)': 'Mass Flow Rate (kg/hr)', # Corrected column name here
    'TS Inlet Temperature (Deg C)': 'TS Inlet Temp (Deg C)',
    'UA (kcal/hr.m².°C)': 'Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)', # Corrected column name here
    'HE Duty (kcal/h)': 'Heat Exchanger Duty (kcal/hr)', # Corrected column name here
    'Brake Power/Fan, Summer (kW)': 'Break Power/Fan Summer (kW)',
    'Brake Power/Fan, Winter (kW)': 'Break Power/Fan Winter (kW)', # Corrected column name here
    'TS Outlet Temperature (Deg C)': 'TS Outlet Temperature (Deg C)',
    'Air Mass Flow (kg/h)': 'Air Mass Flow (kg/h)'
})


# Convert required columns to float, coercing errors
cols_to_convert = [
    'Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)',
    'Heat Exchanger Duty (kcal/hr)',
    'Break Power/Fan Summer (kW)',
    'Break Power/Fan Winter (kW)'
]
for col in cols_to_convert:
    # Ensure all spaces are correctly handled in the column names before conversion
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# Assuming 'Overall Heat Transfer Co-efficient (UA)' is already in kcal/hr.m².°C
# Rename the column to reflect the correct unit
# df = df.rename(columns={'Overall Heat Transfer Co-efficient (UA)': 'Overall Heat Transfer Co-efficient (UA) (kcal/hr.m².°C)'})


# *** CRUCIAL FIX: Make Heat Exchanger Duty positive (take absolute value) ***
df['Heat Exchanger Duty (kcal/hr)'] = df['Heat Exchanger Duty (kcal/hr)'].abs()

# Group data by the TS Inlet Temperature for distinct curves
grouped = df.groupby('TS Inlet Temp (Deg C)')

# ==============================================================================
# 2. PLOT GENERATION (Two Plots)
# ==============================================================================

# --- Color/Style Maps ---
temp_colors = {
    50.0: '#1f77b4',  # Blue
    55.0: '#ff7f0e',  # Orange
    60.0: '#2ca02c',  # Green
    65.0: '#d62728',  # Red
    70.0: '#9467bd',  # Purple
    75.0: '#8c564b'   # Brown
}

# --- Plot 1: Flow vs UA and Flow vs Duty ---
fig1, ax1 = plt.subplots(figsize=(14, 8))
fig1.suptitle('Air Cooler Performance Curve: UA and Heat Duty vs. Mass Flow Rate', fontsize=16, fontweight='bold')

ax1.set_xlabel('Mass Flow Rate (kg/hr)', fontsize=12)
ax1.set_ylabel('Service Overall Heat Transfer Coefficient (UA) (kcal/hr.m².°C)', color=temp_colors.get(list(temp_colors.keys())[0], 'k'), fontsize=12)
ax1.tick_params(axis='y', labelcolor=temp_colors.get(list(temp_colors.keys())[0], 'k'))

ax3 = ax1.twinx()
# ax3.spines['right'].set_position(('outward', 70)) # Removed the offset
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
        marker='',  # Removed marker
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
        marker='',  # Removed marker
        linewidth=2,  # Increased linewidth for duty curves
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
    color='purple',  # Choose a distinct color
    linestyle='--',
    linewidth=2.5,
    label=f'Design Duty ({DESIGN_DUTY} kcal/hr)'
)

# Add Design UA line to Plot 1
ax1.axhline(
    y=DESIGN_UA,
    color='darkorange',  # Choose a distinct color for UA
    linestyle='-.', # Different linestyle from Design Duty
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
    bbox_to_anchor=(0.5, -0.3),
    ncol=4,
    frameon=False,
    fontsize=9
)

# Save Plot 1
plt.savefig('Air_Cooler_Performance_Curve_UA_Duty.png')
print("Air_Cooler_Performance_Curve_UA_Duty.png saved to Colab session files.")


# --- Plot 2: Flow vs Fan Power ---
fig2, ax2 = plt.subplots(figsize=(14, 8))
fig2.suptitle('Air Cooler Performance Curve: Fan Power vs. Mass Flow Rate', fontsize=16, fontweight='bold')

ax2.set_xlabel('Mass Flow Rate (kg/hr)', fontsize=12)
power_color = 'red' # This will be overridden by temp_colors below
ax2.set_ylabel('Break Power/Fan (kW)', fontsize=12) # Removed color from ylabel
ax2.tick_params(axis='y') # Removed color from ytickparams

# Plotting Fan Power
for name, group in grouped:
    # Summer Power
    ax2.plot(
        group['Mass Flow Rate (kg/hr)'],
        group['Break Power/Fan Summer (kW)'],
        color=temp_colors.get(name, 'k'), # Use temperature colors
        linestyle='-',
        marker='', # Removed marker
        label=f'Summer Power @ {name}°C' # Add temperature to label
    )
    # Winter Power
    ax2.plot(
        group['Mass Flow Rate (kg/hr)'],
        group['Break Power/Fan Winter (kW)'],
        color=temp_colors.get(name, 'k'), # Use temperature colors
        linestyle='--',
        marker='', # Removed marker
        label=f'Winter Power @ {name}°C' # Add temperature to label
    )

# Plot Fixed Rated Power (Dark Dashes Line)
ax2.axhline(
    y=RATED_POWER,
    color='k',
    linestyle='-.',
    linewidth=2.5,
    label=f'Rated Power ({RATED_POWER} kW)'
)

# Adjust AXIS 2 (Fan Power) limits
min_power = df[['Break Power/Fan Summer (kW)', 'Break Power/Fan Winter (kW)']].min().min()
max_power = max(df[['Break Power/Fan Summer (kW)', 'Break Power/Fan Winter (kW)']].max().max(), RATED_POWER)
ax2.set_ylim(min_power * 0.9, max_power * 1.1)

# Combine legends for Plot 2
lines_ax2, labels_ax2 = ax2.get_legend_handles_labels()

# No need for unique_labels_2 dictionary, as we want all labels now
# unique_labels_2 = {}
# for h, l in zip(lines_ax2, labels_ax2):
#     if l != "_nolegend_":
#         unique_labels_2[l] = h

# combined_lines_2 = list(unique_labels_2.values())
# combined_labels_2 = list(unique_labels_2.keys())


fig2.tight_layout(rect=[0, 0.15, 0.95, 0.98])
ax2.legend(
    lines_ax2, # Use all lines and labels
    labels_ax2, # Use all lines and labels
    loc='lower center',
    bbox_to_anchor=(0.5, -0.25),
    ncol=5,
    frameon=False,
    fontsize=9
)


# Save Plot 2
plt.savefig('Air_Cooler_Performance_Curve_Fan_Power.png')
print("Air_Cooler_Performance_Curve_Fan_Power.png saved to Colab session files.")

plt.show()
