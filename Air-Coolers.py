import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import io
from PIL import Image as PILImage
import os

def process_sheet_data(df):
    """
    Process a single sheet's data and return the figure and analysis results.
    """
    try:
        # Data Cleaning and Column Renaming
        df.columns = [
            'Temperature_Inlet',
            'Flowrate_Actual',
            'Flowrate_PD_Limit',
            'Flowrate_Momentum_Limit'
        ]

        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df.dropna(inplace=True)

        # Constraint Analysis
        df['Flowrate_Constraint_Min'] = df[['Flowrate_Actual', 'Flowrate_PD_Limit', 'Flowrate_Momentum_Limit']].min(axis=1)

        active_constraints = np.argmin(df[['Flowrate_Actual', 'Flowrate_PD_Limit', 'Flowrate_Momentum_Limit']].values, axis=1)
        constraint_names = ['Actual Flowrate', 'Pressure Drop (PD)', 'Momentum']
        df['Active_Constraint'] = [constraint_names[i] for i in active_constraints]

        crossover_rows = df[df['Active_Constraint'] != df['Active_Constraint'].shift(1)]

        analysis_text = ""
        if len(crossover_rows) > 1:
            shift_data = crossover_rows.iloc[1]
            shift_temp = shift_data['Temperature_Inlet']
            shift_flow = shift_data['Flowrate_Constraint_Min']
            shift_to = shift_data['Active_Constraint']

            analysis_text = (
                f"Constraint Shift Analysis:\n"
                f"Limiting curve switches at Inlet Temperature: {shift_temp:.1f} °C\n"
                f"Flowrate limit at shift: {shift_flow:.0f} kg/hr\n"
                f"New limiting curve: {shift_to}"
            )
        else:
            shift_temp, shift_flow = None, None
            analysis_text = "Note: The overall limiting curve does not switch within the provided temperature range."

        # Create the plot
        fig, ax = plt.subplots(figsize=(14, 8))

        # Plot lines
        ax.plot(
            df['Temperature_Inlet'],
            df['Flowrate_Actual'],
            label='Area Ratio',
            color='#FF00FF',
            linestyle='-',
            linewidth=3
        )

        ax.plot(
            df['Temperature_Inlet'],
            df['Flowrate_PD_Limit'],
            label='Allowable Pressure Drop Limit (0.7 bar)',
            color='red',
            linestyle='-',
            linewidth=2.5
        )

        ax.plot(
            df['Temperature_Inlet'],
            df['Flowrate_Momentum_Limit'],
            label=r'Inlet Nozzle Momentum Limit ($7000-\rho v^2$)',
            color='#228B22',
            linestyle='-.',
            linewidth=2.5
        )

        # Fill safe operating zone
        ax.fill_between(
            df['Temperature_Inlet'],
            0,
            df['Flowrate_Constraint_Min'],
            color='#87CEFA',
            alpha=0.8,
            label='Safe Operating Zone (Below All Curves)'
        )

        # Annotate shift point
        if shift_temp and shift_flow:
            ax.scatter(
                shift_temp, shift_flow,
                color='black',
                marker='x',
                s=200,
                zorder=5,
                label='Limiting Curve Shift Point'
            )
            ax.annotate(
                f'Limiting Curve Shifts to {shift_to}\n@ {shift_temp:.1f}°C',
                xy=(shift_temp, shift_flow),
                xytext=(shift_temp + 5, shift_flow + 5000),
                arrowprops=dict(facecolor='black', shrink=0.05, width=1, headwidth=8),
                fontsize=10,
                bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.8)
            )

        # Plot aesthetics
        ax.set_title('ACHE Operating Envelope: Tube Side Mass Flowrate vs. Inlet Temperature',
                     fontsize=18, fontweight='bold')
        ax.set_xlabel('Tube Side Inlet Temperature (°C)', fontsize=14)
        ax.set_ylabel('Tube Side Mass Flowrate (kg/hr)', fontsize=14)
        ax.grid(True, linestyle=':', alpha=0.6)
        ax.legend(loc='upper right', fontsize=10)
        ax.set_ylim(bottom=0)

        plt.tight_layout()

        return fig, analysis_text

    except Exception as e:
        print(f"Error processing sheet: {e}")
        return None, f"Error: {str(e)}"


def process_excel_workbook(input_file, output_file=None):
    """
    Read all worksheets from an Excel file, generate plots, and save them back to Excel.

    Parameters:
    -----------
    input_file : str
        Path to the input Excel file containing data in multiple sheets
    output_file : str, optional
        Path to the output Excel file. If None, will create a file with '_output' suffix
    """
    
    # Fix the file extension if it's malformed
    # Get the base name without extension
    base_name = os.path.splitext(input_file)[0]
    
    # If file doesn't end with .xlsx or .xls, try to fix it
    if not (input_file.endswith('.xlsx') or input_file.endswith('.xls')):
        print(f"⚠️  Warning: File has unusual extension. Attempting to fix...")
        # Try to rename the file to .xlsx
        corrected_file = base_name + '.xlsx'
        try:
            os.rename(input_file, corrected_file)
            input_file = corrected_file
            print(f"✓ File renamed to: {corrected_file}")
        except Exception as e:
            print(f"✗ Could not rename file: {e}")
            print(f"Attempting to process anyway...")

    if output_file is None:
        # More robust handling of output filename
        if input_file.endswith('.xlsx'):
            output_file = input_file.replace('.xlsx', '_output.xlsx')
        elif input_file.endswith('.xls'):
            output_file = input_file.replace('.xls', '_output.xlsx')
        else:
            # Fallback for any other extension
            output_file = base_name + '_output.xlsx'

    try:
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(input_file)
        sheet_names = excel_file.sheet_names

        print(f"\n{'='*70}")
        print(f"Processing {len(sheet_names)} sheet(s) from: {input_file}")
        print(f"{'='*70}\n")

        # Create a new Excel writer
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

            for sheet_name in sheet_names:
                print(f"Processing sheet: '{sheet_name}'...")

                # Read the sheet
                df = pd.read_excel(input_file, sheet_name=sheet_name)

                # Process the data and create plot
                fig, analysis_text = process_sheet_data(df)

                if fig is not None:
                    # Save the original data to the output workbook
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    # Save the plot as an image in memory
                    img_buffer = io.BytesIO()
                    fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
                    img_buffer.seek(0)
                    plt.close(fig)

                    # Get the worksheet
                    worksheet = writer.sheets[sheet_name]

                    # Add the analysis text below the data
                    start_row = len(df) + 3
                    worksheet.cell(row=start_row, column=1, value="Analysis Results:")
                    for i, line in enumerate(analysis_text.split('\n')):
                        worksheet.cell(row=start_row + i + 1, column=1, value=line)

                    # Insert the plot image
                    img = Image(img_buffer)
                    # Position the image to the right of the data
                    img.anchor = f'F2'
                    worksheet.add_image(img)

                    print(f"  ✓ Successfully processed '{sheet_name}'")
                    print(f"  {analysis_text.split(chr(10))[0]}\n")
                else:
                    print(f"  ✗ Failed to process '{sheet_name}'\n")

        print(f"{'='*70}")
        print(f"✓ All sheets processed successfully!")
        print(f"Output saved to: {output_file}")
        print(f"{'='*70}\n")

        return output_file

    except Exception as e:
        print(f"\n✗ Error processing workbook: {e}")
        print(f"\n💡 Troubleshooting tips:")
        print(f"  1. Make sure the file is a valid Excel file (.xlsx or .xls)")
        print(f"  2. Check that the file is not corrupted")
        print(f"  3. Verify the file extension is correct")
        print(f"  4. Try opening the file in Excel and saving it again")
        raise


# ==============================================================================
# AUTO-EXECUTE WITH FILE UPLOAD
# ==============================================================================

print("\n" + "🚀 "*35)
print(" "*20 + "ACHE OPERATING ENVELOPE ANALYZER")
print("🚀 "*35 + "\n")

# Try Google Colab upload first
try:
    from google.colab import files
    print("="*70)
    print("📤 GOOGLE COLAB DETECTED")
    print("="*70)
    print("Please click 'Choose Files' below to upload your Excel file...")
    print("(File should contain multiple worksheets with ACHE data)")
    print("="*70 + "\n")

    uploaded = files.upload()

    if uploaded:
        filename = list(uploaded.keys())[0]
        print(f"\n✅ File '{filename}' uploaded successfully!\n")

        # Process the file
        output_file = process_excel_workbook(filename)

        # Download the result
        print("\n📥 Preparing download...")
        files.download(output_file)
        print("✅ Complete! Check your downloads folder.")

except ImportError:
    # Not in Colab, try Jupyter with ipywidgets
    try:
        import ipywidgets as widgets
        from IPython.display import display, clear_output

        print("="*70)
        print("📤 JUPYTER NOTEBOOK DETECTED")
        print("="*70)
        print("Click 'Upload' button below to select your Excel file...")
        print("="*70 + "\n")

        uploader = widgets.FileUpload(
            accept='.xlsx,.xls',
            multiple=False,
            description='📁 Upload Excel',
            button_style='success',
            icon='upload'
        )

        output_area = widgets.Output()

        def on_upload(change):
            with output_area:
                clear_output()
                if uploader.value:
                    uploaded_file = list(uploader.value.values())[0]
                    filename = list(uploader.value.keys())[0]

                    # Save the uploaded file
                    with open(filename, 'wb') as f:
                        f.write(uploaded_file['content'])

                    print(f"\n✅ File '{filename}' uploaded successfully!\n")

                    # Process the file
                    output_file = process_excel_workbook(filename)

                    print(f"\n✅ Processing complete!")
                    print(f"📥 Output file: {output_file}")
                    print(f"\nYou can download it from the file browser on the left →")

        uploader.observe(on_upload, names='value')

        # Display the upload widget
        display(widgets.VBox([
            widgets.HTML("<h3>📊 Upload Your Excel File</h3>"),
            uploader,
            output_area
        ]))

    except ImportError:
        # No interactive widgets available
        print("="*70)
        print("⚠️  INTERACTIVE UPLOAD NOT AVAILABLE")
        print("="*70)
        print("\n📋 Your platform doesn't support automatic file upload.")
        print("\nPlease follow these steps:\n")
        print("STEP 1: Upload your Excel file manually")
        print("  - Use the file browser/upload button in your environment")
        print("  - Or place the file in the same directory as this script\n")
        print("STEP 2: Run the following command:")
        print("  process_excel_workbook('your_filename.xlsx')\n")
        print("Example:")
        print("  process_excel_workbook('ACHE_Data.xlsx')\n")
        print("="*70)
        print("\n💡 TIP: In most platforms, you can drag & drop files into the")
        print("file browser panel to upload them.\n")
        print("="*70 + "\n")
