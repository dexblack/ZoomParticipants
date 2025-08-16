import pandas as pd
import argparse
import os

# --- Default File Names ---
DEFAULT_INPUT_FILE = 'MeetingAttendance.csv'
DEFAULT_OUTPUT_FILE = 'MeetingSummary.xlsx'

def generate_summary_report(input_file, output_file):
    """
    Reads the annotated attendance CSV and generates a multi-sheet XLSX summary report.
    """
    # --- 1. Load the Data ---
    try:
        df = pd.read_csv(input_file)
        print(f"✅ Successfully loaded '{input_file}'.")
    except FileNotFoundError:
        print(f"❌ Error: Input file not found at '{os.path.abspath(input_file)}'.")
        print("Please run the 'attendance_analyzer.py' script first to generate this file.")
        return
    except Exception as e:
        print(f"❌ An error occurred while reading the file: {e}")
        return

    # --- 2. Perform Calculations ---
    total_attendees = len(df)
    unmatched_attendees = len(df[df['match rule'] == 'No Match Found'])
    unregistered_attendees = len(df[df['registered/unregistered'] == 'Unregistered'])

    # Quorum calculation: Find unique local groups represented by a delegate
    # Exclude any delegates whose group could not be identified ('Unknown')
    delegates_df = df[df['delegate/observer'] == 'Delegate']
    represented_groups_df = delegates_df[delegates_df['local group'] != 'Unknown']
    total_groups_with_delegates = represented_groups_df['local group'].nunique()

    # Create a summary DataFrame for the report sheet
    summary_data = {
        'Metric': [
            'Total Attendees (in Zoom)',
            'Total Groups with at least one Delegate',
            'Unmatched Attendees (for manual check)',
            'Unregistered Attendees (in Zoom but not registered)'
        ],
        'Value': [
            total_attendees,
            total_groups_with_delegates,
            unmatched_attendees,
            unregistered_attendees
        ]
    }
    summary_df = pd.DataFrame(summary_data)
    print("📊 Statistics calculated.")

    # --- 3. Prepare Filtered Lists for Other Sheets ---
    unmatched_df = df[df['match rule'] == 'No Match Found'].copy()
    unregistered_df = df[df['registered/unregistered'] == 'Unregistered'].copy()
    
    # Sort the delegate roster by local group then name for easy roll call
    delegate_roster_df = delegates_df.sort_values(by=['local group', 'zoom_user_name']).copy()

    # --- 4. Write to a Multi-Sheet XLSX File ---
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name='Summary Report', index=False)
            delegate_roster_df.to_excel(writer, sheet_name='Delegate Roster', index=False)
            unmatched_df.to_excel(writer, sheet_name='Unmatched Attendees', index=False)
            unregistered_df.to_excel(writer, sheet_name='Unregistered Attendees', index=False)
            df.to_excel(writer, sheet_name='Full Annotated Data', index=False)
        
        print(f"\n🎉 Success! Your meeting summary has been saved to '{output_file}'")
    except Exception as e:
        print(f"❌ An error occurred while writing the Excel file: {e}")

def main():
    """Main function to parse arguments and run the summary generation."""
    parser = argparse.ArgumentParser(
        description="Generate an XLSX summary report from the annotated attendance CSV."
    )
    parser.add_argument(
        '--input', 
        default=DEFAULT_INPUT_FILE, 
        help=f"Path to the input annotated attendance CSV file. (default: {DEFAULT_INPUT_FILE})"
    )
    parser.add_argument(
        '--output', 
        default=DEFAULT_OUTPUT_FILE,
        help=f"Path for the output XLSX summary file. (default: {DEFAULT_OUTPUT_FILE})"
    )
    
    args = parser.parse_args()
    
    generate_summary_report(args.input, args.output)

if __name__ == '__main__':
    main()