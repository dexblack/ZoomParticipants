import pandas as pd
import argparse
import re
import csv
import os

# --- Default File Names ---
DEFAULT_PARTICIPANTS_FILE = 'MeetingParticipants.csv'
DEFAULT_REGISTRANTS_FILE = 'CiviCRM_Participant_Search.xlsx'
DEFAULT_DELEGATES_FILE = 'delegates.xlsx'
DEFAULT_GROUPS_FILE = 'GNSW Local Groups.txt'
DEFAULT_OUTPUT_FILE = 'MeetingAttendance.csv'

def load_data(participants_file, registrants_file, delegates_file, groups_file):
    """Loads all input files into pandas DataFrames or lists."""
    try:
        participants_df = pd.read_csv(participants_file, usecols=['user_name', 'email']).fillna('')
        registrants_df = pd.read_excel(registrants_file, usecols=['Billing-Email', 'Local Group', 'First Name', 'Last Name', 'Preferred Name']).fillna('')
        # Delegates file has no header
        delegates_df = pd.read_excel(delegates_file, header=None, names=['local_group', 'full_name', 'email']).fillna('')
        
        with open(groups_file, 'r') as f:
            groups_list = [line.strip() for line in f if line.strip()]
            
        print("✅ All files loaded successfully.")
        return participants_df, registrants_df, delegates_df, groups_list
    except FileNotFoundError as e:
        print(f"❌ Error: File not found - {e.filename}. Please check the file path.")
        exit(1)
    except Exception as e:
        print(f"❌ An error occurred while loading files: {e}")
        exit(1)

def build_group_lookup(groups_list):
    """Creates a flexible lookup map for member group names."""
    lookup = {}
    for group in groups_list:
        # Cleaned name: 'Kiama Greens' -> 'kiama'
        cleaned = re.sub(r'\s+greens$', '', group, flags=re.IGNORECASE).strip().lower()
        lookup[cleaned] = group
        
        # Initials: 'Canada Bay Greens' -> 'cbg'
        words = cleaned.split()
        if len(words) > 1:
            initials = "".join(word[0] for word in words)
            lookup[initials] = group
            
    return lookup

def find_best_group_match(user_name, group_lookup):
    """Search for any group name within the user_name string."""
    clean_name = normalise_name(user_name)
    nospace_name = clean_name.replace(" ", "")
    
    for group_key, group_value in group_lookup.items():
        gk = group_key.lower()
        
        # First pass: strict word boundary
        pattern = r'\b' + re.escape(gk) + r'\b'
        if re.search(pattern, clean_name):
            return group_value
        
        # Second pass: ignore spaces (handles 'southernhighlands')
        if gk.replace(" ", "") in nospace_name:
            return group_value
    
    return "Unknown"

def normalise_name(name):
    """Lowercase and strip punctuation except spaces."""
    text = name.lower()
    text = re.sub(r'[^a-z\s]', ' ', text)  # keep only letters and spaces
    text = re.sub(r'(they|them|her|him|she|he)', '', text)
    return re.sub(r'\s+', ' ', text).strip()

def analyze_attendance(participants_df, registrants_df, delegates_df, groups_list, delegates_not_registered):
    """
    Performs the core analysis, matching participants against other lists.
    """
    group_lookup = build_group_lookup(groups_list)
    results = []

    for _, p_row in participants_df.iterrows():
        zoom_user_name = p_row['user_name']
        zoom_email = str(p_row['email']).lower().strip()
        
        # --- Initialize variables for this participant ---
        is_registered = False
        is_delegate = False
        match_rules = []
        
        # Initial group guess based on user_name
        local_group = find_best_group_match(zoom_user_name, group_lookup)

        # Match by name (if email fails)
        norm_zoom_name = re.sub(f"{local_group}", '', normalise_name(zoom_user_name))
        for _, r_row in registrants_df.iterrows():
            sort_name = r_row['Last Name'] + ', ' + r_row['First Name']
            if ',' in sort_name:
                last, first = [s.strip() for s in sort_name.split(',', 1)]
                norm_reg_name = normalise_name(f"{first}{last}")
                if norm_reg_name in norm_zoom_name or norm_zoom_name in norm_reg_name:
                    is_registered = True
                    match_rules.append(f"Registered: Zoom name '{norm_zoom_name}' ~ CiviCRM Sort Name '{sort_name}'")
                    # Use registrant’s Local Group if available
                    if pd.notna(r_row.get('Local Group', None)) and r_row['Local Group'].strip():
                        local_group = r_row['Local Group'].strip()
                        match_rules.append(f"Local Group refined from registrant record: '{local_group}'")
                    break

        # Match against Delegates
        norm_zoom_name = normalise_name(norm_zoom_name)
         # Match by name (if email fails or doesn't exist)
        if is_registered or not delegates_not_registered:
            for _, d_row in delegates_df.iterrows():
                delegate_name = normalise_name(d_row['full_name'])
                if delegate_name in norm_zoom_name or norm_zoom_name in delegate_name:
                    is_delegate = True
                    if 'Unknown' == local_group:
                        local_group = d_row['local_group']
                    match_rules.append(f"Delegate: Zoom name '{norm_zoom_name}' ~ Delegate name '{delegate_name}'")
                    break

        # --- Compile the final record for this participant ---
        results.append({
            'zoom_user_name': zoom_user_name,
            'email': p_row['email'],
            'local group': local_group,
            'registered/unregistered': 'Registered' if is_registered else 'Unregistered',
            'delegate/observer': 'Delegate' if is_delegate else 'Observer',
            'match rule': ' | '.join(match_rules) if match_rules else 'No Match Found'
        })
        
    return results

def write_output(results, output_file):
    """Writes the final analysis to a CSV file."""
    if not results:
        print("⚠️ No participants found to process.")
        return
        
    headers = [
        'zoom_user_name', 'email', 'local group', 
        'registered/unregistered', 'delegate/observer', 'match rule'
    ]
    
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            writer.writerows(results)
        print(f"\n✅ Analysis complete. Output saved to '{output_file}'")
    except Exception as e:
        print(f"❌ Error writing to output file: {e}")


def main():
    """Main function to parse arguments and run the analysis."""
    parser = argparse.ArgumentParser(description="Analyze Zoom attendance against registrant and delegate lists.")
    parser.add_argument(
        '--participants', 
        default=DEFAULT_PARTICIPANTS_FILE, 
        help=f"Path to the Zoom participants CSV file. (default: {DEFAULT_PARTICIPANTS_FILE})"
    )
    parser.add_argument(
        '--registrants', 
        default=DEFAULT_REGISTRANTS_FILE,
        help=f"Path to the registrants XLSX file. (default: {DEFAULT_REGISTRANTS_FILE})"
    )
    parser.add_argument(
        '--delegates', 
        default=DEFAULT_DELEGATES_FILE,
        help=f"Path to the delegates XLSX file. (default: {DEFAULT_DELEGATES_FILE})"
    )
    parser.add_argument(
        '--delegates_not_registered', 
        action='store_true',
        help="Assume Delegates are not Registered attendees."
    )
    parser.add_argument(
        '--groups', 
        default=DEFAULT_GROUPS_FILE,
        help=f"Path to the member groups TXT file. (default: {DEFAULT_GROUPS_FILE})"
    )
    parser.add_argument(
        '--output', 
        default=DEFAULT_OUTPUT_FILE,
        help=f"Path for the output CSV file. (default: {DEFAULT_OUTPUT_FILE})"
    )
    
    args = parser.parse_args()

    print("🚀 Starting attendance analysis...")
    print(f"Participants file: {os.path.abspath(args.participants)}")
    print(f"Registrants file:  {os.path.abspath(args.registrants)}")
    print(f"Delegates file:    {os.path.abspath(args.delegates)}")
    print(f"Groups file:       {os.path.abspath(args.groups)}")
    
    participants, registrants, delegates, groups = load_data(
        args.participants, args.registrants, args.delegates, args.groups
    )
    
    analysis_results = analyze_attendance(participants, registrants, delegates, groups, args.delegates_not_registered)
    
    write_output(analysis_results, args.output)

if __name__ == '__main__':
    main()