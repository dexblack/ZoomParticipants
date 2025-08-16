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
        registrants_df = pd.read_excel(registrants_file, usecols=['Billing-Email', 'First Name', 'Last Name', 'Preferred Name']).fillna('')
        # Delegates file has no header
        delegates_df = pd.read_excel(delegates_file, header=None, names=['full_name', 'email']).fillna('')
        
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
    """Attempts to find a member group from the user_name string."""
    parts = user_name.lower().split()
    # Check for multi-word group names first (e.g., 'canada bay')
    if len(parts) >= 2:
        two_word_key = f"{parts[0]} {parts[1]}"
        if two_word_key in group_lookup:
            return group_lookup[two_word_key]
    # Check for single-word or initials
    if len(parts) >= 1:
        one_word_key = parts[0]
        if one_word_key in group_lookup:
            return group_lookup[one_word_key]
    return "Unknown"

def normalize_name(name):
    """Prepares a name string for comparison."""
    return re.sub(r'[^a-z0-9]', '', name.lower())

def analyze_attendance(participants_df, registrants_df, delegates_df, groups_list):
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
        
        # 1. Find Member Group
        local_group = find_best_group_match(zoom_user_name, group_lookup)

        # 2. Isolate the potential name from the user_name
        # This new block cleanly handles observers who don't provide a group name.
        potential_name = re.sub(r'\(.*?\)', '', zoom_user_name).strip() # Remove (pronouns) etc.
        if local_group != "Unknown":
            # If a group was found, attempt to remove it to isolate the person's name.
            # We use the cleaned group name for more reliable stripping.
            group_name_base = local_group.replace(" Greens", "")
            potential_name = re.sub(re.escape(group_name_base), '', potential_name, flags=re.IGNORECASE).strip()
        
        # 3. Match against Registrants
        # Attempt 1: Match by email (most reliable)
        if zoom_email:
            reg_match = registrants_df[registrants_df['Billing-Email'].str.lower().str.strip() == zoom_email]
            if not reg_match.empty:
                is_registered = True
                matched_email = reg_match.iloc[0]['Billing-Email']
                match_rules.append(f"Registered: Zoom email '{zoom_email}' == CiviCRM Billing-Email '{matched_email}'")

        # Attempt 2: Match by name (if email fails)
        if not is_registered and len(potential_name.split()) >= 2:
            norm_zoom_name = normalize_name(potential_name)
            for _, r_row in registrants_df.iterrows():
                # CiviCRM 'Sort Name' is often 'LastName, FirstName'
                sort_name = r_row['Last Name'] + ', ' + r_row['First Name']
                if ',' in sort_name:
                    last, first = [s.strip() for s in sort_name.split(',', 1)]
                    norm_reg_name = normalize_name(f"{first}{last}")
                    if norm_reg_name in norm_zoom_name or norm_zoom_name in norm_reg_name:
                         is_registered = True
                         match_rules.append(f"Registered: Zoom name '{potential_name}' ~ CiviCRM Sort Name '{sort_name}'")
                         break # Stop after first name match

        # 4. Match against Delegates
        # Attempt 1: Match by email
        if zoom_email:
            del_match = delegates_df[delegates_df['email'].str.lower().str.strip() == zoom_email]
            if not del_match.empty:
                is_delegate = True
                matched_email = del_match.iloc[0]['email']
                match_rules.append(f"Delegate: Zoom email '{zoom_email}' == Delegate email '{matched_email}'")
        
        # Attempt 2: Match by name (if email fails or doesn't exist)
        if not is_delegate and is_registered and len(potential_name.split()) >= 2: # Only try name-matching if they are at least registered
            norm_zoom_name = normalize_name(potential_name)
            for _, d_row in delegates_df.iterrows():
                delegate_name = d_row['full_name']
                norm_del_name = normalize_name(delegate_name)
                if norm_del_name in norm_zoom_name or norm_zoom_name in norm_del_name:
                    is_delegate = True
                    match_rules.append(f"Delegate: Zoom name '{potential_name}' ~ Delegate name '{delegate_name}'")
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
    
    analysis_results = analyze_attendance(participants, registrants, delegates, groups)
    
    write_output(analysis_results, args.output)

if __name__ == '__main__':
    main()