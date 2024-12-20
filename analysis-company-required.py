import os
import pandas as pd
from collections import Counter

# Define expanded keyword lists
project_keywords = ['Project', 'Task', 'Milestone', 'Progress', 'Delivery', 'Implementation', 'Rollout', 'Launch',
                    'Completion', 'Achievement', 'Outcome']
client_keywords = ['Client', 'Customer', 'Service', 'Request', 'Inquiry', 'Support', 'Collaboration', 'Meeting',
                   'Solution', 'Feedback']
team_keywords = ['Meeting', 'Sync', 'Discussion', 'Update', 'Teamwork', 'Collaboration', 'Coordination', 'Briefing',
                 'Review', 'Planning', 'Strategy']
product_keywords = ['Feature', 'Enhancement', 'Improvement', 'Solution', 'Product', 'Service', 'Quality', 'Vendor',
                    'Supplier', 'Purchase', 'Procurement']

challenges_keywords = ['Issue', 'Problem', 'Blocker', 'Obstacle', 'Delay', 'Error', 'Defect', 'Failure', 'Concern',
                       'Pending', 'Setback', 'Bottleneck']
solutions_keywords = ['Resolution', 'Solution', 'Fix', 'Completed', 'Update', 'Repair', 'Adjusted', 'Improved',
                      'Repaired', 'Upgraded', 'Resolved', 'Optimized']

suggestions_keywords = ['Improve', 'Suggestion', 'Recommendation', 'Enhance', 'Feedback', 'Optimize', 'Upgrade',
                        'Efficiency', 'Streamline', 'Revise', 'Better', 'Solution', 'Action']
process_keywords = ['Process', 'Workflow', 'Automation', 'System', 'Structure', 'Efficiency', 'Policy',
                    'Standardization', 'Procedure']
tools_keywords = ['Tool', 'Platform', 'System', 'Software', 'Application', 'Integration', 'Technology',
                  'Infrastructure', 'Upgrade']

efficiency_keywords = ['Productivity', 'Performance', 'Efficiency', 'Streamline', 'Automate', 'Optimize', 'Minimize',
                       'Maximize', 'Focus', 'Output', 'Results']
teamwork_keywords = ['Collaboration', 'Communication', 'Team', 'Engagement', 'Coordination', 'Support', 'Alignment',
                     'Task', 'Role', 'Responsibility', 'Accountability']

processes_keywords = ['Standardization', 'Workflow', 'Policy', 'Documentation', 'Procedure', 'Guidelines',
                      'Best Practices', 'Governance', 'Compliance', 'Strategy', 'Control', 'Structure']
improvements_keywords = ['Refinement', 'Adjustments', 'Adaptation', 'Policy Change', 'Realignment', 'Restructure',
                         'Modification', 'Re-engineering', 'Transition']


# Function to read CSV file and select the appropriate one
def load_csv():
    # List all CSV files in the current directory
    files = [f for f in os.listdir() if f.endswith('.csv')]
    print("Available CSV files:")

    for idx, file in enumerate(files):
        print(f"{idx + 1}. {file}")

    # User selects a file
    file_index = int(input("Enter the number of the CSV file to analyze: ")) - 1
    selected_file = files[file_index]

    # Load the CSV file into a pandas dataframe
    df = pd.read_csv(selected_file)
    print(f"Loaded file: {selected_file}")
    return df


# Function to perform analysis on Key Achievements (Project Contributions)
def key_achievements(df):
    sender_count = df['SenderName'].value_counts().head(5)
    subject_counts = Counter()

    for subject in df['Subject'].dropna():
        for keyword in project_keywords + client_keywords + team_keywords + product_keywords:
            if keyword.lower() in subject.lower():
                subject_counts[keyword] += 1

    return sender_count, subject_counts


# Function to analyze Challenges and Solutions
def challenges_and_solutions(df):
    challenges = []
    solutions = []

    for subject in df['Subject'].dropna():
        if any(keyword.lower() in subject.lower() for keyword in challenges_keywords):
            challenges.append(subject)
        if any(keyword.lower() in subject.lower() for keyword in solutions_keywords):
            solutions.append(subject)

    return challenges, solutions


# Function to analyze Suggestions for Improvement
def suggestions_for_improvement(df):
    suggestions = []

    for subject in df['Subject'].dropna():
        if any(keyword.lower() in subject.lower() for keyword in
               suggestions_keywords + process_keywords + tools_keywords):
            suggestions.append(subject)

    return suggestions


# Function to analyze Team/Department Efficiency
def team_efficiency(df):
    efficiency = []

    for subject in df['Subject'].dropna():
        if any(keyword.lower() in subject.lower() for keyword in efficiency_keywords + teamwork_keywords):
            efficiency.append(subject)

    return efficiency


# Function to analyze Company Processes
def company_processes(df):
    processes = []

    for subject in df['Subject'].dropna():
        if any(keyword.lower() in subject.lower() for keyword in processes_keywords + improvements_keywords):
            processes.append(subject)

    return processes


# Main function to run the analysis
def main():
    # Load the data
    df = load_csv()

    # Perform the analysis
    sender_count, subject_counts = key_achievements(df)
    challenges, solutions = challenges_and_solutions(df)
    suggestions = suggestions_for_improvement(df)
    efficiency = team_efficiency(df)
    processes = company_processes(df)

    # Display the results
    print("\n--- Key Achievements (Project Contributions) ---")
    print(f"Top senders of emails: \n{sender_count}")
    print(f"Keyword occurrences in email subjects: \n{subject_counts}")

    print("\n--- Challenges and Solutions ---")
    print(f"Identified challenges: \n{challenges}")
    print(f"Suggested solutions: \n{solutions}")

    print("\n--- Suggestions for Improvement ---")
    print(f"Suggestions for improvement: \n{suggestions}")

    print("\n--- Team/Department Efficiency ---")
    print(f"Emails related to team efficiency: \n{efficiency}")

    print("\n--- Company Processes ---")
    print(f"Emails related to company processes: \n{processes}")


if __name__ == "__main__":
    main()
