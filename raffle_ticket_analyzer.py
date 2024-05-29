import pandas as pd

# Load the file
file_path = r'path_to_file'  # Replace with the path to your file
sheet_name = 'Sheet1' # Replace if necessary

# Read the file
df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# Define the ticket price
ticket_price = 5

# Define discount rules
def apply_discount(ticket_count):
    free_tickets = 0
    if ticket_count >= 40:
        free_tickets = 25
    elif ticket_count >= 30:
        free_tickets = 15
    elif ticket_count >= 15:
        free_tickets = 7
    elif ticket_count >= 5:
        free_tickets = 2
    return ticket_count + free_tickets

# Function to extract the number of tickets from the "xn" format
def extract_ticket_count(ticket_str):
    if isinstance(ticket_str, str) and ticket_str.startswith('x'):
        try:
            return int(ticket_str[1:])
        except ValueError:
            return 0
    return 0

# Initialize variables
total_tickets = 0
total_amount = 0.0
i = 0
n = len(df)
people_data = []

# Iterate over the rows in the dataframe
while i < n:
    person_name = df.iloc[i, 0]  # Name of the person
    i += 1  # Move to the next row for ticket info
    person_tickets = 0
    
    # Check if we are not at the end of the dataframe and next row is not a person's name
    while i < n and isinstance(df.iloc[i, 0], str) and df.iloc[i, 0].startswith('x'):
        ticket_info = df.iloc[i, 0]
        person_tickets += extract_ticket_count(ticket_info)
        i += 1

    # Apply discount rules to calculate effective ticket count
    effective_tickets = apply_discount(person_tickets)
    
    # Calculate the amount to be paid for the actual tickets bought
    amount_spent = person_tickets * ticket_price
    total_amount += amount_spent

    # Add the effective tickets to the total tickets
    total_tickets += effective_tickets

    # Store data for each person
    people_data.append({
        'Name': person_name,
        'Tickets Bought': person_tickets,
        'Effective Tickets': effective_tickets,
        'Amount Spent (EUR)': amount_spent
    })

# Create a DataFrame from the collected data
people_df = pd.DataFrame(people_data)

# Calculate winning probability for each person
people_df['Winning Probability (%)'] = (people_df['Effective Tickets'] / total_tickets) * 100

# Add summary row to the DataFrame
summary_data = {
    'Name': 'Total',
    'Tickets Bought': people_df['Tickets Bought'].sum(),
    'Effective Tickets': total_tickets,
    'Amount Spent (EUR)': total_amount,
    'Winning Probability (%)': '' 
}

# Append the summary row
summary_df = pd.DataFrame([summary_data])
people_df = pd.concat([people_df, summary_df], ignore_index=True)

print(f"The total amount of tickets purchased is: {total_amount:.2f} EUR")
print(f"The total number of tickets (including free ones) is: {total_tickets}")

print("\nDetailed Information:")
print(people_df)

# Saving output to new Excel file
output_file_path = 'output_with_individual_details_and_summary.xlsx'
people_df.to_excel(output_file_path, index=False)



# Top 10 people based on winning probability

# Calculate winning probability for each person
people_df['Winning Probability (%)'] = (people_df['Effective Tickets'] / total_tickets) * 100

# Sort the DataFrame based on the winning probability in descending order
people_df_sorted = people_df.sort_values(by='Winning Probability (%)', ascending=False)

# List the top 10 people based on their probability of winning along with their respective tickets
top_10_people = people_df_sorted.head(10)

print("Top 10 People Based on Probability of Winning:")
print(top_10_people[['Name', 'Tickets Bought', 'Winning Probability (%)']])

# Saving output to new Excel file
output_file_path = 'output_top_10_people_with_tickets.xlsx'
top_10_people.to_excel(output_file_path, index=False)
