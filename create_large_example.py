"""
Script to create a large example Excel file for the Capstone Defense Scheduler
120 defenses, 30 supervisors, 20 topics, dates from May 11-28 (excluding weekends)
"""

import pandas as pd
from datetime import datetime, timedelta
import numpy as np

# Set random seed for reproducibility
np.random.seed(42)

# Generate dates from May 11 to May 28, 2026, excluding weekends
start_date = datetime(2026, 5, 11)
end_date = datetime(2026, 5, 28)
dates = []
current_date = start_date
while current_date <= end_date:
    # Monday = 0, Sunday = 6
    if current_date.weekday() < 5:  # Monday to Friday
        dates.append(current_date.strftime('%Y-%m-%d'))
    current_date += timedelta(days=1)

print(f"Generated {len(dates)} working days (excluding weekends)")

# Topics
topics = [
    "NLP", "Finance", "ML", "Data_Science", "Cybersecurity",
    "Web_Development", "Mobile_Apps", "Cloud_Computing", "IoT",
    "Blockchain", "AI_Ethics", "Computer_Vision", "Robotics",
    "Game_Development", "Database_Systems", "Networking",
    "Software_Engineering", "Human_Computer_Interaction",
    "Distributed_Systems", "Quantum_Computing"
]

# Generate 120 projects
num_projects = 120
num_supervisors = 30

# Create supervisors
supervisors = [f"Prof_{chr(65+i//26)}{chr(65+i%26)}" if i >= 26 else f"Prof_{chr(65+i)}" 
               for i in range(num_supervisors)]

# Generate projects
projects_data = []
for i in range(num_projects):
    project_id = f"P{i+1:03d}"
    topic = np.random.choice(topics)
    supervisor = np.random.choice(supervisors)
    required_panelists = np.random.choice([2, 3], p=[0.7, 0.3])  # 70% need 2, 30% need 3
    
    projects_data.append({
        "project_id": project_id,
        "topic": topic,
        "supervisor": supervisor,
        "required_panelists": required_panelists
    })

projects = pd.DataFrame(projects_data)

# Generate panelists (more than supervisors to have enough capacity)
# Include supervisors as panelists, plus additional panelists
num_panelists = 50  # 30 supervisors + 20 additional
panelist_ids = supervisors.copy()

# Add additional panelists
for i in range(num_panelists - num_supervisors):
    panelist_ids.append(f"Prof_{chr(90-num_supervisors+i//26)}{chr(65+i%26)}" if i >= 26 else f"Prof_{chr(90-num_supervisors+i)}")

# Panelists with varying capacity
# Need enough capacity for ~280 panel slots (120 projects * ~2.3 avg panelists)
panelists_data = []
for i, panelist_id in enumerate(panelist_ids):
    # Supervisors can handle more panels
    if panelist_id in supervisors:
        max_panels = np.random.choice([5, 6, 7, 8], p=[0.2, 0.3, 0.3, 0.2])
    else:
        max_panels = np.random.choice([4, 5, 6], p=[0.3, 0.5, 0.2])
    
    panelists_data.append({
        "panelist_id": panelist_id,
        "max_panels": max_panels
    })

panelists = pd.DataFrame(panelists_data)

# Generate panelist expertise matrix
# Each panelist should have expertise in 3-6 topics
panelist_topics_data = []
for panelist_id in panelist_ids:
    row = {"panelist_id": panelist_id}
    # Each panelist has expertise in 3-6 random topics
    num_expertise = np.random.randint(3, 7)
    expert_topics = np.random.choice(topics, size=num_expertise, replace=False)
    
    for topic in topics:
        row[topic] = 1 if topic in expert_topics else 0
    
    panelist_topics_data.append(row)

panelist_topics = pd.DataFrame(panelist_topics_data)

# Generate time slots
# No rooms - each time slot is unique by date and time only
# Time slots from 9:30 to 17:30, 30 minutes each = 16 slots per day
slots_data = []
slot_id_counter = 1

# Generate 30-minute time slots from 9:30 to 17:30
times = []
start_hour = 9
start_minute = 30
end_hour_limit = 17
end_minute_limit = 30

current_hour = start_hour
current_minute = start_minute

while (current_hour < end_hour_limit) or (current_hour == end_hour_limit and current_minute < end_minute_limit):
    # Format start time
    start_str = f"{current_hour:02d}:{current_minute:02d}"
    
    # Calculate end time (30 minutes later)
    next_minute = current_minute + 30
    next_hour = current_hour
    if next_minute >= 60:
        next_minute -= 60
        next_hour += 1
    
    # Format end time
    end_str = f"{next_hour:02d}:{next_minute:02d}"
    
    # Add time slot
    times.append(f"{start_str}-{end_str}")
    
    # Move to next slot
    current_hour = next_hour
    current_minute = next_minute

for date in dates:
    for time in times:
        slots_data.append({
            "slot_id": f"S{slot_id_counter:03d}",
            "date": date,
            "time": time
        })
        slot_id_counter += 1

slots = pd.DataFrame(slots_data)
print(f"Generated {len(slots)} time slots")

# Note: Availability will be generated after fake panelists are added (if needed)

# Verify feasibility and add fake panelists if needed
print("\n" + "="*80)
print("FEASIBILITY CHECK")
print("="*80)

total_required = projects['required_panelists'].sum()
total_capacity = panelists['max_panels'].sum()

print(f"Total panelist slots needed: {total_required}")
print(f"Total panelist capacity: {total_capacity}")
print(f"Capacity ratio: {total_capacity/total_required:.2f}x")

# Add fake panelists if capacity is insufficient
if total_capacity < total_required:
    capacity_deficit = total_required - total_capacity
    print(f"\n⚠️  Capacity deficit: {capacity_deficit} slots")
    print("Adding fake panelists to meet capacity requirements...")
    
    # Calculate how many fake panelists needed (each with capacity 5-6)
    avg_fake_capacity = 5.5
    num_fake_needed = int(np.ceil(capacity_deficit / avg_fake_capacity)) + 2  # Add 2 extra for buffer
    
    fake_panelist_start = len(panelist_ids)
    for i in range(num_fake_needed):
        fake_id = f"Fake_Panelist_{i+1:02d}"
        panelist_ids.append(fake_id)
        
        # Fake panelists have high capacity
        max_panels = np.random.choice([5, 6, 7], p=[0.3, 0.5, 0.2])
        panelists_data.append({
            "panelist_id": fake_id,
            "max_panels": max_panels
        })
        
        # Fake panelists have expertise in many topics (to be flexible)
        num_expertise = np.random.randint(8, 15)  # 8-14 topics
        expert_topics = np.random.choice(topics, size=num_expertise, replace=False)
        
        row = {"panelist_id": fake_id}
        for topic in topics:
            row[topic] = 1 if topic in expert_topics else 0
        panelist_topics_data.append(row)
    
    print(f"✅ Added {num_fake_needed} fake panelists")
    
    # Update DataFrames
    panelists = pd.DataFrame(panelists_data)
    panelist_topics = pd.DataFrame(panelist_topics_data)
    
    # Recalculate capacity
    total_capacity = panelists['max_panels'].sum()
    print(f"New total capacity: {total_capacity}")
    print(f"New capacity ratio: {total_capacity/total_required:.2f}x")
else:
    print("\n✅ Capacity is sufficient, no fake panelists needed")

# Generate availability matrix for all panelists (including fake ones if added)
print("\nGenerating availability matrix...")
availability_data = []
slot_ids = slots['slot_id'].tolist()

for panelist_id in panelist_ids:
    row = {"panelist_id": panelist_id}
    
    # Fake panelists are always available (100%)
    if panelist_id.startswith("Fake_Panelist"):
        # All fake panelists available in all slots
        for slot_id in slot_ids:
            row[slot_id] = 1
    else:
        # Real panelists: Use pattern-based availability for better feasibility
        # Higher availability (95-100%) for large datasets to ensure feasibility
        base_availability = np.random.uniform(0.95, 1.0)
        
        # Create availability pattern: available most of the time with some gaps
        for slot_id in slot_ids:
            # Use base availability with small random variation
            available = 1 if np.random.random() < base_availability else 0
            row[slot_id] = available
    
    availability_data.append(row)

availability = pd.DataFrame(availability_data)

# Ensure at least some panelists are always available in each slot (for feasibility)
print("Ensuring minimum availability per slot...")
for slot_id in slot_ids:
    slot_availability = availability[slot_id].sum()
    # If less than 10 panelists available in a slot, make more available
    if slot_availability < 10:
        needed = 10 - slot_availability
        unavailable = availability[availability[slot_id] == 0].index[:needed]
        for idx in unavailable:
            availability.loc[idx, slot_id] = 1

# Check each project
infeasible = []
for _, project in projects.iterrows():
    topic = project['topic']
    supervisor = project['supervisor']
    required = project['required_panelists']
    
    # Find eligible panelists
    eligible = []
    for panelist_id in panelist_ids:
        if panelist_id == supervisor:
            continue
        expertise = panelist_topics.loc[
            panelist_topics.panelist_id == panelist_id, topic
        ].values[0]
        if expertise == 1:
            eligible.append(panelist_id)
    
    if len(eligible) < required:
        infeasible.append({
            'project_id': project['project_id'],
            'topic': topic,
            'required': required,
            'eligible': len(eligible)
        })

if infeasible:
    print(f"\n⚠️  {len(infeasible)} projects may have insufficient eligible panelists:")
    for item in infeasible[:10]:  # Show first 10
        print(f"  {item['project_id']} ({item['topic']}): needs {item['required']}, has {item['eligible']}")
    if len(infeasible) > 10:
        print(f"  ... and {len(infeasible) - 10} more")
    
    # Try to fix by adding more expertise
    print("\nAttempting to fix by adding more expertise...")
    for item in infeasible:
        project = projects[projects.project_id == item['project_id']].iloc[0]
        topic = project['topic']
        supervisor = project['supervisor']
        required = item['required']
        
        # Find panelists without this expertise
        missing_expertise = panelist_topics[
            (panelist_topics.panelist_id != supervisor) & 
            (panelist_topics[topic] == 0)
        ]
        
        # Add expertise to random panelists until we have enough
        needed = required - item['eligible']
        if len(missing_expertise) >= needed:
            add_to = missing_expertise.sample(n=needed)
            for idx in add_to.index:
                panelist_topics.loc[idx, topic] = 1
    
    print("✅ Fixed by adding additional expertise")
else:
    print("\n✅ All projects have sufficient eligible panelists!")

# Create Excel file
output_file = "capstone_scheduler_large_example.xlsx"

print(f"\nCreating Excel file: {output_file}")

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    projects.to_excel(writer, sheet_name='Projects', index=False)
    panelists.to_excel(writer, sheet_name='Panelists', index=False)
    panelist_topics.to_excel(writer, sheet_name='Panelist_Topics', index=False)
    slots.to_excel(writer, sheet_name='Time_Slots', index=False)
    availability.to_excel(writer, sheet_name='Availability', index=False)

# Count fake panelists
fake_panelists = [p for p in panelist_ids if p.startswith('Fake_Panelist')]
fake_count = len(fake_panelists)
real_count = len(panelists) - fake_count

print(f"\n✅ Large example Excel file created: {output_file}")
print(f"\nSummary:")
print(f"  - {len(projects)} projects")
print(f"  - {len(panelists)} panelists:")
print(f"    • {num_supervisors} supervisors")
print(f"    • {real_count - num_supervisors} additional real panelists")
if fake_count > 0:
    print(f"    • {fake_count} fake panelists (added for capacity)")
print(f"  - {len(topics)} topics")
print(f"  - {len(slots)} time slots across {len(dates)} working days")
print(f"  - Date range: {dates[0]} to {dates[-1]}")
print(f"  - Total capacity: {total_capacity} panel slots")
print(f"  - Total needed: {total_required} panel slots")
if fake_count > 0:
    print(f"\n⚠️  Note: {fake_count} fake panelists were added to meet capacity requirements.")
    print(f"    These can be replaced with real panelists in your actual data.")

