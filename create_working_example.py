"""
Create a working example Excel file for capstone defense scheduling.
This version is designed to be more feasible with better constraints.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta

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
        dates.append(current_date.strftime("%Y-%m-%d"))
    current_date += timedelta(days=1)

print(f"Generated {len(dates)} working days (excluding weekends)")

# Generate projects - reduce to 80 for better feasibility
num_projects = 80
projects_data = []

# Topics - reduce to 15 for better coverage
topics = [
    "NLP", "Finance", "ML", "Data_Science", "Cybersecurity",
    "Web_Development", "Mobile_Apps", "Cloud_Computing", "IoT", "Blockchain",
    "AI_Ethics", "Computer_Vision", "Robotics", "Game_Development", "Database_Systems"
]

# Supervisors - 25 supervisors
num_supervisors = 25
supervisors = [f"Supervisor_{i+1:02d}" for i in range(num_supervisors)]

for i in range(num_projects):
    project_id = f"P{i+1:03d}"
    topic = np.random.choice(topics)
    supervisor = np.random.choice(supervisors)
    # Most projects need 2 panelists, some need 3
    required_panelists = np.random.choice([2, 3], p=[0.8, 0.2])
    
    projects_data.append({
        "project_id": project_id,
        "topic": topic,
        "supervisor": supervisor,
        "required_panelists": required_panelists
    })

projects = pd.DataFrame(projects_data)

# Generate panelists - more panelists with higher capacity
num_panelists = 60  # Increased from 50
panelist_ids = supervisors.copy()

# Add additional panelists
for i in range(num_panelists - num_supervisors):
    panelist_ids.append(f"Prof_{chr(65+i%26)}{chr(65+(i//26)%26)}")

# Panelists with higher capacity
panelists_data = []
for i, panelist_id in enumerate(panelist_ids):
    # Supervisors can handle more panels
    if panelist_id in supervisors:
        max_panels = np.random.choice([6, 7, 8, 9], p=[0.2, 0.3, 0.3, 0.2])
    else:
        max_panels = np.random.choice([5, 6, 7], p=[0.3, 0.5, 0.2])
    
    panelists_data.append({
        "panelist_id": panelist_id,
        "max_panels": max_panels
    })

panelists = pd.DataFrame(panelists_data)

# Generate panelist expertise matrix
# Each panelist should have expertise in 4-8 topics (more coverage)
panelist_topics_data = []
for panelist_id in panelist_ids:
    row = {"panelist_id": panelist_id}
    # Each panelist has expertise in 4-8 random topics
    num_expertise = np.random.randint(4, 9)
    expert_topics = np.random.choice(topics, size=num_expertise, replace=False)
    
    for topic in topics:
        row[topic] = 1 if topic in expert_topics else 0
    
    panelist_topics_data.append(row)

panelist_topics = pd.DataFrame(panelist_topics_data)

# Generate time slots
# No rooms - each time slot is unique by date and time only
# Time slots from 9:30 to 17:30, 30 minutes each
slots_data = []
slot_id_counter = 1

# Generate 30-minute time slots from 9:30 to 17:30
times = []
start_hour = 9
start_minute = 30
end_hour = 17
end_minute = 30

current_hour = start_hour
current_minute = start_minute

while (current_hour < end_hour) or (current_hour == end_hour and current_minute <= end_minute):
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
    
    # Calculate how many fake panelists needed
    avg_fake_capacity = 5.5
    num_fake_needed = int(np.ceil(capacity_deficit / avg_fake_capacity)) + 2
    
    for i in range(num_fake_needed):
        fake_id = f"Fake_Panelist_{i+1:02d}"
        panelist_ids.append(fake_id)
        
        # Fake panelists have high capacity
        max_panels = np.random.choice([6, 7, 8], p=[0.3, 0.5, 0.2])
        panelists_data.append({
            "panelist_id": fake_id,
            "max_panels": max_panels
        })
        
        # Fake panelists have expertise in many topics
        num_expertise = np.random.randint(10, 16)  # 10-15 topics
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

# Generate availability matrix for all panelists
print("\nGenerating availability matrix...")
availability_data = []
slot_ids = slots['slot_id'].tolist()

for panelist_id in panelist_ids:
    row = {"panelist_id": panelist_id}
    
    # Fake panelists are always available (100%)
    if panelist_id.startswith("Fake_Panelist"):
        for slot_id in slot_ids:
            row[slot_id] = 1
    else:
        # Real panelists: Very high availability (98-100%) for better feasibility
        base_availability = np.random.uniform(0.98, 1.0)
        
        for slot_id in slot_ids:
            available = 1 if np.random.random() < base_availability else 0
            row[slot_id] = available
    
    availability_data.append(row)

availability = pd.DataFrame(availability_data)

# Ensure at least 15 panelists are always available in each slot
print("Ensuring minimum availability per slot...")
for slot_id in slot_ids:
    slot_availability = availability[slot_id].sum()
    if slot_availability < 15:
        needed = 15 - slot_availability
        unavailable = availability[availability[slot_id] == 0].index[:needed]
        for idx in unavailable:
            availability.loc[idx, slot_id] = 1

# Verify each project has enough eligible panelists
print("\nChecking project eligibility...")
infeasible = []
for _, project in projects.iterrows():
    project_id = project['project_id']
    topic = project['topic']
    supervisor = project['supervisor']
    required = int(project['required_panelists'])
    
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
            'project_id': project_id,
            'topic': topic,
            'required': required,
            'eligible': len(eligible)
        })

if infeasible:
    print(f"\n⚠️  {len(infeasible)} projects may have insufficient eligible panelists:")
    for item in infeasible[:10]:
        print(f"  {item['project_id']} ({item['topic']}): needs {item['required']}, has {item['eligible']}")
    
    # Fix by adding more expertise
    print("\nAttempting to fix by adding more expertise...")
    for item in infeasible:
        project = projects[projects.project_id == item['project_id']].iloc[0]
        topic = project['topic']
        supervisor = project['supervisor']
        required = item['required']
        
        missing_expertise = panelist_topics[
            (panelist_topics.panelist_id != supervisor) & 
            (panelist_topics[topic] == 0)
        ]
        
        needed = required - item['eligible']
        if len(missing_expertise) >= needed:
            add_to = missing_expertise.sample(n=needed)
            for idx in add_to.index:
                panelist_topics.loc[idx, topic] = 1
    
    print("✅ Fixed by adding additional expertise")
else:
    print("\n✅ All projects have sufficient eligible panelists!")

# Create Excel file
output_file = "capstone_scheduler_working_example.xlsx"

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

print(f"\n✅ Working example Excel file created: {output_file}")
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
print(f"  - Slot utilization: {(len(projects)/len(slots)*100):.1f}%")
if fake_count > 0:
    print(f"\n⚠️  Note: {fake_count} fake panelists were added to meet capacity requirements.")
    print(f"    These can be replaced with real panelists in your actual data.")

