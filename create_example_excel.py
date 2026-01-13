"""
Script to create an example Excel file for the Capstone Defense Scheduler
This creates a working example with 8 projects, 10 panelists, and 12 time slots
"""

import pandas as pd
from datetime import datetime, timedelta

# Create example data
projects = pd.DataFrame({
    "project_id": ["P01", "P02", "P03", "P04", "P05", "P06", "P07", "P08"],
    "topic": ["NLP", "Finance", "ML", "NLP", "Finance", "ML", "NLP", "Finance"],
    "supervisor": ["Prof_A", "Prof_B", "Prof_C", "Prof_D", "Prof_E", "Prof_F", "Prof_G", "Prof_H"],
    "required_panelists": [2, 2, 2, 2, 2, 2, 2, 2]
})

panelists = pd.DataFrame({
    "panelist_id": ["Prof_A", "Prof_B", "Prof_C", "Prof_D", "Prof_E", 
                    "Prof_F", "Prof_G", "Prof_H", "Prof_I", "Prof_J"],
    "max_panels": [4, 4, 3, 3, 3, 3, 2, 2, 2, 2]
})

panelist_topics = pd.DataFrame({
    "panelist_id": ["Prof_A", "Prof_B", "Prof_C", "Prof_D", "Prof_E", 
                    "Prof_F", "Prof_G", "Prof_H", "Prof_I", "Prof_J"],
    "NLP": [1, 0, 0, 1, 0, 0, 1, 0, 1, 1],
    "Finance": [0, 1, 0, 0, 1, 0, 0, 1, 1, 0],
    "ML": [0, 0, 1, 1, 0, 1, 1, 0, 0, 1],
    "Data_Science": [1, 1, 1, 0, 0, 0, 0, 1, 1, 0]
})

# Create 12 time slots across 3 days
slots = pd.DataFrame({
    "slot_id": ["S01", "S02", "S03", "S04", "S05", "S06", "S07", "S08", "S09", "S10", "S11", "S12"],
    "date": ["2026-06-12", "2026-06-12", "2026-06-12", "2026-06-12",
             "2026-06-13", "2026-06-13", "2026-06-13", "2026-06-13",
             "2026-06-14", "2026-06-14", "2026-06-14", "2026-06-14"],
    "time": ["09-10", "10-11", "11-12", "14-15",
             "09-10", "10-11", "11-12", "14-15",
             "09-10", "10-11", "11-12", "14-15"],
    "room": ["R1", "R1", "R2", "R2", "R1", "R1", "R2", "R2", "R1", "R1", "R2", "R2"]
})

# Create availability matrix
availability = pd.DataFrame({
    "panelist_id": ["Prof_A", "Prof_B", "Prof_C", "Prof_D", "Prof_E", 
                    "Prof_F", "Prof_G", "Prof_H", "Prof_I", "Prof_J"],
    "S01": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "S02": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "S03": [1, 1, 0, 1, 1, 1, 1, 1, 1, 1],
    "S04": [1, 1, 1, 1, 1, 0, 1, 1, 1, 1],
    "S05": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "S06": [1, 0, 1, 1, 1, 1, 1, 1, 1, 1],
    "S07": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "S08": [1, 1, 1, 0, 1, 1, 1, 1, 1, 1],
    "S09": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "S10": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "S11": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "S12": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
})

# Create Excel file
output_file = "capstone_scheduler_example.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    projects.to_excel(writer, sheet_name='Projects', index=False)
    panelists.to_excel(writer, sheet_name='Panelists', index=False)
    panelist_topics.to_excel(writer, sheet_name='Panelist_Topics', index=False)
    slots.to_excel(writer, sheet_name='Time_Slots', index=False)
    availability.to_excel(writer, sheet_name='Availability', index=False)

print(f"✅ Example Excel file created: {output_file}")
print(f"\nFile contains:")
print(f"  - {len(projects)} projects")
print(f"  - {len(panelists)} panelists")
print(f"  - {len(slots)} time slots")
print(f"  - {len([c for c in panelist_topics.columns if c != 'panelist_id'])} topics")

