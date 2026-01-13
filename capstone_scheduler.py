"""
Capstone Defense Scheduling Algorithm

This module implements a comprehensive algorithm to match capstone defenses with panelists
and schedule them into 1-hour calendar slots. The algorithm:

1. Groups panelists by topics - Ensures topic compatibility between projects and panelists
2. Assigns panelists to projects - Matches panelists based on expertise while respecting constraints
3. Schedules into calendar slots - Assigns defenses to available time slots

Constraints:
- Topic compatibility (panelists must have expertise in project topic)
- Supervisor exclusion (supervisors cannot be on their own project's panel)
- Panelist capacity limits (respects maximum number of panels per panelist)
- Panelist availability (no double booking in time slots)
- Each project gets exactly the required number of panelists
- Each project is scheduled exactly once
"""

import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp
from collections import defaultdict
from typing import Dict, List, Tuple, Optional


def group_panelists_by_topics(panelist_topics_df: pd.DataFrame) -> Dict[str, List[str]]:
    """
    Groups panelists by their topic expertise.
    
    Args:
        panelist_topics_df: DataFrame with panelist_id and topic columns (1/0 for expertise)
    
    Returns:
        Dictionary mapping topic -> list of panelist IDs who have expertise
    """
    topic_groups = defaultdict(list)
    
    # Get all topics (exclude panelist_id column)
    topics = [col for col in panelist_topics_df.columns if col != "panelist_id"]
    
    for _, row in panelist_topics_df.iterrows():
        panelist_id = row["panelist_id"]
        for topic in topics:
            if row[topic] == 1:  # Has expertise in this topic
                topic_groups[topic].append(panelist_id)
    
    return dict(topic_groups)


def assign_panelists_to_projects(
    projects_df: pd.DataFrame,
    panelists_df: pd.DataFrame,
    panelist_topics_df: pd.DataFrame
) -> Tuple[pd.DataFrame, bool]:
    """
    Assigns panelists to projects using MILP optimization.
    
    Args:
        projects_df: DataFrame with columns [project_id, topic, supervisor, required_panelists]
        panelists_df: DataFrame with columns [panelist_id, max_panels]
        panelist_topics_df: DataFrame with panelist_id and topic columns (1/0 for expertise)
    
    Returns:
        Tuple of (assignment DataFrame, success boolean)
    """
    solver = pywraplp.Solver.CreateSolver("SCIP")
    if not solver:
        print("SCIP solver not available, using default solver")
        solver = pywraplp.Solver.CreateSolver("CBC")
    
    # Decision variables: y[i, j] = 1 if panelist j is assigned to project i
    y = {}
    for i in projects_df.project_id:
        for j in panelists_df.panelist_id:
            y[i, j] = solver.BoolVar(f"y_{i}_{j}")
    
    # Constraint 1: Each project must have exactly the required number of panelists
    for i in projects_df.project_id:
        required = int(projects_df.loc[projects_df.project_id == i, "required_panelists"].values[0])
        solver.Add(sum(y[i, j] for j in panelists_df.panelist_id) == required)
    
    # Constraint 2: Topic compatibility - panelist must have expertise in project topic
    for _, p in projects_df.iterrows():
        project_id = p.project_id
        topic = p.topic
        for j in panelists_df.panelist_id:
            # Check if panelist has expertise in this topic
            expertise = panelist_topics_df.loc[
                panelist_topics_df.panelist_id == j, topic
            ].values
            if len(expertise) > 0:
                has_expertise = int(expertise[0])
                solver.Add(y[project_id, j] <= has_expertise)
    
    # Constraint 3: Supervisor cannot be on their own project's panel
    for _, p in projects_df.iterrows():
        supervisor = p.supervisor
        solver.Add(y[p.project_id, supervisor] == 0)
    
    # Constraint 4: Panelist capacity - respect max_panels limit
    for j in panelists_df.panelist_id:
        max_panels = int(panelists_df.loc[panelists_df.panelist_id == j, "max_panels"].values[0])
        solver.Add(sum(y[i, j] for i in projects_df.project_id) <= max_panels)
    
    # Solve
    status = solver.Solve()
    
    if status == pywraplp.Solver.OPTIMAL or status == pywraplp.Solver.FEASIBLE:
        # Extract assignments
        assignments = []
        for i in projects_df.project_id:
            for j in panelists_df.panelist_id:
                if y[i, j].solution_value() == 1:
                    assignments.append({
                        "project_id": i,
                        "panelist_id": j
                    })
        
        assignment_df = pd.DataFrame(assignments)
        return assignment_df, True
    else:
        print(f"Solver status: {status}")
        print("No feasible solution found for panel assignment!")
        print("\n" + "="*80)
        print("DIAGNOSTICS: Checking constraints...")
        print("="*80)
        
        # Diagnostic: Check if each project has enough eligible panelists
        for _, p in projects_df.iterrows():
            project_id = p.project_id
            topic = p.topic
            supervisor = p.supervisor
            required = int(p.required_panelists)
            
            # Find eligible panelists (have expertise AND not the supervisor)
            eligible = []
            for j in panelists_df.panelist_id:
                if j == supervisor:
                    continue
                expertise = panelist_topics_df.loc[
                    panelist_topics_df.panelist_id == j, topic
                ].values
                if len(expertise) > 0 and int(expertise[0]) == 1:
                    eligible.append(j)
            
            print(f"\nProject {project_id} ({topic}):")
            print(f"  Required panelists: {required}")
            print(f"  Supervisor: {supervisor} (excluded)")
            print(f"  Eligible panelists: {len(eligible)}")
            if len(eligible) > 0:
                print(f"    {', '.join(eligible)}")
            else:
                print(f"    ⚠️  NO ELIGIBLE PANELISTS!")
            if len(eligible) < required:
                print(f"  ❌ PROBLEM: Only {len(eligible)} eligible panelists but need {required}")
        
        # Diagnostic: Check panelist capacity
        print("\n" + "-"*80)
        print("Panelist Capacity Check:")
        print("-"*80)
        total_required = projects_df.required_panelists.sum()
        total_capacity = panelists_df.max_panels.sum()
        print(f"Total panelist slots needed: {total_required}")
        print(f"Total panelist capacity: {total_capacity}")
        if total_capacity < total_required:
            print(f"  ❌ PROBLEM: Total capacity ({total_capacity}) < Total needed ({total_required})")
        
        print("\n" + "="*80)
        print("SUGGESTIONS:")
        print("="*80)
        print("1. Check if projects have enough eligible panelists (expertise + not supervisor)")
        print("2. Increase panelist max_panels capacity if too low")
        print("3. Add more panelists with expertise in needed topics")
        print("4. Reduce required_panelists for some projects")
        print("="*80 + "\n")
        
        return pd.DataFrame(columns=["project_id", "panelist_id"]), False


def schedule_defenses(
    projects_df: pd.DataFrame,
    panel_assignment_df: pd.DataFrame,
    slots_df: pd.DataFrame,
    availability_df: pd.DataFrame
) -> Tuple[pd.DataFrame, bool]:
    """
    Schedules projects into time slots using MILP optimization.
    
    Args:
        projects_df: DataFrame with columns [project_id, topic, supervisor, required_panelists]
        panel_assignment_df: DataFrame with columns [project_id, panelist_id]
        slots_df: DataFrame with columns [slot_id, date, time] (room optional)
        availability_df: DataFrame with panelist_id and slot_id columns (1/0 for availability)
    
    Returns:
        Tuple of (schedule DataFrame, success boolean)
    """
    if len(panel_assignment_df) == 0:
        print("No panel assignments available. Run panel assignment first.")
        return pd.DataFrame(), False
    
    solver = pywraplp.Solver.CreateSolver("SCIP")
    if not solver:
        solver = pywraplp.Solver.CreateSolver("CBC")
    
    # Set time limit for large problems (300 seconds = 5 minutes)
    if len(projects_df) > 50:
        solver.SetTimeLimit(300000)  # 5 minutes in milliseconds
    
    # Decision variables: x[i, t] = 1 if project i is scheduled in slot t
    x = {}
    for i in projects_df.project_id:
        for t in slots_df.slot_id:
            x[i, t] = solver.BoolVar(f"x_{i}_{t}")
    
    # Constraint 1: Each project must be scheduled exactly once
    for i in projects_df.project_id:
        solver.Add(sum(x[i, t] for t in slots_df.slot_id) == 1)
    
    # Constraint 2: All assigned panelists must be available for the chosen slot
    # Pre-compute availability to speed up constraint generation
    availability_dict = {}
    for _, row in availability_df.iterrows():
        panelist_id = row["panelist_id"]
        availability_dict[panelist_id] = {}
        for slot_id in slots_df.slot_id:
            if slot_id in row.index:
                availability_dict[panelist_id][slot_id] = int(row[slot_id])
    
    for _, row in panel_assignment_df.iterrows():
        project_id = row["project_id"]
        panelist_id = row["panelist_id"]
        
        if panelist_id in availability_dict:
            for t in slots_df.slot_id:
                avail_value = availability_dict[panelist_id].get(t, 0)
                # If panelist is not available, project cannot be scheduled in this slot
                solver.Add(x[project_id, t] <= avail_value)
    
    # Constraint 3: No double booking - each panelist can only be in one slot at a time
    # Build a map of panelist -> assigned projects for efficiency
    panelist_projects_map = {}
    for _, row in panel_assignment_df.iterrows():
        panelist_id = row["panelist_id"]
        project_id = row["project_id"]
        if panelist_id not in panelist_projects_map:
            panelist_projects_map[panelist_id] = []
        panelist_projects_map[panelist_id].append(project_id)
    
    # Add constraints: each panelist can only be in one slot at a time
    for panelist_id in panelist_projects_map:
        assigned_projects = panelist_projects_map[panelist_id]
        for t in slots_df.slot_id:
            # At most one of the assigned projects can be scheduled in this slot
            solver.Add(
                sum(x[i, t] for i in assigned_projects) <= 1
            )
    
    # Constraint 4: Prevent consecutive slots for same panelist (defenses may take 1 hour)
    # Only apply if slots are 30 minutes or less (to avoid over-constraining)
    # Check slot duration by examining time format
    sample_time = slots_df.iloc[0]['time'] if len(slots_df) > 0 else ""
    slot_duration_30min = False
    if ':' in sample_time and '-' in sample_time:
        try:
            start_str, end_str = sample_time.split('-')
            start_h, start_m = map(int, start_str.split(':'))
            end_h, end_m = map(int, end_str.split(':'))
            duration_minutes = (end_h * 60 + end_m) - (start_h * 60 + start_m)
            if duration_minutes <= 30:
                slot_duration_30min = True
        except:
            pass
    
    if slot_duration_30min:
        print("Note: 30-minute slots detected. Adding consecutive slot constraints...")
        # Build consecutive slot pairs (same date, consecutive times, any room)
        consecutive_pairs = set()
        slots_by_datetime = {}
        for _, slot_row in slots_df.iterrows():
            slot_id = slot_row['slot_id']
            date = slot_row['date']
            time = slot_row['time']
            slots_by_datetime[slot_id] = (date, time)
        
        # Find consecutive slot pairs
        for t1 in slots_df.slot_id:
            date1, time1 = slots_by_datetime[t1]
            try:
                start_time_str = time1.split('-')[0]
                start_hour, start_minute = map(int, start_time_str.split(':'))
                # Calculate next slot time (30 minutes later)
                next_minute = start_minute + 30
                next_hour = start_hour
                if next_minute >= 60:
                    next_minute -= 60
                    next_hour += 1
                next_time_str = f"{next_hour:02d}:{next_minute:02d}"
                
                for t2 in slots_df.slot_id:
                    if t1 < t2:  # Avoid duplicates
                        date2, time2 = slots_by_datetime[t2]
                        if date2 == date1:
                            next_start_time = time2.split('-')[0]
                            if next_start_time == next_time_str:
                                consecutive_pairs.add((t1, t2))
            except (ValueError, IndexError):
                continue
        
        print(f"Found {len(consecutive_pairs)} consecutive slot pairs")
        
        # Add constraints: panelists can't be in consecutive slots
        constraint_count = 0
        for panelist_id in panelist_projects_map:
            assigned_projects = panelist_projects_map[panelist_id]
            for t1, t2 in consecutive_pairs:
                # Can't be in both consecutive slots
                solver.Add(
                    sum(x[i, t1] for i in assigned_projects) + 
                    sum(x[i, t2] for i in assigned_projects) <= 1
                )
                constraint_count += 1
        
        print(f"Added {constraint_count} consecutive slot constraints")
    else:
        print("Note: Slot duration > 30 minutes. Skipping consecutive slot constraints.")
    
    # Solve
    print(f"Solving scheduling problem: {len(projects_df)} projects, {len(slots_df)} slots...")
    status = solver.Solve()
    
    if status == pywraplp.Solver.OPTIMAL:
        print("✅ Optimal solution found!")
    elif status == pywraplp.Solver.FEASIBLE:
        print("✅ Feasible solution found (may not be optimal)")
    elif status == pywraplp.Solver.INFEASIBLE:
        print("❌ Problem is infeasible - no solution exists")
    elif status == pywraplp.Solver.NOT_SOLVED:
        print("⚠️  Solver did not complete (timeout or other issue)")
    
    if status == pywraplp.Solver.OPTIMAL or status == pywraplp.Solver.FEASIBLE:
        # Extract schedule
        schedule = []
        for i in projects_df.project_id:
            for t in slots_df.slot_id:
                if x[i, t].solution_value() == 1:
                    slot_info = slots_df[slots_df.slot_id == t].iloc[0]
                    # Get panelists for this project
                    project_panelists = panel_assignment_df[
                        panel_assignment_df.project_id == i
                    ]["panelist_id"].tolist()
                    
                    # Get project topic
                    project_topic = projects_df.loc[
                        projects_df.project_id == i, "topic"
                    ].values[0]
                    
                    schedule_entry = {
                        "slot_id": t,
                        "date": slot_info.date,
                        "time": slot_info.time,
                        "project_id": i,
                        "topic": project_topic,
                        "panelists": ", ".join(project_panelists),
                        "num_panelists": len(project_panelists)
                    }
                    # Add room only if it exists in the original data
                    if 'room' in slot_info.index:
                        schedule_entry["room"] = slot_info.room
                    schedule.append(schedule_entry)
        
        schedule_df = pd.DataFrame(schedule)
        
        # Compute and assign rooms to each scheduled project
        # Assign rooms based on concurrent defenses (same date and time)
        schedule_df = assign_rooms_to_schedule(schedule_df)
        
        return schedule_df, True
    else:
        print(f"Solver status: {status}")
        print("No feasible solution found for scheduling!")
        print("\n" + "="*80)
        print("DIAGNOSTICS: Checking scheduling constraints...")
        print("="*80)
        
        # Diagnostic: Check if we have enough slots
        num_projects = len(projects_df)
        num_slots = len(slots_df)
        print(f"\nSlot availability:")
        print(f"  Projects to schedule: {num_projects}")
        print(f"  Available time slots: {num_slots}")
        if num_slots < num_projects:
            print(f"  ❌ PROBLEM: Not enough slots ({num_slots}) for all projects ({num_projects})")
        else:
            print(f"  ✅ Sufficient slots available")
        
        # Diagnostic: Check panelist availability for each project
        print(f"\nPanelist availability check:")
        problematic_projects = []
        projects_with_slots = []
        
        # Pre-compute availability for efficiency
        availability_dict = {}
        for _, row in availability_df.iterrows():
            panelist_id = row["panelist_id"]
            availability_dict[panelist_id] = {}
            for slot_id in slots_df.slot_id:
                if slot_id in row.index:
                    availability_dict[panelist_id][slot_id] = int(row[slot_id])
        
        for _, project_row in projects_df.iterrows():
            project_id = project_row['project_id']
            assigned_panelists = panel_assignment_df[
                panel_assignment_df.project_id == project_id
            ]['panelist_id'].tolist()
            
            if len(assigned_panelists) == 0:
                problematic_projects.append({
                    'project_id': project_id,
                    'panelists': [],
                    'common_slots': 0,
                    'reason': 'No panelists assigned'
                })
                continue
            
            # Check if all panelists are available in at least one common slot
            common_slots = None
            for panelist_id in assigned_panelists:
                if panelist_id in availability_dict:
                    # Get slots where this panelist is available
                    panelist_avail_slots = [
                        slot_id for slot_id in slots_df.slot_id
                        if availability_dict[panelist_id].get(slot_id, 0) == 1
                    ]
                    
                    if common_slots is None:
                        common_slots = set(panelist_avail_slots)
                    else:
                        common_slots = common_slots.intersection(set(panelist_avail_slots))
                else:
                    common_slots = set()  # Panelist not in availability matrix
            
            if common_slots is None or len(common_slots) == 0:
                problematic_projects.append({
                    'project_id': project_id,
                    'panelists': assigned_panelists,
                    'common_slots': 0
                })
            else:
                projects_with_slots.append({
                    'project_id': project_id,
                    'common_slots': len(common_slots)
                })
        
        if problematic_projects:
            print(f"  ❌ {len(problematic_projects)} projects have no common available slots for their panelists:")
            for item in problematic_projects[:10]:  # Show first 10
                if 'reason' in item:
                    print(f"    • {item['project_id']}: {item['reason']}")
                else:
                    print(f"    • {item['project_id']}: panelists {item['panelists']} have no common slots")
            if len(problematic_projects) > 10:
                print(f"    ... and {len(problematic_projects) - 10} more")
        else:
            print(f"  ✅ All projects have at least one common slot for their panelists")
        
        if projects_with_slots:
            avg_slots = sum(p['common_slots'] for p in projects_with_slots) / len(projects_with_slots)
            print(f"  📊 Average common slots per project: {avg_slots:.1f}")
        
        # Diagnostic: Check slot capacity
        print(f"\nSlot capacity check:")
        slots_per_day = slots_df.groupby('date').size()
        unique_dates = len(slots_df['date'].unique())
        print(f"  Average slots per day: {slots_per_day.mean():.1f}")
        print(f"  Total days: {unique_dates}")
        print(f"  Projects per day needed: {num_projects / unique_dates:.1f}")
        print(f"  Max slots in a day: {slots_per_day.max()}")
        
        # Check room utilization (if rooms are specified)
        if 'room' in slots_df.columns:
            rooms = slots_df['room'].unique()
            print(f"  Number of rooms: {len(rooms)}")
            print(f"  Rooms: {', '.join(rooms)}")
        else:
            print(f"  Note: No room information (rooms will be assigned later)")
        
        # Check if we can fit all projects
        total_slot_capacity = len(slots_df)
        if total_slot_capacity < num_projects:
            print(f"  ❌ CRITICAL: Only {total_slot_capacity} slots but need {num_projects} projects!")
        else:
            utilization = (num_projects / total_slot_capacity) * 100
            print(f"  📊 Slot utilization: {utilization:.1f}% ({num_projects}/{total_slot_capacity})")
        
        print("\n" + "="*80)
        print("SUGGESTIONS:")
        print("="*80)
        if num_slots < num_projects:
            print("1. ❌ URGENT: Add more time slots (need at least " + str(num_projects) + " slots)")
        else:
            print("1. Increase number of time slots (more slots per day or more days)")
        print("2. Increase panelist availability (reduce conflicts, aim for 95%+ availability)")
        print("3. Add more rooms to allow parallel defenses")
        print("4. Reduce number of projects or spread over more days")
        if len(problematic_projects) > 0:
            print("5. Fix projects with no common slots (see list above)")
        print("="*80 + "\n")
        
        return pd.DataFrame(), False


def assign_rooms_to_schedule(schedule_df: pd.DataFrame) -> pd.DataFrame:
    """
    Assigns rooms to scheduled projects based on concurrent defenses.
    Projects scheduled at the same date and time get different rooms.
    
    Args:
        schedule_df: DataFrame with scheduled projects (must have 'date', 'time' columns)
    
    Returns:
        DataFrame with 'room' column added
    """
    if schedule_df.empty:
        return schedule_df
    
    # Group by date and time to find concurrent defenses
    schedule_df = schedule_df.copy()
    schedule_df['room'] = None
    
    # Group by date and time
    grouped = schedule_df.groupby(['date', 'time'])
    
    max_rooms_needed = 0
    for (date, time), group in grouped:
        num_concurrent = len(group)
        
        # Assign rooms: R1, R2, R3, etc.
        for idx, row_idx in enumerate(group.index):
            room_name = f"R{idx + 1}"
            schedule_df.loc[schedule_df.index == row_idx, 'room'] = room_name
        
        # Track maximum rooms needed
        if num_concurrent > max_rooms_needed:
            max_rooms_needed = num_concurrent
    
    if max_rooms_needed > 0:
        print(f"Room assignment: Maximum {max_rooms_needed} rooms needed for concurrent defenses")
    
    return schedule_df


def match_defenses_and_panelists(
    projects_df: pd.DataFrame,
    panelists_df: pd.DataFrame,
    panelist_topics_df: pd.DataFrame,
    slots_df: pd.DataFrame,
    availability_df: pd.DataFrame
) -> Dict:
    """
    Complete algorithm to match capstone defenses with panelists and schedule them.
    
    Args:
        projects_df: DataFrame with columns [project_id, topic, supervisor, required_panelists]
        panelists_df: DataFrame with columns [panelist_id, max_panels]
        panelist_topics_df: DataFrame with panelist_id and topic columns (1/0 for expertise)
        slots_df: DataFrame with columns [slot_id, date, time, room]
        availability_df: DataFrame with panelist_id and slot_id columns (1/0 for availability)
    
    Returns:
        Dictionary containing:
            - topic_groups: Dict grouping panelists by topics
            - panel_assignment: DataFrame of project-panelist assignments
            - schedule: DataFrame of scheduled defenses
            - success: Boolean indicating if algorithm succeeded
    """
    # Step 1: Group panelists by topics
    topic_groups = group_panelists_by_topics(panelist_topics_df)
    
    # Step 2: Assign panelists to projects
    panel_assignment, assign_success = assign_panelists_to_projects(
        projects_df, panelists_df, panelist_topics_df
    )
    
    # Step 3: Schedule defenses
    if assign_success:
        schedule, schedule_success = schedule_defenses(
            projects_df, panel_assignment, slots_df, availability_df
        )
    else:
        schedule = pd.DataFrame()
        schedule_success = False
    
    return {
        "topic_groups": topic_groups,
        "panel_assignment": panel_assignment,
        "schedule": schedule,
        "success": assign_success and schedule_success
    }


def print_summary_report(result: Dict, projects_df: pd.DataFrame):
    """
    Prints a summary report grouped by topics.
    
    Args:
        result: Result dictionary from match_defenses_and_panelists
        projects_df: DataFrame with project information
    """
    if not result["success"]:
        print("Cannot generate summary: algorithm did not complete successfully.")
        return
    
    topic_groups = result["topic_groups"]
    panel_assignment = result["panel_assignment"]
    schedule = result["schedule"]
    
    print("=" * 80)
    print("SUMMARY: PANELISTS GROUPED BY TOPICS")
    print("=" * 80)
    
    # Get all unique topics from projects
    all_topics = projects_df.topic.unique()
    
    for topic in all_topics:
        print(f"\n📚 Topic: {topic}")
        print("-" * 80)
        
        # Get panelists with this topic expertise
        topic_panelists = topic_groups.get(topic, [])
        print(f"Available Panelists: {', '.join(topic_panelists)}")
        
        # Get projects with this topic
        topic_projects = projects_df[projects_df.topic == topic]
        print(f"Projects: {', '.join(topic_projects.project_id.tolist())}")
        
        # Show assignments for this topic
        topic_assignments = panel_assignment[
            panel_assignment.project_id.isin(topic_projects.project_id)
        ]
        
        if len(topic_assignments) > 0:
            print("\nAssignments:")
            for _, row in topic_assignments.iterrows():
                project_id = row["project_id"]
                panelist_id = row["panelist_id"]
                # Get schedule info
                project_schedule = schedule[schedule.project_id == project_id]
                if len(project_schedule) > 0:
                    sched = project_schedule.iloc[0]
                    room_info = f" in {sched.room}" if 'room' in sched.index and pd.notna(sched.room) else ""
                    print(f"  • {project_id} → {panelist_id} (Scheduled: {sched.date} {sched.time}{room_info})")
                else:
                    print(f"  • {project_id} → {panelist_id}")
        
        # Show schedule for this topic
        topic_schedule = schedule[schedule.topic == topic]
        if len(topic_schedule) > 0:
            print("\nScheduled Defenses:")
            for _, row in topic_schedule.iterrows():
                room_info = f" | Room: {row.room}" if 'room' in row.index and pd.notna(row.room) else ""
                print(f"  • {row.date} {row.time}{room_info} | Project: {row.project_id}")
                print(f"    Panelists: {row.panelists}")
    
    print("\n" + "=" * 80)
    print("CALENDAR VIEW (Chronological)")
    print("=" * 80)
    
    # Sort by date and time
    schedule_sorted = schedule.sort_values(["date", "time"])
    for _, row in schedule_sorted.iterrows():
        room_info = f" | Room: {row.room}" if 'room' in row.index and pd.notna(row.room) else ""
        print(f"\n{row.date} {row.time}{room_info}")
        print(f"  Project: {row.project_id} ({row.topic})")
        print(f"  Panelists: {row.panelists}")


if __name__ == "__main__":
    # Example usage
    # Note: Make sure each project has enough eligible panelists (expertise + not supervisor)
    projects = pd.DataFrame({
        "project_id": ["P01", "P02", "P03"],
        "topic": ["NLP", "Finance", "NLP"],
        "supervisor": ["Prof_A", "Prof_B", "Prof_D"],
        "required_panelists": [2, 2, 2]
    })
    
    panelists = pd.DataFrame({
        "panelist_id": ["Prof_A", "Prof_B", "Prof_C", "Prof_D", "Prof_E", "Prof_F"],
        "max_panels": [3, 3, 2, 2, 2, 2]
    })
    
    # Updated: Added Prof_F with Finance expertise to make P02 feasible
    panelist_topics = pd.DataFrame({
        "panelist_id": ["Prof_A", "Prof_B", "Prof_C", "Prof_D", "Prof_E", "Prof_F"],
        "NLP": [1, 0, 1, 1, 1, 0],
        "Finance": [0, 1, 1, 0, 0, 1],  # Added Prof_F with Finance expertise
        "ML": [1, 1, 0, 1, 0, 0]
    })
    
    slots = pd.DataFrame({
        "slot_id": ["S01", "S02", "S03", "S04", "S05"],
        "date": ["2026-06-12", "2026-06-12", "2026-06-12", "2026-06-13", "2026-06-13"],
        "time": ["10-11", "11-12", "14-15", "10-11", "11-12"],
        "room": ["R1", "R1", "R2", "R1", "R2"]
    })
    
    availability = pd.DataFrame({
        "panelist_id": ["Prof_A", "Prof_B", "Prof_C", "Prof_D", "Prof_E", "Prof_F"],
        "S01": [1, 1, 1, 1, 0, 1],
        "S02": [0, 1, 1, 1, 1, 1],
        "S03": [1, 0, 1, 0, 1, 1],
        "S04": [1, 1, 0, 1, 1, 1],
        "S05": [1, 1, 1, 1, 1, 1]
    })
    
    print("Running complete algorithm...")
    result = match_defenses_and_panelists(
        projects, panelists, panelist_topics, slots, availability
    )
    
    if result["success"]:
        print("\n✅ Algorithm completed successfully!")
        print(f"\nTopic Groups: {len(result['topic_groups'])} topics")
        print(f"Panel Assignments: {len(result['panel_assignment'])} assignments")
        print(f"Scheduled Defenses: {len(result['schedule'])} defenses")
        print("\n")
        print_summary_report(result, projects)
    else:
        print("\n❌ Algorithm failed. Check constraints and data.")

