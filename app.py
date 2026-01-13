import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
from capstone_scheduler import (
    match_defenses_and_panelists,
    print_summary_report,
    group_panelists_by_topics
)
import plotly.express as px
import plotly.graph_objects as go

# Page configuration
st.set_page_config(
    page_title="Capstone Defense Scheduler",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'result' not in st.session_state:
    st.session_state.result = None
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False

def load_data_from_excel(uploaded_file):
    """Load all required DataFrames from Excel file."""
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        required_sheets = {
            'projects': 'Projects',
            'panelists': 'Panelists',
            'panelist_topics': 'Panelist_Topics',
            'slots': 'Time_Slots',
            'availability': 'Availability'
        }
        
        data = {}
        for key, sheet_name in required_sheets.items():
            if sheet_name in sheet_names:
                data[key] = pd.read_excel(excel_file, sheet_name=sheet_name)
            else:
                st.error(f"❌ Missing sheet: {sheet_name}")
                return None
        
        return data
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None

def create_calendar_view(schedule_df, slots_df):
    """Create an interactive calendar view of the schedule."""
    if schedule_df.empty:
        return None
    
    # The schedule_df from the algorithm should already have date, time, room
    # But if it doesn't, merge with slots_df
    calendar = schedule_df.copy()
    
    if 'date' not in calendar.columns:
        # Merge schedule with slot details if date is missing
        slot_cols = ['slot_id', 'date', 'time']
        if 'room' in slots_df.columns:
            slot_cols.append('room')
        calendar = calendar.merge(
            slots_df[slot_cols],
            on='slot_id',
            how='left'
        )
    
    # Ensure date column exists and convert to datetime
    if 'date' in calendar.columns:
        calendar['date'] = pd.to_datetime(calendar['date'])
        # Create datetime for timeline
        if 'time' in calendar.columns:
            # Extract start time from time range (e.g., "09-10" -> "09:00")
            time_start = calendar['time'].str.split('-').str[0] + ':00'
            calendar['datetime'] = pd.to_datetime(
                calendar['date'].astype(str) + ' ' + time_start,
                format='%Y-%m-%d %H:%M',
                errors='coerce'
            )
    else:
        st.error("Date column not found in schedule data")
        return None
    
    return calendar

def export_results_to_excel(result, projects_df, slots_df):
    """Export results to Excel file."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Panel assignments
        result['panel_assignment'].to_excel(writer, sheet_name='Panel_Assignments', index=False)
        
        # Schedule
        slot_cols = ['slot_id', 'date', 'time']
        if 'room' in slots_df.columns:
            slot_cols.append('room')
        schedule_with_details = result['schedule'].merge(
            slots_df[slot_cols],
            on='slot_id',
            how='left'
        )
        schedule_with_details.to_excel(writer, sheet_name='Schedule', index=False)
        
        # Summary by project
        summary_data = []
        for _, row in result['schedule'].iterrows():
            project_id = row['project_id']
            project_info = projects_df[projects_df.project_id == project_id].iloc[0]
            panelists = result['panel_assignment'][
                result['panel_assignment'].project_id == project_id
            ]['panelist_id'].tolist()
            
            summary_row = {
                'Project_ID': project_id,
                'Topic': project_info['topic'],
                'Supervisor': project_info['supervisor'],
                'Date': row['date'],
                'Time': row['time'],
                'Panelists': ', '.join(panelists),
                'Number_of_Panelists': len(panelists)
            }
            if 'room' in row.index and pd.notna(row.get('room')):
                summary_row['Room'] = row['room']
            summary_data.append(summary_row)
        
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
        
        # Topic groups
        topic_groups_data = []
        for topic, panelists in result['topic_groups'].items():
            topic_groups_data.append({
                'Topic': topic,
                'Panelists': ', '.join(panelists),
                'Count': len(panelists)
            })
        pd.DataFrame(topic_groups_data).to_excel(writer, sheet_name='Topic_Groups', index=False)
    
    output.seek(0)
    return output

# Main App
st.markdown('<div class="main-header">📅 Capstone Defense Scheduler</div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("📤 Data Input")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xls'],
        help="Upload an Excel file with sheets: Projects, Panelists, Panelist_Topics, Time_Slots, Availability"
    )
    
    if uploaded_file is not None:
        data = load_data_from_excel(uploaded_file)
        
        if data:
            st.session_state.data_loaded = True
            st.success("✅ Data loaded successfully!")
            
            # Show data summary
            with st.expander("📊 Data Summary"):
                st.write(f"**Projects:** {len(data['projects'])}")
                st.write(f"**Panelists:** {len(data['panelists'])}")
                st.write(f"**Time Slots:** {len(data['slots'])}")
                
                # Calculate slots per day
                if 'date' in data['slots'].columns:
                    slots_per_day = data['slots'].groupby('date').size()
                    unique_dates = len(data['slots']['date'].unique())
                    avg_slots_per_day = slots_per_day.mean()
                    min_slots_per_day = slots_per_day.min()
                    max_slots_per_day = slots_per_day.max()
                    
                    st.write(f"**Working Days:** {unique_dates}")
                    st.write(f"**Slots per Day:**")
                    st.write(f"  - Average: {avg_slots_per_day:.1f}")
                    st.write(f"  - Min: {int(min_slots_per_day)}")
                    st.write(f"  - Max: {int(max_slots_per_day)}")
                
                st.write(f"**Topics:** {len([c for c in data['panelist_topics'].columns if c != 'panelist_id'])}")
            
            # Store data in session state
            st.session_state.projects = data['projects']
            st.session_state.panelists = data['panelists']
            st.session_state.panelist_topics = data['panelist_topics']
            st.session_state.slots = data['slots']
            st.session_state.availability = data['availability']
    
    st.markdown("---")
    st.header("⚙️ Settings")
    
    # Run algorithm button
    if st.session_state.data_loaded:
        if st.button("🚀 Run Scheduling Algorithm", type="primary", use_container_width=True):
            with st.spinner("Running optimization algorithm..."):
                # Capture stdout to show diagnostics
                import sys
                from io import StringIO
                
                # Redirect stdout to capture diagnostic messages
                old_stdout = sys.stdout
                sys.stdout = captured_output = StringIO()
                
                try:
                    result = match_defenses_and_panelists(
                        st.session_state.projects,
                        st.session_state.panelists,
                        st.session_state.panelist_topics,
                        st.session_state.slots,
                        st.session_state.availability
                    )
                    st.session_state.result = result
                finally:
                    # Restore stdout
                    sys.stdout = old_stdout
                    output = captured_output.getvalue()
                    
                    # Display diagnostics if any
                    if output:
                        with st.expander("📋 Algorithm Diagnostics", expanded=not result['success']):
                            st.text(output)
                
                if result['success']:
                    st.success("✅ Scheduling completed successfully!")
                else:
                    st.error("❌ Scheduling failed. Check constraints.")
                    st.info("💡 Expand 'Algorithm Diagnostics' above to see detailed error information.")

# Main content area
if not st.session_state.data_loaded:
    st.info("👈 Please upload an Excel file in the sidebar to begin.")
    
    # Show template structure
    with st.expander("📋 Excel File Template Structure"):
        st.markdown("""
        Your Excel file should contain the following sheets:
        
        1. **Projects** - Columns: `project_id`, `topic`, `supervisor`, `required_panelists`
        2. **Panelists** - Columns: `panelist_id`, `max_panels`
        3. **Panelist_Topics** - Columns: `panelist_id`, plus one column per topic (values: 1 or 0)
        4. **Time_Slots** - Columns: `slot_id`, `date`, `time` (room optional)
        5. **Availability** - Columns: `panelist_id`, plus one column per slot_id (values: 1 or 0)
        """)
        
        # Example template download
        st.markdown("### Download Template")
        template_data = {
            'Projects': pd.DataFrame({
                'project_id': ['P01', 'P02'],
                'topic': ['NLP', 'Finance'],
                'supervisor': ['Prof_A', 'Prof_B'],
                'required_panelists': [2, 2]
            }),
            'Panelists': pd.DataFrame({
                'panelist_id': ['Prof_A', 'Prof_B', 'Prof_C'],
                'max_panels': [3, 3, 2]
            }),
            'Panelist_Topics': pd.DataFrame({
                'panelist_id': ['Prof_A', 'Prof_B', 'Prof_C'],
                'NLP': [1, 0, 1],
                'Finance': [0, 1, 1]
            }),
            'Time_Slots': pd.DataFrame({
                'slot_id': ['S01', 'S02'],
                'date': ['2026-06-12', '2026-06-12'],
                'time': ['10-11', '11-12'],
                'room': ['R1', 'R1']
            }),
            'Availability': pd.DataFrame({
                'panelist_id': ['Prof_A', 'Prof_B', 'Prof_C'],
                'S01': [1, 1, 1],
                'S02': [1, 1, 1]
            })
        }
        
        template_output = io.BytesIO()
        with pd.ExcelWriter(template_output, engine='openpyxl') as writer:
            for sheet_name, df in template_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        template_output.seek(0)
        
        st.download_button(
            label="📥 Download Template Excel File",
            data=template_output,
            file_name="capstone_scheduler_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    # Tabs for different views
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Overview", 
        "📅 Calendar View", 
        "👥 Panel Assignments", 
        "📋 Schedule Details",
        "📤 Export Results"
    ])
    
    with tab1:
        st.header("Data Overview")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Projects", len(st.session_state.projects))
        with col2:
            st.metric("Panelists", len(st.session_state.panelists))
        with col3:
            st.metric("Time Slots", len(st.session_state.slots))
        with col4:
            topics = [c for c in st.session_state.panelist_topics.columns if c != 'panelist_id']
            st.metric("Topics", len(topics))
        
        # Projects table
        st.subheader("Projects")
        st.dataframe(st.session_state.projects, use_container_width=True)
        
        # Panelists table
        st.subheader("Panelists")
        st.dataframe(st.session_state.panelists, use_container_width=True)
        
        # Topic distribution chart
        st.subheader("Topic Distribution")
        topic_counts = st.session_state.projects['topic'].value_counts()
        fig = px.bar(
            x=topic_counts.index,
            y=topic_counts.values,
            labels={'x': 'Topic', 'y': 'Number of Projects'},
            title="Projects by Topic"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        st.header("📅 Calendar View")
        
        if st.session_state.result and st.session_state.result['success']:
            calendar = create_calendar_view(
                st.session_state.result['schedule'],
                st.session_state.slots
            )
            
            if calendar is not None:
                # Calendar visualization
                calendar['date_str'] = calendar['date'].dt.strftime('%Y-%m-%d')
                
                # Group by date
                st.subheader("Schedule by Date")
                
                dates = sorted(calendar['date_str'].unique())
                selected_date = st.selectbox("Select Date", dates)
                
                day_schedule = calendar[calendar['date_str'] == selected_date].sort_values('time')
                
                if not day_schedule.empty:
                    for _, row in day_schedule.iterrows():
                        with st.container():
                            col1, col2, col3 = st.columns([2, 2, 4])
                            with col1:
                                st.write(f"**{row['time']}**")
                            with col2:
                                if 'room' in row.index and pd.notna(row.get('room')):
                                    st.write(f"Room: {row['room']}")
                                else:
                                    st.write("Room: To be assigned")
                            with col3:
                                st.write(f"**{row['project_id']}** ({row['topic']})")
                                st.write(f"Panelists: {row['panelists']}")
                            st.divider()
                else:
                    st.info(f"No defenses scheduled on {selected_date}")
                
                # Gantt chart
                st.subheader("Gantt Chart View")
                fig = px.timeline(
                    calendar,
                    x_start='datetime',
                    x_end=calendar['datetime'] + pd.Timedelta(hours=1),
                    y='project_id',
                    color='topic',
                    labels={'project_id': 'Project', 'datetime': 'Time'},
                    title="Defense Schedule Timeline"
                )
                fig.update_layout(yaxis_autorange="reversed")
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Run the scheduling algorithm to see the calendar view.")
    
    with tab3:
        st.header("👥 Panel Assignments")
        
        if st.session_state.result and st.session_state.result['success']:
            # Group by topic
            topic_groups = st.session_state.result['topic_groups']
            
            st.subheader("Panelists by Topic")
            for topic, panelists_list in topic_groups.items():
                with st.expander(f"📚 {topic} ({len(panelists_list)} panelists)"):
                    st.write(", ".join(panelists_list))
            
            # Assignments by project
            st.subheader("Assignments by Project")
            assignments = st.session_state.result['panel_assignment']
            
            for project_id in st.session_state.projects['project_id']:
                project_assignments = assignments[assignments['project_id'] == project_id]
                if not project_assignments.empty:
                    project_info = st.session_state.projects[
                        st.session_state.projects['project_id'] == project_id
                    ].iloc[0]
                    
                    with st.expander(f"**{project_id}** - {project_info['topic']} (Supervisor: {project_info['supervisor']})"):
                        panelist_list = project_assignments['panelist_id'].tolist()
                        st.write("**Panelists:** " + ", ".join(panelist_list))
        else:
            st.info("Run the scheduling algorithm to see panel assignments.")
    
    with tab4:
        st.header("📋 Schedule Details")
        
        if st.session_state.result and st.session_state.result['success']:
            schedule = st.session_state.result['schedule'].copy()
            
            # Check if date column exists, if not merge with slots
            if 'date' not in schedule.columns:
                schedule = schedule.merge(
                    st.session_state.slots[['slot_id', 'date', 'time'] + (['room'] if 'room' in st.session_state.slots.columns else [])],
                    on='slot_id',
                    how='left'
                )
            
            # Sort by date and time if date exists
            if 'date' in schedule.columns:
                schedule['date'] = pd.to_datetime(schedule['date'])
                schedule = schedule.sort_values(['date', 'time'])
                
                # Format date for display
                schedule_display = schedule.copy()
                schedule_display['date'] = schedule_display['date'].dt.strftime('%Y-%m-%d')
                
                st.dataframe(
                    schedule_display[['date', 'time'] + (['room'] if 'room' in schedule_display.columns else []) + ['project_id', 'topic', 'panelists']],
                    use_container_width=True
                )
            else:
                # Fallback if date is missing
                st.dataframe(
                    schedule[['slot_id', 'project_id', 'topic', 'panelists']],
                    use_container_width=True
                )
                st.warning("Date information not available. Showing schedule by slot_id.")
            
            # Statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Scheduled", len(schedule))
            with col2:
                if 'date' in schedule.columns:
                    unique_dates = schedule['date'].nunique()
                    st.metric("Days", unique_dates)
                else:
                    st.metric("Days", "N/A")
            with col3:
                if 'room' in schedule.columns:
                    unique_rooms = schedule['room'].nunique()
                    st.metric("Rooms Used", unique_rooms)
                else:
                    st.metric("Rooms Used", "N/A")
        else:
            st.info("Run the scheduling algorithm to see schedule details.")
    
    with tab5:
        st.header("📤 Export Results")
        
        if st.session_state.result and st.session_state.result['success']:
            st.success("✅ Results are ready for export!")
            
            # Export to Excel
            excel_output = export_results_to_excel(
                st.session_state.result,
                st.session_state.projects,
                st.session_state.slots
            )
            
            st.download_button(
                label="📥 Download Results as Excel",
                data=excel_output,
                file_name=f"capstone_schedule_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Preview of what will be exported
            st.subheader("Export Preview")
            st.markdown("""
            The Excel file will contain the following sheets:
            - **Panel_Assignments**: Project-panelist assignments
            - **Schedule**: Complete schedule with dates, times, and rooms
            - **Summary**: Summary view with all key information
            - **Topic_Groups**: Panelists grouped by topic expertise
            """)
        else:
            st.info("Run the scheduling algorithm to export results.")

