# Capstone Defense Scheduler

A comprehensive optimization-based scheduling system for matching capstone defense projects with panelists and scheduling them into time slots. The system uses Mixed Integer Linear Programming (MILP) to ensure all constraints are satisfied while finding optimal or feasible solutions.

## 🚀 Features

- **Topic-Based Panelist Matching**: Groups panelists by expertise and matches them to projects
- **Intelligent Scheduling**: Optimizes defense scheduling into time slots with automatic room assignment
- **Constraint Satisfaction**: Handles multiple constraints including:
  - Topic compatibility (panelists must have expertise in project topic)
  - Supervisor exclusion (supervisors cannot be on their own project's panel)
  - Panelist capacity limits (respects maximum number of panels per panelist)
  - Panelist availability (no double booking in time slots)
  - Consecutive slot prevention (for 30-minute slots)
- **Automatic Room Assignment**: Computes and assigns rooms to concurrent defenses
- **Interactive Web Interface**: User-friendly Streamlit app with visualizations
- **Comprehensive Diagnostics**: Detailed error messages and feasibility checks

## 📋 Requirements

- Python 3.8+
- See `requirements.txt` for all dependencies

## 🛠️ Installation

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd app-panelists
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## 🎯 Quick Start

### Option 1: Streamlit Web App (Recommended)

1. **Run the app:**
   ```bash
   streamlit run app.py
   ```

2. **Open your browser:**
   - The app will automatically open at `http://localhost:8501`
   - Or navigate to the URL shown in the terminal

3. **Upload your data:**
   - Click "Upload Excel File" in the sidebar
   - Use one of the example files or create your own (see Data Format below)
   - Click "🚀 Run Scheduling Algorithm"

4. **View results:**
   - **Overview**: Data summary and statistics
   - **Calendar View**: Interactive calendar and Gantt chart
   - **Panel Assignments**: View assignments by topic and project
   - **Schedule Details**: Complete schedule table with room assignments
   - **Export Results**: Download results as Excel file

### Option 2: Python Module

```python
from capstone_scheduler import match_defenses_and_panelists, print_summary_report
import pandas as pd

# Load your data
excel_file = pd.ExcelFile('your_data.xlsx')
projects = pd.read_excel(excel_file, sheet_name='Projects')
panelists = pd.read_excel(excel_file, sheet_name='Panelists')
panelist_topics = pd.read_excel(excel_file, sheet_name='Panelist_Topics')
slots = pd.read_excel(excel_file, sheet_name='Time_Slots')
availability = pd.read_excel(excel_file, sheet_name='Availability')

# Run the algorithm
result = match_defenses_and_panelists(
    projects, panelists, panelist_topics, slots, availability
)

if result["success"]:
    print_summary_report(result, projects)
    print("\nSchedule with rooms:")
    print(result["schedule"])
else:
    print("Scheduling failed. Check diagnostics.")
```

### Option 3: Jupyter Notebook

Open `capstone_defense_scheduler.ipynb` and run the cells interactively.

## 📊 Data Format

Your Excel file should contain the following sheets:

### 1. Projects
- `project_id`: Unique identifier (e.g., "P001")
- `topic`: Topic/category (e.g., "NLP", "Finance", "ML")
- `supervisor`: Supervisor ID (cannot be assigned as panelist)
- `required_panelists`: Number of panelists needed (typically 2-3)

### 2. Panelists
- `panelist_id`: Unique identifier (e.g., "Prof_A")
- `max_panels`: Maximum number of panels this panelist can serve on

### 3. Panelist_Topics
- `panelist_id`: Unique identifier
- One column per topic with values:
  - `1`: Panelist has expertise in this topic
  - `0`: Panelist does not have expertise

### 4. Time_Slots
- `slot_id`: Unique identifier (e.g., "S001")
- `date`: Date in YYYY-MM-DD format (e.g., "2026-05-11")
- `time`: Time range (e.g., "09:30-10:00", "10:00-10:30")
- `room`: Optional - Rooms are automatically assigned by the algorithm

### 5. Availability
- `panelist_id`: Unique identifier
- One column per `slot_id` with values:
  - `1`: Panelist is available in this slot
  - `0`: Panelist is not available

## 📁 Project Structure

```
app-panelists/
├── app.py                          # Streamlit web application
├── capstone_scheduler.py           # Core algorithm module
├── capstone_defense_scheduler.ipynb # Jupyter notebook with examples
├── create_working_example.py       # Generate working example data
├── create_large_example.py          # Generate large example data
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── .gitignore                      # Git ignore rules
└── *.xlsx                          # Example Excel files
```

## 🔧 Algorithm Details

The algorithm uses **Mixed Integer Linear Programming (MILP)** with Google OR-Tools:

1. **Panelist Assignment Phase:**
   - Assigns panelists to projects based on expertise
   - Respects capacity constraints
   - Excludes supervisors from their own projects

2. **Scheduling Phase:**
   - Assigns projects to time slots
   - Ensures all panelists are available
   - Prevents double booking
   - Handles consecutive slot constraints (for 30-minute slots)

3. **Room Assignment Phase:**
   - Automatically computes rooms needed
   - Assigns rooms (R1, R2, R3...) to concurrent defenses
   - Ensures no conflicts for same date/time

## 🌐 Deployment to Streamlit Cloud

1. **Push to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin <your-github-repo-url>
   git push -u origin main
   ```

2. **Deploy on Streamlit Cloud:**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Sign in with GitHub
   - Click "New app"
   - Select your repository
   - Set main file path to: `app.py`
   - Click "Deploy"

3. **Your app will be live at:**
   `https://<your-username>-app-panelists.streamlit.app`

## 📝 Example Files

- `capstone_scheduler_working_example.xlsx`: 80 projects, optimized for feasibility
- `capstone_scheduler_large_example.xlsx`: 120 projects, larger dataset
- Use these as templates or test data

## 🐛 Troubleshooting

### Scheduling Fails

1. **Check diagnostics:** Expand "Algorithm Diagnostics" in the app
2. **Common issues:**
   - Not enough eligible panelists for a topic
   - Insufficient panelist capacity
   - Too many conflicts in availability
   - Not enough time slots

### Solutions

- Increase panelist availability (aim for 95%+)
- Add more panelists with needed expertise
- Increase `max_panels` for panelists
- Add more time slots or working days
- Use the working example generator: `python3 create_working_example.py`

## 📄 License

This project is provided as-is for educational and research purposes.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📧 Support

For issues or questions, please open an issue on GitHub.

---

**Note:** The algorithm automatically assigns rooms to scheduled defenses. You don't need to specify rooms in the input data - they will be computed based on concurrent defenses at the same date and time.
