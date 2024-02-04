import streamlit as st
import json
from datetime import datetime, timedelta
from itertools import combinations
import random
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
from icalendar import Calendar, Event, vDatetime
import pytz

# CLASS LOGGER FUNCTIONS

# Function to display a single class entry
def display_class_entry(class_entry):
    st.write(f"Class Name: {class_entry['name'].capitalize()}")
    st.write(f"Group/Section: {class_entry['group']}")
    for schedule in class_entry['schedule']:
        st.write(f"Day: {schedule['day']}, Start Time: {schedule['start_time']}, End Time: {schedule['end_time']}")

# Function to save class data to a JSON file
def save_class_to_file(class_data, filename="classes.json"):
    try:
        # Read existing data from the file
        with open(filename, "r") as file:
            data = json.load(file)
    except FileNotFoundError:
        # If file does not exist, start with an empty list
        data = []

    # Append new class data
    data.append(class_data)

    # Write updated data back to the file
    with open(filename, "w") as file:
        json.dump(data, file, indent=4)

# ----------------------------

# TIMETABLE CREATOR FUNCTIONS

# Additional functions for timetable
def parse_time(time_str):
    return datetime.strptime(time_str, '%H:%M').time()

def times_overlap(start1, end1, start2, end2):
    return max(start1, start2) < min(end1, end2)

def has_conflict(class1, class2):
    for session1 in class1['schedule']:
        for session2 in class2['schedule']:
            if session1['day'] == session2['day']:
                if times_overlap(session1['start_time'], session1['end_time'],
                                 session2['start_time'], session2['end_time']):
                    return True
    return False

def has_unique_classes(combination):
    class_names = [cls['name'] for cls in combination]
    return len(class_names) == len(set(class_names))

def has_free_days(combination, free_days):
    # Checks if the combination respects the free days
    scheduled_days = {session['day'] for cls in combination for session in cls['schedule']}
    return all(day not in scheduled_days for day in free_days)

def get_unique_time_slots(classes):
    time_slots = set()
    for cls in classes:
        for session in cls['schedule']:
            time_slots.add((session['start_time'], session['end_time']))
    return sorted(time_slots, key=lambda x: x[0])

def get_random_light_color():
    return "{:02x}{:02x}{:02x}".format(random.randint(100, 255), random.randint(100, 255), random.randint(100, 255))

def create_single_sheet_xlsx_timetables(combinations, filename, time_slots, classes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Timetables"

    header_color = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
    class_colors = {cls['name']: PatternFill(start_color=get_random_light_color(), end_color=get_random_light_color(), fill_type="solid") for cls in classes}

    current_row = 1
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

    for i, combo in enumerate(combinations):
        current_row += 1  # Move to the next row for the header

        # Set the header row and column
        for col, day in enumerate(['Time'] + days, start=1):
            cell = ws.cell(row=current_row, column=col, value=day)
            if col != 1:  # Apply header color to days, not to the 'Time' column
                cell.fill = header_color

        # Create a dark grey border style
        dark_grey_side = Side(border_style="thin", color="404040")
        dark_grey_border = Border(left=dark_grey_side, right=dark_grey_side, top=dark_grey_side, bottom=dark_grey_side)

        # Apply the border to the header row
        for col in range(1, 7):  # Columns A-G
            ws.cell(row=current_row, column=col).border = dark_grey_border

        # Populate the timetable
        for time_slot in time_slots:
            current_row += 1
            ws.cell(row=current_row, column=1, value=f"{time_slot[0].strftime('%H:%M')} - {time_slot[1].strftime('%H:%M')}")
            ws.cell(row=current_row, column=1).fill = header_color

            for cls in combo:
                for session in cls['schedule']:
                    if (session['start_time'], session['end_time']) == time_slot:
                        day_col = days.index(session['day']) + 2
                        cell = ws.cell(row=current_row, column=day_col)
                        class_room_info = session.get('class_room', 'N/A')  # Retrieve class room or default to 'N/A'
                        cell.value = f"{cls['name']}\n(Group {cls['group']}\nRoom: {class_room_info})"
                        cell.fill = class_colors[cls['name']]
                        cell.alignment = Alignment(wrap_text=True)

        # Apply the border to each cell in the timetable
        for row in ws.iter_rows(min_row=current_row - len(time_slots), max_row=current_row, min_col=1, max_col=6):
            for cell in row:
                cell.border = dark_grey_border

        current_row += 5 

    # Set column widths
    for i, column_width in enumerate([20] + [25] * 5, start=1):
        ws.column_dimensions[get_column_letter(i)].width = column_width

    wb.save(filename)

# ----------------------------

# CALENDAR ICS GENERATOR FUNCTIONS
    
def generate_ics_file_for_classes(selected_classes, classes, start_date_str, end_date_str, filename="class_schedule.ics"):
    cal = Calendar()
    
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
    end_date = datetime.strptime(end_date_str, '%Y-%m-%d')

    for class_name in selected_classes:
        cls = next((item for item in classes if item["name"] == class_name), None)
        if cls is None:
            continue
        
        for session in cls['schedule']:
            day, start_time_str, end_time_str = session['day'], session['start_time'], session['end_time']
            if not all([day, start_time_str, end_time_str]):
                # Skip if any essential component is missing
                continue

            first_occurrence_date = start_date
            # Adjust first_occurrence_date to the first occurrence of the session day
            while first_occurrence_date.strftime('%A') != day:
                first_occurrence_date += timedelta(days=1)

            if first_occurrence_date > end_date:
                # If the first occurrence is beyond the semester end, skip this session
                continue

            start_datetime = datetime.combine(first_occurrence_date, datetime.strptime(start_time_str, '%H:%M').time())
            end_datetime = datetime.combine(first_occurrence_date, datetime.strptime(end_time_str, '%H:%M').time())

            event = Event()
            event.add('summary', f"{cls['name']} - Group {cls['group']}")
            event.add('location', session.get('class_room', 'N/A'))
            event.add('dtstart', start_datetime)
            event.add('dtend', end_datetime)
            event.add('rrule', {'freq': 'weekly', 'until': end_date})
            cal.add_component(event)

    with open(filename, 'wb') as f:
        f.write(cal.to_ical())
    
    return filename


# ----------------------------

# Main app function
def main():
    st.title("Class Schedule Management 🎓")
    st.markdown("""
        <style>
        /* Global styles */
        html, body, [class*="st-"] {
            font-family: 'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;
        }
        /* Style buttons for interactivity */
        .stButton>button {
            border: none;
            border-radius: px;
            padding: 10px 24px;
            margin: 5px 0px;
            color: white;
            background-color: #008CBA;
        }
        .stButton>button:hover {
            background-color: #005f73;
        }
        </style>
    """, unsafe_allow_html=True)

    # Credit for the creator
    st.markdown("Created by [Nicolas Cantarovici](https://www.linkedin.com/in/nicolas-cantarovici-3b85a0198)")
    st.markdown("""
    <div >
        <ol>
            <li><strong>Log Your Classes:</strong> Start by entering all classes you're interested in. Whether you're attending or just considering, log them to see all possible timetables.</li>
            <li><strong>Create Timetables:</strong> Navigate to the 'Timetable Creator' tab to generate custom timetables based on your classes. Experiment with different combinations to suit your schedule best.</li>
            <li><strong>Download & Share:</strong> Once satisfied, download your timetable and share it with friends or keep it for your reference.</li>
        </ol>
    </div>
""", unsafe_allow_html=True)


    # Create tabs for different functionalities
    tab1, tab2, tab3 = st.tabs(["📚 Class Logger", "⏱️ Timetable Creator", "🗓️ Add to Calendar"])

    with tab1:
        class_logger()

    with tab2:
        timetable_creator()

    with tab3:
        generate_ics_tab()

    

def class_logger():
    # Initialize session state for number of days
    if 'num_days' not in st.session_state:
        st.session_state.num_days = 1

    # Input for class name and group/section
    class_name = st.text_input("Class Name")
    group_section = st.text_input("Group/Section")

    # Input for number of days and dynamic schedule inputs
    num_days = st.number_input("Number of days per week", min_value=1, max_value=10, step=1, key='num_days')
    schedule_entries = []
    for i in range(num_days):
        cols = st.columns(4)
        with cols[0]:
            day = st.selectbox(f"Day {i+1}", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"], key=f"day{i}")
        with cols[1]:
            start_time = st.time_input(f"Start Time {i+1}", key=f"start_time{i}")
        with cols[2]:
            end_time = st.time_input(f"End Time {i+1}", key=f"end_time{i}")
        with cols[3]:
            class_room = st.text_input(f"Class Room {i+1}", key=f"class_room{i}")
        schedule_entries.append((day, start_time, end_time, class_room))

    # Submission button
    submit_button = st.button("Log Class")

    # Handling the submission
    if submit_button and class_name and group_section:
        schedule = [
            {
                "day": day, 
                "start_time": start_time.strftime("%H:%M"), 
                "end_time": end_time.strftime("%H:%M"),
                "class_room": class_room
            } for day, start_time, end_time, class_room in schedule_entries
        ]
        
        new_class = {
            "name": class_name.capitalize(),
            "group": group_section,
            "schedule": schedule
        }

        # Save to file
        save_class_to_file(new_class)

        # Display the logged class
        display_class_entry(new_class)
        st.success("Class logged and saved successfully.")

def timetable_creator():
    if not os.path.exists('classes.json'):
        st.warning("Please log a class first before generating a timetable.")
        return
    # Load and process classes
    with open('classes.json', 'r') as file:
        classes = json.load(file)

    parsed_classes = [{
        **cls, 
        'schedule': [{
            **session, 
            'start_time': parse_time(session['start_time']), 
            'end_time': parse_time(session['end_time'])
        } for session in cls['schedule']]
    } for cls in classes]

    # Get unique class names
    class_names = list(set(cls['name'] for cls in parsed_classes))

    # Add multiselect for mandatory classes
    mandatory_classes = st.multiselect("Select classes that must be included in every timetable", class_names)

    num_classes = st.number_input("Number of classes to attend", min_value=len(mandatory_classes), max_value=len(class_names), step=1)  

    # User input for selecting free days
    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    free_days = st.multiselect("Select days when you don't want to have classes", days_of_week)

    generate_button = st.button("Generate Timetable")

    if generate_button:
        try:
            class_combinations = combinations(parsed_classes, num_classes)
            viable_combinations = [
                combo for combo in class_combinations 
                if not any(has_conflict(cls1, cls2) for cls1, cls2 in combinations(combo, 2))
                and has_unique_classes(combo)
                and has_free_days(combo, free_days)
                and all(any(cls['name'] == mandatory_class for cls in combo) for mandatory_class in mandatory_classes)
            ]
            st.session_state['viable_combinations'] = viable_combinations

            if not viable_combinations:
                st.warning("No timetables found with the current criteria. Consider adjusting the number of classes, mandatory classes, or selected free days.")
                return

             # Extract unique time slots and pass them along with the combinations and classes
            time_slots = get_unique_time_slots(parsed_classes)
            filename = 'timetables.xlsx'
            create_single_sheet_xlsx_timetables(viable_combinations, filename, time_slots, classes)


            # Provide download link
            st.success("Timetable generated successfully.")
            with open(filename, "rb") as file:
                btn = st.download_button(
                        label="Download Timetable",
                        data=file,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"An error occurred: {e}")

def generate_ics_tab():
    if not os.path.exists('classes.json'):
        st.warning("No classes logged yet. Please log some classes first.")
        return
        
    with open('classes.json', 'r') as file:
        classes = json.load(file)

    class_options = [f"{cls['name']} - Group {cls['group']}" for cls in classes]
    selected_classes = st.multiselect("Select classes you will be attending to add to your calendar", class_options)

    start_date = st.date_input("Semester Start Date")
    end_date = st.date_input("Semester End Date")

    generate_button = st.button("Generate ICS File")

    if generate_button:
        if start_date and end_date:
            selected_class_names = [option.split(" - Group ")[0] for option in selected_classes]
            filename = generate_ics_file_for_classes(selected_class_names, classes, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
            with open(filename, "rb") as file:
                st.download_button(
                    label="Download ICS File",
                    data=file,
                    file_name=filename,
                    mime="text/calendar"
                )
        else:
            st.error("Please make sure both Semester Start and End Dates are selected.")



if __name__ == "__main__":
    main()
