
# Class Schedule Management App

Effortlessly explore all possible timetable combinations, generating highly personalized and flexible class schedules.

## Features
- **Class Logging:** Input details of all potential classes (frequency, start and end times).
- **Timetable Generator:** Select classes to attend and set days off to generate custom timetables.
- **ICS Calendar Export:** Create a final timetable and download it as an ICS file for integration with calendar applications.

## Installation
Clone the repository and install dependencies:
```bash
git clone https://github.com/nicocanta20/Class-Schedule-Management
cd class-schedule-management-app
pip install -r requirements.txt
```

## Setup

**Generate MongoDB Credentials**

Set up a database on [MongoDB Atlas](https://www.mongodb.com/cloud/atlas). After creating your cluster, obtain your MongoDB user and password.

**Create a `.streamlit/secrets.toml` file**

In the root directory of the project, create a directory named `.streamlit` if it doesn't already exist. Inside this directory, create a file named `secrets.toml`. Add your MongoDB connection details to this file:

```toml
# .streamlit/secrets.toml
[MONGODB]
user = "your_mongodb_user"
password = "your_mongodb_password"
```

## Usage
To use the app, run the following command:
```
streamlit run streamlit_app.py
```
Then, access it in your web browser at `http://localhost:8501`.
