# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
from datetime import datetime
import folium
import os
import sys
from collections import Counter
# ------
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# ------
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash.development.base_component import Component

# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)
# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
current_file = os.path.basename(__file__)
script_dir = os.path.dirname(os.path.abspath(__file__))
# data_path = 'data/MarCom_Responses.xlsx'
# file_path = os.path.join(script_dir, data_path)
# data = pd.read_excel(file_path)
# df = data.copy()

# Define the Google Sheets URL
sheet_url = "https://docs.google.com/spreadsheets/d/1EFbKxXM_qBrD6PkxoYrOoZSnIfFpsaNY1NOIHrg_x0g/edit#gid=1782637761"

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials
encoded_key = os.getenv("GOOGLE_CREDENTIALS")

if encoded_key:
    json_key = json.loads(base64.b64decode(encoded_key).decode("utf-8"))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
else:
    creds_path = r"C:\Users\CxLos\OneDrive\Documents\BMHC\Data\bmhc-timesheet-4808d1347240.json"
    if os.path.exists(creds_path):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    else:
        raise FileNotFoundError("Service account JSON file not found and GOOGLE_CREDENTIALS is not set.")

# Authorize and load the sheet
client = gspread.authorize(creds)
sheet = client.open_by_url(sheet_url)
data = pd.DataFrame(client.open_by_url(sheet_url).sheet1.get_all_records())
df = data.copy()

# Strip whitespace
df.columns = df.columns.str.strip()
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Define a discrete color sequence
color_sequence = px.colors.qualitative.Plotly

# Filtered df where 'Date of Activity:' is in January
df['Date of Activity'] = pd.to_datetime(df['Date of Activity'], errors='coerce')
df = df[df['Date of Activity'].dt.month == 5]

# Get the reporting month:
current_month = datetime(2025, 5, 1).strftime("%B")
report_year = datetime(2025, 5, 1).strftime("%Y")
# -------------------------------------------------
# print(df)
# print(df[["Date of Activity", "Total travel time (minutes):"]])
# print('Total Marketing Events: ', len(df))
# print('Column Names: \n', df.columns.tolist())
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())
# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

columns =  [
    'Timestamp', 
    'Date of Activity', 
    'Person submitting this form:', 
    'Activity Duration (minutes):', 
    'Total travel time (minutes):',
    'What type of MARCOM activity are you reporting?', 
    'BMHC Activity:', 
    'Care Network Activity:', 
    'Brief activity description:', 
    'Activity Status', 
    'Community Outreach Activity:', 
    'Community Education Activity:', 
    'Any recent or planned changes to BMHC lead services or programs?', 
    'Entity Name:', 
    'Email Address'
]

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

#  Please provide public information:    137
# Please explain event-oriented:        13

# ============================== Data Preprocessing ========================== #

# Check for duplicate columns
# duplicate_columns = df.columns[df.columns.duplicated()].tolist()
# print(f"Duplicate columns found: {duplicate_columns}")
# if duplicate_columns:
#     print(f"Duplicate columns found: {duplicate_columns}")

# Rename columns
df.rename(
    columns={
        "What type of MARCOM activity are you reporting?": "MC Activity",
        "Activity Duration (minutes):": "Activity Duration",
        "Total travel time (minutes):": "Travel",
        "Person submitting this form:": "Person",
        "BMHC Activity:" : "BMHC Activity",
        "Care Network Activity:" : "Care Activity",
        "Community Outreach Activity:" : "Outreach Activity",
        "Community Education Activity:" : "Education Activity",
        "Activity Status" : "Activity Status",
        "Entity Name:" : "Entity",
    }, 
inplace=True)

# Fill Missing Values
# df['Please provide public information:'] = df['Please provide public information:'].fillna('N/A')
# df['Please explain event-oriented:'] = df['Please explain event-oriented:'].fillna('N/A')

# print(df.dtypes)

# ========================= Filtered DataFrames ========================== #

# -------------------------- MarCom Events --------------------------- #

marcom_events = len(df)
# print("Total Marcom events:", marcom_events)

# ---------------------------- MarCom Hours ---------------------------- #

# Remove the word 'hours' from the 'Activity duration:' column
df['Activity Duration'] = df['Activity Duration'].astype(str)  # Convert to string

df['Activity Duration'] = (
    df['Activity Duration']
    .str.replace(' hours', '', regex=False)
    .str.replace(' hour', '', regex=False)
)

df['Activity Duration'] = pd.to_numeric(df['Activity Duration'], errors='coerce')
# print('Column Names: \n', df.columns)

marcom_hours = df.groupby('Activity Duration').size().reset_index(name='Count')
marcom_hours = df['Activity Duration'].sum()/60
marcom_hours=round(marcom_hours)
# print('Total Activity Duration:', sum_activity_duration, 'hours')

# ------------------------- Travel Time ------------------------------ #

# print("Unique before:", df['Total travel time (minutes):'].unique().tolist())
# print("Travel time value counts", df['Total travel time (minutes):'].value_counts())

df["Travel"] = (
    df["Travel"]
    .replace('', 0)  # Replace empty strings
  
    .astype(float)   # Convert to float for summing in hours
)

df_mc_travel =df[["Date of Activity", "Travel"]]
# print(df_mc_travel.head())

mc_travel = df["Travel"].sum()/60
mc_travel = round(mc_travel)
# print("Total travel time:", mc_travel)

# --------------------------- MarCom Activity -------------------------- #

# Group by "Which MarCom activity category are you submitting an entry for?"
df_activities = df.groupby('MC Activity').size().reset_index(name='Count')
# print('Activities:\n', df_activities)

activity_bar=px.bar(
    df_activities,
    x='MC Activity',
    y='Count',
    color='MC Activity',
    text='Count',
).update_layout(
    height=600, 
    width=950,
    title=dict(
        text=f'{current_month} MarCom Activities',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,  # Rotate x-axis labels for better readability
        tickfont=dict(size=18),  # Adjust font size for the tick labels
        title=dict(
            # text=None,
            text="Name",
            font=dict(size=20),  # Font size for the title
        ),
        showticklabels=False  # Hide x-tick labels
        # showticklabels=True  # Hide x-tick labels
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  # Font size for the title
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        y=1,  # Position legend at the top
        xanchor="left",  # Anchor legend to the left
        yanchor="top",  # Anchor legend to the top
        # visible=False
        visible=True
    ),
    hovermode='closest', # Display only one hover label per trace
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Person Pie Chart
activity_pie=px.pie(
    df_activities,
    names="MC Activity",
    values='Count'  # Specify the values parameter
).update_layout(
    height=600,
    width=950,
    title=f'{current_month} Ratio of MarCom Activity',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=80,
    textposition='auto',
    insidetextorientation='horizontal', 
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Status</b>: %{value}<extra></extra>',
)

# -------------------------- Person Completing Form ------------------------- #

df['Person'] = (
    df['Person']
        .astype(str)
        .str.strip()
        .replace({
            'Felicia Chanlder' : 'Felicia Chandler',
        })
    )

df_person = df.groupby('Person').size().reset_index(name='Count')
# print(df_person.value_counts())

person_bar=px.bar(
    df_person,
    x='Person',
    y='Count',
    color='Person',
    text='Count',
).update_layout(
    height=440, 
    width=780,
    title=dict(
        text='People Submitting Forms',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=0,  # Rotate x-axis labels for better readability
        tickfont=dict(size=18),  # Adjust font size for the tick labels
        title=dict(
            text=None,
            # text="Name",
            font=dict(size=20),  # Font size for the title
        ),
        showticklabels=False  # Hide x-tick labels
        # showticklabels=True  # Hide x-tick labels
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  # Font size for the title
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        y=1,  # Position legend at the top
        xanchor="left",  # Anchor legend to the left
        yanchor="top",  # Anchor legend to the top
        # visible=False
        visible=True
    ),
    margin=dict(t=60, r=0, b=30, l=150),
    hovermode='closest', # Display only one hover label per trace
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Person Pie Chart
person_pie=px.pie(
    df_person,
    names="Person",
    values='Count'  # Specify the values parameter
).update_layout(
    title='Ratio of People Filling Out Forms',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=0,
    textposition='auto',
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Status</b>: %{value}<extra></extra>',
)

# --------------------- Activity Status --------------------- #

df_activity_status = df.groupby('Activity Status').size().reset_index(name='Count')

# print("Activity Status Unique Before: \n", df_activity_status['Activity Status'].unique().tolist())

# exclude null values:
df_activity_status = df_activity_status[df_activity_status['Activity Status'].notnull()]
df_activity_status = df_activity_status[df_activity_status['Activity Status'].str.strip() != '']

status_bar=px.bar(
    df_activity_status,
    x='Activity Status',
    y='Count',
    color='Activity Status',
    text='Count',
).update_layout(
    height=460, 
    width=780,
    title=dict(
        text='Activity Status',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=0,  # Rotate x-axis labels for better readability
        tickfont=dict(size=18),  # Adjust font size for the tick labels
        title=dict(
            # text=None,
            text="Status",
            font=dict(size=20),  # Font size for the title
        ),
        # showticklabels=False  # Hide x-tick labels
        showticklabels=True  # Hide x-tick labels
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  # Font size for the title
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        y=1,  # Position legend at the top
        xanchor="left",  # Anchor legend to the left
        yanchor="top",  # Anchor legend to the top
        # visible=False
        visible=True
    ),
    hovermode='closest', # Display only one hover label per trace
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',
    hovertemplate='<b>Status:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

status_pie=px.pie(
    df_activity_status,
    names="Activity Status",
    values='Count'  # Specify the values parameter
).update_layout(
    title='Activity Status',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=0,
    textposition='auto',
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Status</b>: %{value}<extra></extra>',
)

# --------------------------- BMHC Activity -------------------------- #

# print("BMHC Activity Unique Before:", df["BMHC Activity"].unique().tolist())

bmhc_activity_unique = [
'Add/ Review Content', 'Website Troubleshooting', 'Organization', 'Organizational Efficiency', 'Impact Metrics', 'Care Network Related Strategy', 'Communications Support', 'Communication & Correspondence', 'Research & Planning', 'Organization Strategy', 'Record Keeping & Documentation', 'Key or Special Event Support', 'BMHC Co-Branding', 'Marketing Promotion', 'Update Newsletter', 'Compliance & Policy Enforcement', 'Office Management'
]

bmhc_activity_categories = [
    "Communication & Correspondence",
    "Compliance & Policy Enforcement",
    "HR Support",
    "Office Management",
    "Record Keeping & Documentation",
    "Research & Planning"
]

df['BMHC Activity'] = (
    df['BMHC Activity']
    .str.strip()
    .replace({
        "" : "",
        'Add/ Review Content': 'Communication & Correspondence',
        'Website Troubleshooting': 'Office Management',
        'Organization': 'Office Management',
        'Organizational Efficiency': 'Office Management',
        'Impact Metrics': 'Research & Planning',
        'Care Network Related Strategy': 'Research & Planning',
        'Communications Support': 'Communication & Correspondence',
        'Communication & Correspondence': 'Communication & Correspondence',
        'Research & Planning': 'Research & Planning',
        'Organization Strategy': 'Research & Planning',
        'Record Keeping & Documentation': 'Record Keeping & Documentation',
        'Key or Special Event Support': 'Office Management',
        'BMHC Co-Branding': 'Communication & Correspondence',
        'Marketing Promotion': 'Communication & Correspondence',
        'Update Newsletter': 'Communication & Correspondence',
        'Compliance & Policy Enforcement': 'Compliance & Policy Enforcement',
        'Office Management': 'Office Management'
    })
)

# Identify unexpected/unapproved categories
bmhc_unexpected = df[~df['BMHC Activity'].isin(bmhc_activity_categories)]
# print("BMHC Activity Unexpected: \n", bmhc_unexpected['BMHC Activity'].unique().tolist())

# print("BMHC Activity Unique After:", df["BMHC Activity"].unique().tolist())

# Product Type dataframe:
bmhc_activity = df.groupby('BMHC Activity').size().reset_index(name='Count')

bmhc_bar=px.bar(
    bmhc_activity,
    x='BMHC Activity',
    y='Count',
    color='BMHC Activity',
    text='Count',
).update_layout(
    height=990, 
    width=1700,
    title=dict(
        text=f'{current_month} BMHC Activities',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,
        tickfont=dict(size=18),  
        title=dict(
            # text=None,
            text="Activity",
            font=dict(size=20), 
        ),
        showticklabels=False  
        # showticklabels=True  
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20), 
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        y=1,  # Position legend at the top
        xanchor="left",  # Anchor legend to the left
        yanchor="top",  # Anchor legend to the top
        # visible=False
        visible=True
    ),
    hovermode='closest', # Display only one hover label per trace
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
    margin=dict(t=50, r=50, b=30, l=40),  # Remove margins
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Person Pie Chart
bmhc_pie=px.pie(
    bmhc_activity,
    names="BMHC Activity",
    values='Count'  # Specify the values parameter
).update_layout(
    height=950,
    width=1700,
    title=f'{current_month} Ratio of BMHC Activities',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=90,
    textposition='auto',
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Activity</b>: %{value}<extra></extra>',
)

# ============================ Care Network Activity ================================ #

# print("Care Activity Unique Before: \n", df["Care Activity"].unique().tolist()) 

care_activity_unique = [
'Website Updates', 'Meeting', 'Internal Communications', 'No Product - Organizational Efficiency', 'No product - organizational strategy', 'Community Collaboration', 'Organizational Support', 'No Product - Organizational Support', 'Human Resources Training for Efficiency', 'Clinical Provider', 'Workforce Development', 'Special Announcement', 'AmeriCorps Responsibility', 'no product', 'MarCom Report', 'Distribution list', 'Announcement', 'Government', 'Internal Communications, No Product - Organizational Efficiency', 'No Product - Organizational Efficiency,', 'Timesheet', 'Event Support-Catering', 'Event Support', 'Marcom Report', 'No product- Organizational Efficiency', 'no product- organization efficiency', 'Newsletter', 'Flyer', 'Internal Communications, Meeting', 'Organizational Efficiency', 'Co-Branding, Community Collaboration, Planning - BMHC - Austin Public Health - Sustainable Foods Prostate Cancer Class: Thursday, April 24th at Metropolitan A.M.E.', 'Mental Hellness - The Bartley Method - video editing', 'no product - organizational efficiency', 'no product - Organizational Efficiency', 'Co-Branding, Flyer', 'Archive list', 'No product', 'MarCom Impact Report', 'Writing, Editing, Proofing', 'Social Media Post', 'Flyer, Registration for prostate cancer class', 'Promotion', 'Organized press release pdfs', 'Academic', 'SDoH Provider'
]

care_activity_categories = [
    "Academic",
    "Clinical Provider",
    "Give Back Program",
    "Government",
    "Religious",
    "SDoH Provider",
    "Workforce Development",
    "No Product",
]

df['Care Activity'] = (
    df['Care Activity']
    .str.strip()
    .replace({
        "" : "",
        'Website Updates': 'Academic',
        'Meeting': 'Academic',
        'Internal Communications': 'Academic',
        'No Product - Organizational Efficiency': 'No Product',
        'No product - organizational strategy': 'No Product',
        'Community Collaboration': 'Give Back Program',
        'Organizational Support': 'No Product',
        'No Product - Organizational Support': 'No Product',
        'Human Resources Training for Efficiency': 'Workforce Development',
        'Clinical Provider': 'Clinical Provider',
        'Workforce Development': 'Workforce Development',
        'Special Announcement': 'Academic',
        'AmeriCorps Responsibility': 'Workforce Development',
        'no product': 'No Product',
        'MarCom Report': 'Academic',
        'Distribution list': 'Academic',
        'Announcement': 'Academic',
        'Government': 'Government',
        'Internal Communications, No Product - Organizational Efficiency': 'No Product',
        'No Product - Organizational Efficiency,': 'No Product',
        'Timesheet': 'No Product',
        'Event Support-Catering': 'Give Back Program',
        'Event Support': 'Give Back Program',
        'Marcom Report': 'Academic',
        'No product- Organizational Efficiency': 'No Product',
        'no product- organization efficiency': 'No Product',
        'Newsletter': 'Academic',
        'Flyer': 'Academic',
        'Internal Communications, Meeting': 'Academic',
        'Organizational Efficiency': 'No Product',
        'Co-Branding, Community Collaboration, Planning - BMHC - Austin Public Health - Sustainable Foods Prostate Cancer Class: Thursday, April 24th at Metropolitan A.M.E.': 'Give Back Program',
        'Mental Hellness - The Bartley Method - video editing': 'Academic',
        'no product - organizational efficiency': 'No Product',
        'no product - Organizational Efficiency': 'No Product',
        'Co-Branding, Flyer': 'Academic',
        'Archive list': 'Academic',
        'No product': 'No Product',
        'MarCom Impact Report': 'Academic',
        'Writing, Editing, Proofing': 'Academic',
        'Social Media Post': 'Academic',
        'Flyer, Registration for prostate cancer class': 'Academic',
        'Promotion': 'Academic',
        'Organized press release pdfs': 'Academic',
        'Academic': 'Academic',
        'SDoH Provider': 'SDoH Provider',
    })
)

# normalized_categories = {cat.lower().strip(): cat for cat in _categories}
# counter = Counter()

# for entry in df['Support']:
#     items = [i.strip().lower() for i in entry.split(",")]
#     for item in items:
#         if item in normalized_categories:
#             counter[normalized_categories[item]] += 1

# # Display the result
# # for category, count in counter.items():
# #     print(f"Support Counts: \n {category}: {count}")

# df_ = pd.DataFrame(counter.items(), columns=['', 'Count']).sort_values(by='Count', ascending=False)

# Identify unexpected/unapproved categories
care_unexpected = df[~df['Care Activity'].isin(care_activity_categories)]
# print("Care Activity Unexpected: \n", care_unexpected['Care Activity'].unique().tolist())

# print("Care Activity Unique After:", df["Care Activity"].unique().tolist())

# Product Type dataframe:
care_activity = df.groupby('Care Activity').size().reset_index(name='Count')

care_bar = px.bar(
    care_activity,
    x='Care Activity',
    y='Count',
    color='Care Activity',
    text='Count',
).update_layout(
    height=990, 
    width=1700,
    title=dict(
        text=f'{current_month} Care Activities',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,
        tickfont=dict(size=18),  
        title=dict(
            text="Activity",
            font=dict(size=20), 
        ),
        showticklabels=False
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20), 
        ),
    ),
    legend=dict(
        title_text='',
        orientation="v",
        x=1.05,
        y=1,
        xanchor="left",
        yanchor="top",
        visible=True
    ),
    hovermode='closest',
    bargap=0.08,
    bargroupgap=0,
    margin=dict(t=50, r=50, b=30, l=40),
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Person Pie Chart
care_pie = px.pie(
    care_activity,
    names="Care Activity",
    values='Count'
).update_layout(
    height=950,
    width=1700,
    title=f'{current_month} Ratio of Care Activities',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=90,
    textposition='auto',
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Activity</b>: %{value}<extra></extra>',
)

# ============================ Outreach Activity ================================ #

# print("Outreach Activity Unique Before: \n", df["Outreach Activity"].unique().tolist())
# print("Outreach Activity value counts", df['Outreach Activity'].value_counts())

outreach_unique = [
'', 'Created special HTML BMHC Announcement of the CommUnity Care Clinic Measles Clinic in MailChimp', 'Updated email distribution list for special HTML BMHC Announcement of the CommUnity Care Clinic Measles Clinic', 'Health awareness, equity, tips, and services info for the public', 'No Product - Organizational Human Resources Support', 'Event (in-person), Social Media Post, Visuals', 'None', 'Visuals', 'Vector Logo', 'Timesheet', 'Overcoming Mental Hellness logo', 'PSA/ Commercial', 'Blogs', 'DACC Meeting', 'Newsletter plan', 'Social Media Post, schedule SWAD post', 'Newsletter/ Social Media Analytics', 'BMHC Services', 'Social Media Post, Shared partner post with Areebah', 'meeting with Areebah', 'Social Media Post, schedule CUC post on social media', 'Social Media Post', 'Quarterly Team Meeting', 'Updated Newsletter', 'Visuals, updated food sustainable flyer , created Mr. Larry Wallace Sr. Congratulations flyer , Stress Awareness flyer', 'updated Overcoming Mental Hellness Logo', 'Website Updates', 'Special Event planning', 'Event (in-person)', 'Website', 'Event Logistics', 'Key Event Logistics', 'Videography', 'Event (in-person), Special events set up', 'purchase food for special events', 'Social Media Post, Visuals', 'Updated the March Impact Report', 'updated Q2 Report', 'Social Media Post, update Social Media Coverage page visual', 'Visuals, update board slide', 'Visuals, Updated the March Impact Report', 'Visuals, Updated Bartley’s Way Documents', 'Created the Q2 Board Member Meeting slides', 'updated Military slides', 'Meeting with Areebah', 'meeting with Carlos Bautista', 'Updated Q2 Report', 'Key Leader Meeting', 'Created and sent Newsletter Plan or 4/25', 'Visuals, updated and sent DACC flyer for approval', 'updated newsletter', 'Social Media Post, reviewed partner posts of social media', 'Key Leader huddle'
]

outreach_categories = [
    "Event (in-person)",
    "Handouts",
    "Press Releases",
    "PSA/ Commercial",
    "Social Media Post",
    "Videography",
    "Visuals",
    "Website",
    "N/A",
]

df['Outreach Activity'] = (
    df['Outreach Activity']
    .str.strip()
    .replace({
        "" : "N/A",
        'Created special HTML BMHC Announcement of the CommUnity Care Clinic Measles Clinic in MailChimp': 'Press Releases',
        'Updated email distribution list for special HTML BMHC Announcement of the CommUnity Care Clinic Measles Clinic': 'Press Releases',
        'Health awareness, equity, tips, and services info for the public': 'Social Media Post',
        'No Product - Organizational Human Resources Support': '',
        'Event (in-person), Social Media Post, Visuals': 'Event (in-person)',
        'None': '',
        'Visuals': 'Visuals',
        'Vector Logo': 'Visuals',
        'Timesheet': '',
        'Overcoming Mental Hellness logo': 'Visuals',
        'PSA/ Commercial': 'PSA/ Commercial',
        'Blogs': 'Press Releases',
        'DACC Meeting': 'Event (in-person)',
        'Newsletter plan': 'Press Releases',
        'Social Media Post, schedule SWAD post': 'Social Media Post',
        'Newsletter/ Social Media Analytics': 'Social Media Post',
        'BMHC Services': 'Handouts',
        'Social Media Post, Shared partner post with Areebah': 'Social Media Post',
        'meeting with Areebah': 'Event (in-person)',
        'Social Media Post, schedule CUC post on social media': 'Social Media Post',
        'Social Media Post': 'Social Media Post',
        'Quarterly Team Meeting': 'Event (in-person)',
        'Updated Newsletter': 'Press Releases',
        'Visuals, updated food sustainable flyer , created Mr. Larry Wallace Sr. Congratulations flyer , Stress Awareness flyer': 'Visuals',
        'updated Overcoming Mental Hellness Logo': 'Visuals',
        'Website Updates': 'Website',
        'Special Event planning': 'Event (in-person)',
        'Event (in-person)': 'Event (in-person)',
        'Website': 'Website',
        'Event Logistics': 'Event (in-person)',
        'Key Event Logistics': 'Event (in-person)',
        'Videography': 'Videography',
        'Event (in-person), Special events set up': 'Event (in-person)',
        'purchase food for special events': 'Event (in-person)',
        'Social Media Post, Visuals': 'Social Media Post',
        'Updated the March Impact Report': 'Press Releases',
        'updated Q2 Report': 'Press Releases',
        'Social Media Post, update Social Media Coverage page visual': 'Social Media Post',
        'Visuals, update board slide': 'Visuals',
        'Visuals, Updated the March Impact Report': 'Visuals',
        'Visuals, Updated Bartley’s Way Documents': 'Visuals',
        'Created the Q2 Board Member Meeting slides': 'Visuals',
        'updated Military slides': 'Visuals',
        'Meeting with Areebah': 'Event (in-person)',
        'meeting with Carlos Bautista': 'Event (in-person)',
        'Updated Q2 Report': 'Press Releases',
        'Key Leader Meeting': 'Event (in-person)',
        'Created and sent Newsletter Plan or 4/25': 'Press Releases',
        'Visuals, updated and sent DACC flyer for approval': 'Visuals',
        'updated newsletter': 'Press Releases',
        'Social Media Post, reviewed partner posts of social media': 'Social Media Post',
        'Key Leader huddle': 'Event (in-person)',
    })
)

# Identify unexpected/unapproved categories
outreach_unexpected = df[~df['Outreach Activity'].isin(outreach_categories)]
# print("Outreach Activity Unexpected: \n", outreach_unexpected['Outreach Activity'].unique().tolist())

normalized_categories = {cat.lower().strip(): cat for cat in outreach_categories}
counter = Counter()

for entry in df['Outreach Activity']:
    items = [i.strip().lower() for i in entry.split(",")]
    for item in items:
        if item in normalized_categories:
            counter[normalized_categories[item]] += 1 # Count occurrences

# Display the result
# for category, count in counter.items():
#     print(f"Support Counts: \n {category}: {count}")

df_outreach = pd.DataFrame(counter.items(), columns=['Outreach Activity', 'Count']).sort_values(by='Count', ascending=False)

# Filtered dataframe to exclude all rows = 'N/A':
df_outreach = df_outreach[df_outreach['Outreach Activity'] != 'N/A']

# print("Outreach Activity Unique After:", df["Outreach Activity"].unique().tolist())

outreach_bar = px.bar(
    df_outreach,
    x='Outreach Activity',
    y='Count',
    color='Outreach Activity',
    text='Count',
).update_layout(
    height=990, 
    width=1700,
    title=dict(
        text=f'{current_month} Outreach Activities',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,
        tickfont=dict(size=18),  
        title=dict(
            text="Activity",
            font=dict(size=20), 
        ),
        showticklabels=False
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20), 
        ),
    ),
    legend=dict(
        title_text='',
        orientation="v",
        x=1.05,
        y=1,
        xanchor="left",
        yanchor="top",
        visible=True
    ),
    hovermode='closest',
    bargap=0.08,
    bargroupgap=0,
    margin=dict(t=50, r=50, b=30, l=40),
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Person Pie Chart
outreach_pie = px.pie(
    df_outreach,
    names="Outreach Activity",
    values='Count'
).update_layout(
    height=950,
    width=1700,
    title=f'{current_month} Ratio of Outreach Activities',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=0,
    textposition='auto',
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Activity</b>: %{value}<extra></extra>',
)

# ============================ Education Activity ================================ #

# print("Education Activity Unique Before: \n", df["Education Activity"].unique().tolist())

education_activity_unique = [
    '', 
    'Meeting with Areebah', 
    'Event',
    'Newsletter', 
    'PSA / Commercial', 
    'Impact Report/Timesheet', 
    'Social Media Post', 
    'Meeting', 
    'Handout, Visual', 
    'Visual',
    'Newsletter, Newsletter - Layout', 
    'Newsletter, Did You Know Articles',
    'Newsletter, Visual', 
    'PSA / Commercial, Special Announcement: Overcoming Mental Hellness', 
    'Special Announcement: Overcoming Mental Hellness', 
    'Visual, Research Survey Marketing Material for Manor Project', 
    'Videography',
    'What\'s in your soul?" Meeting Review', 
    'Visual, Update Impact In Action Powerpoint',
    'Visual, Research/Sent Marketing Material Recommendations for Manor Project',
    'Researched and Sent Schedule Post List to Areebah', 
    'Newsletter, Announcement', 
    'Reviewed Clockify Training', 
    'Impact Report edit', 
    'Timesheet', 
    'Updated Impact Report', 
    'Key Leader Huddle', 
    'Researched Pricing for Marketing Material for Manor Project',
    'Watched Clockify Training Video', 
    'Social Media Post, Visual', 
    'Announcement of new committee member', 
    'Visual, Announcement for Rod Sigler',
    'Community Outreach Activity (Physical Events)'
]

education_activity_categories = [
    "Event",
    "Handout",
    "Newsletter",
    "PSA / Commercial",
    "Social Media Post",
    "Videography",
    "Visual"
    "Meeting",
]

# Filter out null or empty strings after stripping spaces
df = df[df["Education Activity"].notnull()]
df = df[df["Education Activity"].str.strip() != '']

# Map various known strings to allowed categories (adjust mappings as needed)
mapping = {
    # Straight mappings
    'Event': 'Event',
    'Newsletter': 'Newsletter',
    'PSA / Commercial': 'PSA / Commercial',
    'Social Media Post': 'Social Media Post',
    'Videography': 'Videography',
    'Visual': 'Visual',
    'Handout': 'Handout',
    
    # Variations and multi-items mapped to one of the categories
    'Newsletter - Layout': 'Newsletter',
    'Newsletter, Newsletter - Layout': 'Newsletter',
    'Newsletter, Did You Know Articles': 'Newsletter',
    'Newsletter, Visual': 'Newsletter',
    'Newsletter, Announcement': 'Newsletter',
    'Social Media Post, Visual': 'Social Media Post',
    'Visual, Announcement for Rod Sigler': 'Visual',
    'Visual, Update Impact In Action Powerpoint': 'Visual',
    'Visual, Research Survey Marketing Material for Manor Project': 'Visual',
    'Visual, Research/Sent Marketing Material Recommendations for Manor Project': 'Visual',
    
    # Some that could be categorized as Event or Newsletter depending on your rules:
    'Community Outreach Activity (Physical Events)': 'Event',
    
    # Misc that don't fit well - treat as unknown or ignore
    'Meeting with Areebah': 'Meeting',
    'Meeting': 'Meeting',
    'Impact Report/Timesheet': 'Other',
    'Special Announcement: Overcoming Mental Hellness': 'Other',
    'What\'s in your soul?" Meeting Review': 'Meeting',
    'Researched and Sent Schedule Post List to Areebah': 'Other',
    'Reviewed Clockify Training': 'Other',
    'Impact Report edit': 'Other',
    'Timesheet': 'Other',
    'Updated Impact Report': 'Other',
    'Key Leader Huddle': 'Meeting',
    'Researched Pricing for Marketing Material for Manor Project': 'Other',
    'Watched Clockify Training Video': 'Other',
    'Announcement of new committee member': 'Other',
}

# Clean up column values by mapping and stripping
def map_education_activity(val):
    val = val.strip()
    # If exact mapping found, return it
    if val in mapping:
        return mapping[val]
    # If not found, check if comma-separated list and map first known category found
    parts = [part.strip() for part in val.split(",")]
    for part in parts:
        if part in mapping and mapping[part] is not None:
            return mapping[part]
    return None  # Uncategorized or unknown

df['Education Activity Clean'] = df['Education Activity'].apply(map_education_activity)

# Drop rows where cleaned category is None (uncategorized)
df = df[df['Education Activity Clean'].notnull()]

# Now group and count
education_activity_counts = df.groupby('Education Activity Clean').size().reset_index(name='Count')

# print("Education Activity Unique After:\n", education_activity_counts)

# If you want to count multiple activities in one cell separately, use this:
counter = Counter()
for entry in df['Education Activity']:
    parts = [p.strip() for p in entry.split(",")]
    for p in parts:
        cat = map_education_activity(p)
        if cat:
            counter[cat] += 1

education_activity = pd.DataFrame(counter.items(), columns=['Education Activity', 'Count']).sort_values(by='Count', ascending=False)

# print("\nCount of individual activities split by comma:\n", education_activity)

education_bar = px.bar(
    education_activity,
    x='Education Activity',
    y='Count',
    color='Education Activity',
    text='Count',
).update_layout(
    height=990, 
    width=1700,
    title=dict(
        text=f'{current_month} Education Activities',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,
        tickfont=dict(size=18),  
        title=dict(
            text="Activity",
            font=dict(size=20), 
        ),
        showticklabels=False
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20), 
        ),
    ),
    legend=dict(
        title_text='',
        orientation="v",
        x=1.05,
        y=1,
        xanchor="left",
        yanchor="top",
        visible=True
    ),
    hovermode='closest',
    bargap=0.08,
    bargroupgap=0,
    margin=dict(t=50, r=50, b=30, l=40),
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Person Pie Chart
education_pie = px.pie(
    education_activity,
    names="Education Activity",
    values='Count'
).update_layout(
    height=950,
    width=1700,
    title=f'{current_month} Ratio of Education Activities',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=70,
    textposition='auto',
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Activity</b>: %{value}<extra></extra>',
)

# ============================ Entity Name ================================ #

# print("Entity Names Unique Before: \n", df["Entity"].unique().tolist())

entity_unique = [
'', 'CommunityCare', 'None', "Black Men's Health Clinic", 'DACC Meeting', "Black Men's Health Clinic, Overcoming Mental Hellness logo", "Austin Public Health, Black Men's Health Clinic, Sustainable Food Center", 'AmeriCorps Duties', 'Austin Public Health, Sustainable Food Center', 'Austin Public Health, SFC', 'SFC'
]

entity_categories = [
    "Austin Public Health",
    "Black Men's Health Clinic",
    "City of Austin AmeriCorps",
    "Sustainable Food Center",
    "Central Health",
    "CommunityCare",
    "GudLife",
    "Integral Care",
    "DACC Meeting",
    "None",
]

df['Entity'] = (
    df['Entity']
    .str.strip()
    .replace({
        "AmeriCorps Duties" : "City of Austin AmeriCorps",
        "SFC" : "Sustainable Food Center",
    })
)

# Identify unexpected/unapproved categories
entity_unexpected = df[~df['Entity'].isin(entity_categories)]
# print("Entity Unexpected: \n", entity_unexpected['Entity'].unique().tolist())

# print("Entity Unique After:", df["Entity"].unique().tolist())

normalized_categories = {cat.lower().strip(): cat for cat in entity_categories}
counter = Counter()

for entry in df['Entity']:
    items = [i.strip().lower() for i in entry.split(",")]
    for item in items:
        if item in normalized_categories:
            counter[normalized_categories[item]] += 1

# for category, count in counter.items():
#     print(f"Support Counts: \n {category}: {count}")

df_entity = pd.DataFrame(counter.items(), columns=['Entity', 'Count']).sort_values(by='Count', ascending=False)

# df_entity = df.groupby('Entity').size().reset_index(name='Count')

entity_bar = px.bar(
    df_entity,
    x='Entity',
    y='Count',
    color='Entity',
    text='Count',
).update_layout(
    height=990, 
    width=1700,
    title=dict(
        text=f'{current_month} Entity Names',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,
        tickfont=dict(size=18),  
        title=dict(
            text="Name",
            font=dict(size=20), 
        ),
        showticklabels=False
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20), 
        ),
    ),
    legend=dict(
        title_text='',
        orientation="v",
        x=1.05,
        y=1,
        xanchor="left",
        yanchor="top",
        visible=True
    ),
    hovermode='closest',
    bargap=0.08,
    bargroupgap=0,
    margin=dict(t=50, r=50, b=30, l=40),
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Entity Pie Chart
entity_pie = px.pie(
    df_entity,
    names="Entity",
    values='Count'
).update_layout(
    height=950,
    width=1700,
    title=f'{current_month} Ratio of Entity Names',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=-80,
    textposition='auto',
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='<b>%{label} Activity</b>: %{value}<extra></extra>',
)

# # ========================== DataFrame Table ========================== #

# MarCom Table
marcom_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df[col] for col in df.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server

app.layout = html.Div(
    children=[ 
    html.Div(
        className='divv', 
        children=[ 
        html.H1(
            'MarCom Report', 
            className='title'),
        html.H1(
            f'{current_month} {report_year}', 
            className='title2'),
    html.Div(
        className='btn-box', 
        children=[
        html.A(
            'Repo',
            href=f'https://github.com/CxLos/MC_Apr_{report_year}',
            className='btn'),
        ]),
    ]),    

# Data Table
# html.Div(
#     className='row0',
#     children=[
#         html.Div(
#             className='table',
#             children=[
#                 html.H1(
#                     className='table-title',
#                     children='Data Table'
#                 )
#             ]
#         ),
#         html.Div(
#             className='table2', 
#             children=[
#                 dcc.Graph(
#                     className='data',
#                     figure=marcom_table
#                 )
#             ]
#         )
#     ]
# ),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_month} MarCom Events']
            ),
            html.Div(
                className='circle',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high2',
                            children=[marcom_events]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high3',
                children=[f'{current_month} MarCom Hours']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[marcom_hours]
                    ),
                        ]
                    )
                ],
            ),
            ]
        ),
    ]
),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_month} MC Travel Hours']
            ),
            html.Div(
                className='circle',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high6',
                            children=[mc_travel]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
                html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high3',
                children=['Blank']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            # children=[marcom_hours]
                    ),
                        ]
                    )
                ],
            ),
            ]
        ),
    ]
),

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=bmhc_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=bmhc_pie
                )
            ]
        ),
    ]
),   
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=care_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=care_pie
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=outreach_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=outreach_pie
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=education_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=education_pie
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=entity_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=entity_pie
                )
            ]
        ),
    ]
),   

# ROW 4
html.Div(
    className='row2',
    children=[
        html.Div(
            className='graph1',
            children=[                
                dcc.Graph(
                    figure=person_bar
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[
                dcc.Graph(
                    figure=person_pie
                )
            ],
        ),
    ]
),

# ROW 4
html.Div(
    className='row2',
    children=[
        html.Div(
            className='graph1',
            children=[                
                dcc.Graph(
                    figure=status_bar
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[
                dcc.Graph(
                    figure=status_pie
                )
            ],
        ),
    ]
),
])

print(f"Serving Flask app '{current_file}'! 🚀")

if __name__ == '__main__':
    app.run_server(debug=
                   True)
                #    False)
# =================================== Updated Database ================================= #

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# updated_path = f'data/MarCom_{current_month}_{report_year}.xlsx'
# data_path = os.path.join(script_dir, updated_path)

# with pd.ExcelWriter(data_path, engine='xlsxwriter') as writer:
#     df.to_excel(
#             writer, 
#             sheet_name=f'MarCom {current_month} {report_year}', 
#             startrow=1, 
#             index=False
#         )

#     # Create the workbook to access the sheet and make formatting changes:
#     workbook = writer.book
#     sheet1 = writer.sheets['MarCom April 2025']
    
#     # Define the header format
#     header_format = workbook.add_format({
#         'bold': True, 
#         'font_size': 13, 
#         'align': 'center', 
#         'valign': 'vcenter',
#         'border': 1, 
#         'font_color': 'black', 
#         'bg_color': '#B7B7B7',
#     })
    
#     # Set column A (Name) to be left-aligned, and B-E to be right-aligned
#     left_align_format = workbook.add_format({
#         'align': 'left',  # Left-align for column A
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })

#     right_align_format = workbook.add_format({
#         'align': 'right',  # Right-align for columns B-E
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })
    
#     # Create border around the entire table
#     border_format = workbook.add_format({
#         'border': 1,  # Add border to all sides
#         'border_color': 'black',  # Set border color to black
#         'align': 'center',  # Center-align text
#         'valign': 'vcenter',  # Vertically center text
#         'font_size': 12,  # Set font size
#         'font_color': 'black',  # Set font color to black
#         'bg_color': '#FFFFFF'  # Set background color to white
#     })

#     # Merge and format the first row (A1:E1) for each sheet
#     sheet1.merge_range('A1:Q1', f'MarCom Report {current_month} {report_year}', header_format)

#     # Set column alignment and width
#     # sheet1.set_column('A:A', 20, left_align_format)  

#     print(f"MarCom Excel file saved to {data_path}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create mc-impact-11-2024
# heroku git:remote -a mc-impact-11-2024
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx