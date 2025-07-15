#HR Attendance Dashboard

Overview
This is a Streamlit-based web application designed to process and analyze HR attendance data from Excel files. It calculates employee presence, generates styled Excel reports, and provides visualizations for attendance statistics. Built as a tool for data science learning, it includes features like anomaly detection and interactive filtering.
Prerequisites

Python 3.8 or higher
Internet connection (for dependencies and deployment)

Installation

Clone the repository:git clone <repository-url>
cd hr-attendance-dashboard

Install the required dependencies:pip install -r requirements.txt

(Note: Create a requirements.txt file with the following content and include it in the repository:streamlit
pandas
plotly
openpyxl
numpy

Usage

Run the app locally:streamlit run app.py

Open your browser at http://localhost:8501.

Upload an Excel file (e.g., Etat_Fin_de_Periode_17524776654333149930951483363918.xlsx) containing attendance data.
Navigate using the sidebar buttons:
Data Processing: View the processed data table and download a styled Excel report.
Stats: Explore summary statistics, anomaly detection, and visualizations (bar chart, box plot, histogram) with a service filter.

Interact with the app:
Use the "Filter by Service" dropdown on the Stats page to analyze specific services.
Download the processed data as an Excel file from the Data Processing page.

Deployment
To deploy the app on Streamlit Community Cloud:

Push your code to a GitHub repository.
Sign in to share.streamlit.io.
Create a new app:
Link your GitHub repository.
Set the main file to app.py.

Deploy the app and share the generated URL with your team.

Notes

The app caps visualization y-axes at 300 for readability, based on a maximum of 192 hours.
The "Presence" calculation uses D - ((H \* 8) + 14) for hourly workers and "Jours de Présences 24/5-23/6" for daily workers.
Customize the Observations column logic (e.g., "Prime anciennete:30 dt" vs. 50 dt) by modifying the process_excel_file function if needed.

Troubleshooting

Ensure all dependencies are up-to-date: pip install --upgrade streamlit pandas plotly openpyxl numpy.
Check for errors in the terminal or browser console if the app fails to load.
Verify the Excel file format matches the expected columns (e.g., "Matricule", "Nom", "Heures Normales", etc.).

Contributions
Feel free to fork this repository, make improvements, and submit pull requests. Suggestions for enhancing visualizations or adding machine learning features are welcome!
