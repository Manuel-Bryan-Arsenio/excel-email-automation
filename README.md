# **Excel Email Automation**

## **Overview**
This project automates the process of sending interview invitations via email, using Python. The program reads candidate information from an Excel file and sends tailored email invitations based on interview format (online or offline). It also checks for scheduling conflicts and missing data, ensuring smooth communication with candidates. This project was assigned to me by the Chair of Aerospace Structures at the Technical University of Munich to automate the email invitation process for interview applicants.

## **Features**
- **Excel Data Extraction**: Automatically reads candidate information (name, email, interview date, time, etc.) from an Excel sheet.
- **Data Cleanup**: Handles missing data by replacing empty fields (e.g., missing interview dates or locations) with default values.
- **Conflict Detection**: Checks for overlapping interview dates and times, notifying if any conflicts are found.
- **Email Automation**: Sends personalized email invitations to candidates based on the interview format (online/offline).
- **Error Handling**: Stops execution if critical issues like scheduling conflicts or missing data are detected.

## **Technologies Used**
- **Pandas**: For data extraction and manipulation from the Excel file.
- **NumPy**: To handle missing data (NaN) values.
- **SMTP**: For sending automated emails via Gmail.
- **Sys**: For error handling and program termination.

## **Future Improvements**
- Optimize message handling to reduce complexity by separating message templates into a reusable module.
- Improve conflict detection to handle multiple candidates with overlapping schedules.
- Add support for multiple email providers besides Gmail.

## **Contributing**
Feel free to fork this project, open issues, and submit pull requests. Contributions are welcome!
"# excel-email-automation" 
