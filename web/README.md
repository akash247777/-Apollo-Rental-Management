# Apollo Rental Management System

A modern web application for managing rental data for retail locations. This application provides functionality for storing, viewing, editing, and reporting on rental agreements.

## Features

- **Secure Login**: Authentication system to protect sensitive data
- **Dashboard**: View and search for site information
- **Data Entry**: Add new site records with comprehensive details
- **Edit Records**: Modify existing site information
- **Report Generation**: Generate different types of reports including:
  - Hike Reports
  - Rent Reports
  - Owner Wise Reports
  - Negotiation Reports
  - Lease Period Reports
  - All Sites Data Reports
- **Excel Integration**: Upload site data from Excel files
- **Responsive Design**: Works on desktop, tablet, and mobile devices

## Application Structure

- `index.html` - Login page
- `dashboard.html` - Main dashboard for viewing site information
- `new-entry.html` - Form for adding new site records
- `reports.html` - Report generation interface
- `css/style.css` - Styling for the entire application
- `js/` - JavaScript files for each page:
  - `login.js` - Login page functionality
  - `dashboard.js` - Dashboard page functionality
  - `new-entry.js` - New entry form functionality
  - `reports.js` - Report generation functionality
- `img/` - Images used in the application

## Setup

1. Clone this repository
2. Set up a web server (Apache, Nginx, etc.) to serve the files
3. Configure the back-end API endpoint in the JavaScript files (currently using dummy data)
4. Set up a database to store the rental data

## Technical Implementation

This web application is built using:

- **HTML5** for structure
- **CSS3** for styling (responsive design, grid layout)
- **JavaScript** for client-side functionality

For a production environment, you would also need:

- A backend server (Node.js, Python, PHP, etc.)
- A database (SQL Server, as used in the original application)
- API endpoints for CRUD operations

## Usage

### Login

Use the following credentials to log in:
- Username: krishna, Password: krishna@123
- Username: kuber, Password: kuber@123

### Navigation

The left sidebar provides navigation between different sections of the application:
- Dashboard
- New Entry
- Reports
- Logout

### Search

Use the search box in the dashboard to find site information by site ID.

### Adding New Records

Click the "New Entry" link in the sidebar to access the form for adding new site records.

### Generating Reports

Click the "Reports" link in the sidebar to access the report generation interface. Select the report type, date range, and any additional filters, then click "Generate Report".

## Conversion Notes

This web application is a conversion of a Python/Tkinter desktop application to a modern web interface. The key improvements include:

- Modern, responsive UI
- Improved navigation
- Streamlined data entry
- Enhanced reporting capabilities
- Consistent design language throughout the application

## Future Enhancements

- Add real-time data validation
- Implement API integration
- Add user management features
- Enhance security with JWT authentication
- Add dark mode theme
- Implement more advanced filtering in reports
- Add data visualization/charts 