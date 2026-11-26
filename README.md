# Darelle Couture

## Fashion, Tailoring & Business Analytics

Darelle Couture is a responsive web application combining
a fashion and tailoring website with an interactive
business analytics dashboard.

The project demonstrates how Excel business data can be
transformed into dynamic web-based reports, pivot tables,
charts, and key performance indicators.

## Project Objectives

- Build a professional fashion and tailoring website.
- Display clothing products and tailoring services.
- Import and process Excel business data.
- Generate dynamic pivot tables.
- Visualize business performance.
- Automatically update dashboard statistics.
- Practice Git and GitHub version control.

## Technologies Used

- HTML5
- CSS3
- JavaScript
- SheetJS
- Chart.js
- Chart.js DataLabels
- Microsoft Excel
- Visual Studio Code
- Git
- GitHub

## Website Features

### Fashion and Tailoring

- Professional homepage
- Product catalogue
- Tailoring services
- Responsive design
- Canadian dollar pricing (CAD)

### Business Analytics

- Excel file import
- Automatic worksheet detection
- Data validation
- Dynamic pivot tables
- Key performance indicators
- Orders by service chart
- Service value chart
- Automatic dashboard updates

### Data Validation

- Supported file format: .xlsx
- Maximum file size: 5 MB
- Maximum records: 10,000
- Required column validation
- Invalid price detection
- Error handling

## Project Structure

    Site-web-darelle/
    |
    |-- Data/
    |-- Images/
    |-- index.html
    |-- analytics.js
    |-- .gitignore
    |-- README.md

## Data Requirements

The Excel worksheet must contain the following columns:

| Column | Description |
|--------|-------------|
| service_name | Service category |
| base_price | Service price in CAD |
| order_id | Order identifier |

## Getting Started

1. Clone the GitHub repository.
2. Open the project in Visual Studio Code.
3. Install the Live Server extension.
4. Open index.html using Live Server.
5. Navigate to Business Analytics.
6. Import a compatible Excel file.

## Data Privacy

Confidential Excel files are excluded from version control.

Only fictional or anonymized business data should be
used for public demonstrations.

File validation is performed in the browser.
This prototype does not implement administrator
authentication or server-side access control.

## Git Branching Strategy

- main: Stable production version.
- develop: Development and integration.
- feature/*: New features and improvements.

## Future Improvements

- Administrator authentication
- Secure backend integration
- Database integration
- Advanced analytics filters
- Product ordering system
- Online payment integration
- Cloud deployment

## Author

Hugo Loic Keundo Nana

IT Support | Networking | Cloud & DevOps Learner

## Project Status

Functional prototype under development.
