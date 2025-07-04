# Python Project for Document Generation

This project provides a set of tools for generating various types of documents, including customs reports, monthly agreements, and acceptance-transfer acts. The application is structured to facilitate easy access to these functionalities through a main menu interface.

## Project Structure

```
python-project
├── src
│   ├── menu.py                     # Main menu for the application
│   ├── custom_reports.py            # Logic for generating customs reports
│   ├── monthly_agreements.py        # Logic for creating monthly agreements
│   └── acceptance_transfer_acts.py  # Logic for creating acceptance-transfer acts
├── templates
│   ├── excel
│   │   └── custom_report_template.xlsx  # Excel template for customs reports
│   └── word
│       ├── monthly_agreement_template.docx  # Word template for monthly agreements
│       └── acceptance_transfer_act_template.docx  # Word template for acceptance-transfer acts
└── README.md                       # Documentation for the project
```

## Setup Instructions

1. Clone the repository to your local machine.
2. Navigate to the project directory.
3. Ensure you have the required libraries installed. You may need libraries such as `pandas`, `openpyxl`, and `python-docx` for handling Excel and Word documents.
4. Run the application by executing `python src/menu.py`.

## Usage Guidelines

- Upon running the application, you will be presented with a menu to choose from the available functionalities.
- Follow the prompts to input the necessary data for generating reports, agreements, or acts.
- The generated documents will be formatted according to the provided templates.

## Contributing

Contributions to enhance the functionality or improve the documentation are welcome. Please submit a pull request with your changes.