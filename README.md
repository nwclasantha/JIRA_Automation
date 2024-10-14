# Jira Task Management with Excel Formatter
![image](https://github.com/user-attachments/assets/c5fb8b48-4b2f-4ed0-ae3d-8924d06d0df8)

This Python script is designed to interact with Jira to fetch and update issues for the **XXX** project, export the issues to an Excel file, and format the Excel file with priority-based coloring and other enhancements. Additionally, it can update Jira issues directly from an Excel file.

## Features
- Fetch Jira issues using Jira REST API.
- Export Jira issues to an Excel file.
- Apply custom formatting to Excel files including colored rows based on priority.
- Update Jira issues directly from an Excel file, including updating fields like summary, priority, and assignee.
- Use Jinja2 templating to generate descriptions for Jira issues.

## Prerequisites

Before running this script, make sure you have installed the required Python packages. You can install them by running:

```bash
pip install requests pandas jinja2 openpyxl
```
