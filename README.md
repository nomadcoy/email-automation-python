# Email Automation with Python

## Overview
This project automates the process of sending emails using Python based on a structured dataset.
It was developed as a portfolio project to demonstrate automation, data handling, and practical scripting for real-world tasks.

The focus is not only on automation itself, but on designing a reproducible and data-driven workflow.

---

## Technologies
- Python
- pandas
- pyautogui
- openpyxl
- PowerShell
- Virtual Environment (venv)

---

## How the Project Works
1. A dataset (Excel or CSV) contains the email recipients and relevant fields used for personalization.
2. The script reads and processes the dataset using pandas.
3. For each record, the script automates the email composition and sending process based on a predefined template.
4. Emails are sent individually, allowing controlled and personalized communication at scale.

---

## Technical Considerations
- This project uses GUI automation with `pyautogui`.
- While the script is running:
  - The computer should not be used for other tasks.
  - Screen resolution and interface layout must remain unchanged.
- This approach was chosen intentionally to explore automation under constrained environments.

---

## Limitations
- GUI-based automation is sensitive to interface changes.
- Accented characters may require specific handling depending on the environment.
- The script is not intended for large-scale or background execution.

---

## Possible Improvements
- Replace GUI automation with SMTP or email APIs.
- Add structured logging and exception handling.
- Implement validation for email addresses.
- Introduce rate control and execution reports.

---

## Project Structure
