# SSMU Club Audit Helper
---
## Welcome!

This is an internally developed program for SSMU, designed to assist auditing process of student-run clubs.

---

## Getting Started

### First time set-up

- First, create a folder called 'Transaction Lists'

> This is where all of your transaction lists should be stored

- Inside that folder, create a folder called 'Audit Report'

> This is where the reports will be automatically generated from the transaction lists

### Required files

- You need at least 3 excel files for the program to run:

> - list of clubs' unique identification (CLU)
- list of concerning keywords you're interested in detecting
- transaction list(s)

- Set the file paths:
The top 4 variables indicate where your files are located / to be located.

>- file_names -> indicates file path to the list of clubs
- keywords -> indicates file path to the list of keywords
- audit_report_file_path -> where the reports are to be populated
- clu_file_path -> the transaction list follows naming convention containing the club's CLU number. Enter file path just up until CLU

## Author
- Richie Youm
