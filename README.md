![Version](https://img.shields.io/badge/version-v0.3-blue)
![Status](https://img.shields.io/badge/status-active-success)
![License](https://img.shields.io/badge/license-MIT-green)
![Excel](https://img.shields.io/badge/Excel-VBA-darkgreen)
![Feature](https://img.shields.io/badge/feature-Table_First_Design-orange)

# Invoice Generator

Reusable Excel VBA toolkit for invoice generation and business workflow automation.

Built for small businesses, freelancers, and operations teams that need a simple, maintainable invoice automation solution using Microsoft Excel.

---

## Architecture

<p align="center">
  <img src="images/architecture.svg" width="800" alt="Invoice Generator Architecture">
</p>

---

## Features

✓ Excel Table-based Data Model

✓ Customer Master Management

✓ Customer ID Dropdown

✓ Automatic Customer Lookup (XLOOKUP)

✓ Customer Information Auto Fill

✓ Invoice Number Generation

✓ Payment Terms Support

✓ Automatic Due Date Calculation

✓ Currency Lookup

✓ Tax Rate Lookup

✓ Invoice Item Table

✓ Input Validation (Excel VBA)

✓ Print-ready Layout

✓ PDF-ready Design

✓ Reusable VBA Modules

✓ Business Workflow Automation

---

## Example Workflow

```text
tblCustomers
        │
        ▼
Invoice Input
        │
        ▼
Customer Lookup
(Customer / Currency / Tax / Payment Terms)
        │
        ▼
tblInvoiceItems
        │
        ▼
Subtotal
        │
        ▼
Tax
        │
        ▼
Total
        │
        ▼
tblSettings
        │
        ▼
Invoice Output
        │
        ▼
Print / PDF
```

---

## Project Structure

```text
Invoice-Generator
│
├── README.md
├── LICENSE
├── images/
├── sample/
└── src/
```

---

## Worksheets

```text
Customer Master
Invoice Input
Invoice Output
Settings
```

---

## Technologies

- Microsoft Excel
- Excel Tables
- Structured References
- XLOOKUP
- Excel VBA
- Data Validation
- Print Layout
- Business Automation

---

## Current Functions

- Customer Master management
- Customer ID dropdown list
- Automatic customer lookup
- Automatic customer information fill
- Automatic currency lookup
- Automatic tax rate lookup
- Payment terms lookup
- Automatic due date calculation
- Invoice number generation
- Invoice item table
- Automatic subtotal calculation
- Automatic tax calculation
- Automatic total calculation
- VBA input validation
- Print-ready invoice layout

---

## Future Roadmap

- PDF Export
- Email Sending
- Multi-Currency Support
- Invoice History
- Automatic Invoice Number Sequence
- Dashboard
- API Integration
- Cloud Integration

---

## Example Use Cases

- Small Business Invoicing
- Freelancer Invoice Generation
- Internal Billing
- Client Billing
- Service Invoice Templates
- Sales Operations
- Business Workflow Automation

---

## Why This Project?

Many businesses still create invoices manually in Excel.

This project demonstrates how repetitive billing tasks can be standardized and automated using Excel Tables, XLOOKUP, Structured References, and reusable VBA modules.

Customer information, invoice numbers, payment terms, due dates, tax rates, currencies, and invoice calculations are automatically managed through a table-first architecture designed for maintainability and scalability.

Rather than being an accounting system, this project focuses on practical business process automation using Microsoft Excel.

---

## Design Principles

- Table First Design
- Reusable VBA Modules
- Structured References
- Maintainable Workbook Architecture
- Business-oriented Automation
- Scalable Excel Solutions

---

## License

MIT License