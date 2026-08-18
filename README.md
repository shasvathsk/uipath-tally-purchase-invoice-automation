# UiPath TallyPrime Purchase Invoice Automation

## Overview

An RPA workflow developed with UiPath to automate purchase invoice entry into TallyPrime using structured invoice data stored in Excel.

The automation reads invoice records from an Excel DataTable and enters the required information into TallyPrime, including invoice details, supplier information, and amounts.

## Problem

Entering large numbers of purchase invoices manually into accounting software is repetitive and time-consuming.

An additional challenge occurs when multiple invoices belong to the same supplier. The supplier ledger should be created only once and reused for subsequent invoices.

This automation addresses both problems.

## Workflow

```text
Excel Invoice Data
        ↓
Read DataTable
        ↓
For Each Invoice
        ↓
Read Supplier Name
        ↓
Check Existing Ledger
        ↓
    ┌───────────────┐
    │               │
    ▼               ▼
Existing        New Supplier
Ledger             │
    │              ▼
    │           Alt + C
    │              ↓
    │        Create Ledger
    │              ↓
    └───────→ Continue
                  ↓
          Enter Purchase Invoice
                  ↓
          Enter Invoice Number
                  ↓
              Enter Date
                  ↓
          Select Supplier
                  ↓
          Enter Invoice Amount
                  ↓
                Save
```

## Duplicate Supplier Handling

The automation maintains a DataTable named `created` containing supplier ledger names that have already been created during the automation run.

Before creating a new ledger, the workflow checks whether the supplier already exists in the DataTable.

Example logic:

```vb
created.AsEnumerable().Any(
    Function(r) r("cust").ToString.Trim.Equals(
        CurrentRow("customer").ToString.Trim,
        StringComparison.OrdinalIgnoreCase
    )
)
```

If the supplier already exists:

```text
Do not create a new ledger
→ Use the existing supplier
```

If the supplier does not exist:

```text
Create the ledger
→ Add the supplier to the tracking DataTable
→ Continue invoice entry
```

This prevents duplicate supplier ledgers when multiple invoices belong to the same supplier.

## Technologies Used

* UiPath Studio
* TallyPrime
* Excel
* DataTables
* VB.NET expressions
* Keyboard automation
* Conditional logic

## UiPath Concepts Demonstrated

* Excel DataTable processing
* For Each Row
* Conditional branching
* DataTable lookup
* Duplicate detection
* Keyboard automation
* TallyPrime workflow automation
* Reusable invoice-processing logic

## Example Scenario

Input Excel:

| Invoice Number | Customer      |  Amount |
| -------------- | ------------- | ------: |
| 23122          | Thomas Boland | 6102.72 |
| 23123          | ABC Traders   | 4500.00 |
| 23124          | Thomas Boland | 2300.00 |

The automation creates:

```text
Thomas Boland → Created once
ABC Traders   → Created once
Thomas Boland → Existing ledger reused
```

The second Thomas Boland invoice does not trigger ledger creation again.

## Business Use Case

This automation can help reduce repetitive accounting data entry for businesses processing multiple purchase invoices.

Potential applications include:

* Purchase invoice entry
* Supplier ledger management
* Accounting data entry
* Bulk invoice processing
* Excel-to-Tally automation
* Repetitive bookkeeping workflows

## Important Note

This project is a demonstration of RPA workflow design and should be adapted and tested against the specific TallyPrime configuration, voucher settings, accounting structure, and business requirements of the organization using it.
