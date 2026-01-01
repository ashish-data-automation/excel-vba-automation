# Excel VBA Automation Projects

This repository contains **real-world Excel VBA automation examples** designed to
reduce manual work, save time, and improve accuracy.

All projects here are based on **practical business problems**, not theoretical examples.

---

## 🔍 What This Repository Is About

Many teams use Excel daily for:
- Reports
- Invoices
- Data processing
- Repetitive calculations

Doing these tasks manually leads to:
- Human errors
- Time wastage
- Inconsistent results

This repository shows how **Excel VBA automation** can solve these problems effectively.

---

## 📂 Projects Included

### 1️⃣ Invoice Automation
📁 `invoice-automation/`

**Problem:**  
Invoices are often prepared manually in Excel, which is repetitive and error-prone.

**Solution:**  
An automated Excel VBA solution that generates invoices from structured input data
with minimal manual effort.

**What it demonstrates:**
- Data-driven invoice creation
- Automated formatting
- Reduced manual intervention
- Business-friendly Excel automation

---

## 🛠 Skills Demonstrated

- Excel VBA macro development
- Automation logic & workflow design
- Structured Excel-based solutions
- Real-life business problem solving

---

## 🎯 Who This Is Useful For

- Small businesses
- Freelancers
- Accounts & finance teams
- Anyone using Excel for repetitive tasks

---

## 📌 Notes

- Projects are kept **simple and practical**
- Code is written for **clarity and usability**
- Each project folder contains its own README with explanation

More automation examples will be added gradually.

---

## 📬 Connect With Me

If you’re looking to automate your Excel workflows or repetitive tasks:

- Fiverr: https://www.fiverr.com/s/2KB3Dgq
- LinkedIn: https://www.linkedin.com/in/ashish-jain-4581ba3a2/

Practical automation is always better than manual work.


---

## 🔧 Sample VBA Macro (Invoice Generation)

Below is a sample VBA macro that demonstrates how invoice automation
can be implemented in Excel.

```vba
Sub GenerateInvoice()

    Dim wsInput As Worksheet
    Dim wsInvoice As Worksheet
    Dim lastRow As Long
    Dim totalAmount As Double
    Dim i As Long

    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsInvoice = ThisWorkbook.Sheets("Invoice")

    wsInvoice.Range("B5:E20").ClearContents
    totalAmount = 0

    wsInvoice.Range("B2").Value = wsInput.Range("B2").Value
    wsInvoice.Range("B3").Value = wsInput.Range("B3").Value

    lastRow = wsInput.Cells(wsInput.Rows.Count, "B").End(xlUp).Row

    For i = 5 To lastRow
        wsInvoice.Cells(i, "B").Value = wsInput.Cells(i, "B").Value
        wsInvoice.Cells(i, "C").Value = wsInput.Cells(i, "C").Value
        wsInvoice.Cells(i, "D").Value = wsInput.Cells(i, "D").Value
        wsInvoice.Cells(i, "E").Value = wsInput.Cells(i, "C").Value * wsInput.Cells(i, "D").Value

        totalAmount = totalAmount + wsInvoice.Cells(i, "E").Value
    Next i

    wsInvoice.Range("E10").Value = totalAmount

    MsgBox "Invoice generated successfully!", vbInformation

End Sub

How to Use

1.Open Excel
2.Press ALT + F11
3.Insert → Module
4.Paste the macro
5.Run GenerateInvoice



