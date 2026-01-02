# Invoice Automation using Excel VBA

> ⚠️ This is a sample automation project created for demonstration purposes.
> It represents a common real-world business use case.

---

## 🧩 Problem
Invoices are often prepared manually in Excel.
This leads to:
- Repetitive work
- Human errors
- Time wastage

---

## 💡 Solution
This project automates invoice generation using **Excel VBA**.
With a single click, invoices are generated from structured input data.

---

## ⚙️ How It Works
1. User enters invoice data in an **Input** sheet
2. VBA macro processes the data
3. Invoice template is filled automatically
4. Total amount is calculated
5. Final invoice is ready to share

---

## 🔁 Input → Process → Output
- **Input:** Customer details, item list, quantity, rate
- **Process:** VBA macro reads data and performs calculations
- **Output:** Ready-to-use invoice

---

## 🔧 Sample VBA Macro (Invoice Generation)

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

```
---

## ▶️ How to Use

1. Open Excel
2. Press **ALT + F11** to open the VBA editor
3. Click **Insert → Module**
4. Paste the macro code
5. Close the editor
6. Run the macro **GenerateInvoice**

The invoice will be generated automatically based on the input data.
