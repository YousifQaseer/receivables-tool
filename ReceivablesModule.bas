
Attribute VB_Name = "ReceivablesModule"

Option Explicit

Public Sub AddReceivable()

    Dim ws As Worksheet
    Dim nextRow As Long
    Dim customerName As String
    Dim amountDue As Double
    Dim dueDate As Date

    Set ws = ThisWorkbook.Sheets("Receivables")

    customerName = InputBox("Enter customer name:")
    amountDue = InputBox("Enter amount due:")
    dueDate = InputBox("Enter due date (MM/DD/YYYY):")

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = customerName
    ws.Cells(nextRow, 2).Value = amountDue
    ws.Cells(nextRow, 3).Value = dueDate
    ws.Cells(nextRow, 4).Value = "Pending"

    MsgBox "Receivable added successfully!", vbInformation

End Sub


Public Sub CalculateTotals()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalBalance As Double
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Receivables")

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    For i = 2 To lastRow
        totalBalance = totalBalance + ws.Cells(i, 2).Value
    Next i

    MsgBox "Total Outstanding Balance: $" & totalBalance, vbInformation

End Sub
