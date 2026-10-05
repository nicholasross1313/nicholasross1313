Attribute VB_Name = "MultiFileReconciliation"
Option Explicit

Public Sub RunMultiFileReconciliation()

    Dim wsOrders As Worksheet
    Dim wsWarehouse As Worksheet
    Dim wsBilling As Worksheet
    Dim wsResult As Worksheet

    Set wsOrders = GetRequiredSheet("Orders")
    Set wsWarehouse = GetRequiredSheet("Warehouse")
    Set wsBilling = GetRequiredSheet("Billing")
    Set wsResult = GetOrCreateSheet("Reconciliation Results")

    Dim orders As Object
    Dim warehouse As Object
    Dim billing As Object
    Dim allKeys As Object

    Set orders = LoadRows(wsOrders, "OrderID")
    Set warehouse = LoadRows(wsWarehouse, "OrderID")
    Set billing = LoadRows(wsBilling, "OrderID")
    Set allKeys = CreateObject("Scripting.Dictionary")

    AddKeys allKeys, orders
    AddKeys allKeys, warehouse
    AddKeys allKeys, billing

    wsResult.Cells.Clear
    WriteHeaders wsResult

    Dim rowOut As Long
    rowOut = 2

    Dim key As Variant
    For Each key In allKeys.Keys
        WriteResultRow wsResult, rowOut, CStr(key), orders, warehouse, billing
        rowOut = rowOut + 1
    Next key

    FormatResult wsResult, rowOut - 1

    MsgBox "Reconciliation complete: " & (rowOut - 2) & " records checked.", vbInformation

End Sub

Private Function GetRequiredSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetRequiredSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetRequiredSheet Is Nothing Then
        Err.Raise vbObjectError + 100, , "Missing required sheet: " & sheetName
    End If
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function LoadRows(ByVal ws As Worksheet, ByVal keyHeader As String) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim col As Long
    For col = 1 To lastCol
        If Len(Trim$(CStr(ws.Cells(1, col).Value))) > 0 Then
            headerMap(Trim$(CStr(ws.Cells(1, col).Value))) = col
        End If
    Next col

    If Not headerMap.Exists(keyHeader) Then
        Err.Raise vbObjectError + 101, , ws.Name & " is missing header: " & keyHeader
    End If

    Dim keyCol As Long
    keyCol = headerMap(keyHeader)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow

        Dim recordKey As String
        recordKey = Trim$(CStr(ws.Cells(r, keyCol).Value))

        If Len(recordKey) > 0 Then

            Dim rowData As Object
            Set rowData = CreateObject("Scripting.Dictionary")

            Dim header As Variant
            For Each header In headerMap.Keys
                rowData(CStr(header)) = ws.Cells(r, headerMap(header)).Value
            Next header

            dict(recordKey) = rowData

        End If
    Next r

    Set LoadRows = dict

End Function

Private Sub AddKeys(ByVal target As Object, ByVal source As Object)
    Dim key As Variant
    For Each key In source.Keys
        If Not target.Exists(CStr(key)) Then
            target(CStr(key)) = True
        End If
    Next key
End Sub

Private Sub WriteHeaders(ByVal ws As Worksheet)

    Dim headers As Variant
    headers = Array( _
        "OrderID", _
        "Customer", _
        "SKU", _
        "OrderedQty", _
        "ShippedQty", _
        "BilledQty", _
        "ExpectedAmount", _
        "InvoiceAmount", _
        "WarehouseStatus", _
        "Result", _
        "Notes" _
    )

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

Private Sub WriteResultRow( _
    ByVal ws As Worksheet, _
    ByVal rowOut As Long, _
    ByVal orderId As String, _
    ByVal orders As Object, _
    ByVal warehouse As Object, _
    ByVal billing As Object)

    Dim hasOrder As Boolean
    Dim hasWarehouse As Boolean
    Dim hasBilling As Boolean

    hasOrder = orders.Exists(orderId)
    hasWarehouse = warehouse.Exists(orderId)
    hasBilling = billing.Exists(orderId)

    Dim customer As String
    Dim sku As String
    Dim orderedQty As Double
    Dim shippedQty As Double
    Dim billedQty As Double
    Dim unitPrice As Double
    Dim expectedAmount As Double
    Dim invoiceAmount As Double
    Dim warehouseStatus As String

    If hasOrder Then
        customer = SafeText(orders(orderId), "Customer")
        sku = SafeText(orders(orderId), "SKU")
        orderedQty = SafeNumber(orders(orderId), "OrderedQty")
        unitPrice = SafeNumber(orders(orderId), "UnitPrice")
        expectedAmount = orderedQty * unitPrice
    End If

    If hasWarehouse Then
        shippedQty = SafeNumber(warehouse(orderId), "ShippedQty")
        warehouseStatus = SafeText(warehouse(orderId), "WarehouseStatus")
    End If

    If hasBilling Then
        billedQty = SafeNumber(billing(orderId), "BilledQty")
        invoiceAmount = SafeNumber(billing(orderId), "InvoiceAmount")
    End If

    Dim result As String
    Dim notes As String

    If Not hasOrder Then
        result = "Extra downstream record"
        notes = "Order ID exists outside the Orders source."
    ElseIf Not hasWarehouse Then
        result = "Missing warehouse record"
        notes = "No warehouse record found."
    ElseIf Not hasBilling Then
        result = "Missing billing record"
        notes = "No billing record found."
    ElseIf SafeText(orders(orderId), "SKU") <> SafeText(warehouse(orderId), "SKU") Then
        result = "SKU mismatch"
        notes = "Order SKU and warehouse SKU are different."
    ElseIf orderedQty <> shippedQty Or orderedQty <> billedQty Then
        result = "Quantity mismatch"
        notes = "Ordered, shipped and billed quantities do not agree."
    ElseIf Abs(expectedAmount - invoiceAmount) > 0.01 Then
        result = "Amount mismatch"
        notes = "Invoice amount differs from ordered quantity × unit price."
    Else
        result = "Matched"
        notes = ""
    End If

    ws.Cells(rowOut, 1).Value = orderId
    ws.Cells(rowOut, 2).Value = customer
    ws.Cells(rowOut, 3).Value = sku
    ws.Cells(rowOut, 4).Value = IIf(hasOrder, orderedQty, "")
    ws.Cells(rowOut, 5).Value = IIf(hasWarehouse, shippedQty, "")
    ws.Cells(rowOut, 6).Value = IIf(hasBilling, billedQty, "")
    ws.Cells(rowOut, 7).Value = IIf(hasOrder, expectedAmount, "")
    ws.Cells(rowOut, 8).Value = IIf(hasBilling, invoiceAmount, "")
    ws.Cells(rowOut, 9).Value = warehouseStatus
    ws.Cells(rowOut, 10).Value = result
    ws.Cells(rowOut, 11).Value = notes

End Sub

Private Function SafeText(ByVal rowData As Object, ByVal fieldName As String) As String
    If rowData.Exists(fieldName) Then
        SafeText = Trim$(CStr(rowData(fieldName)))
    Else
        SafeText = ""
    End If
End Function

Private Function SafeNumber(ByVal rowData As Object, ByVal fieldName As String) As Double
    If rowData.Exists(fieldName) And IsNumeric(rowData(fieldName)) Then
        SafeNumber = CDbl(rowData(fieldName))
    Else
        SafeNumber = 0
    End If
End Function

Private Sub FormatResult(ByVal ws As Worksheet, ByVal lastRow As Long)

    ws.Rows(1).Font.Bold = True
    ws.Columns("A:K").AutoFit
    ws.Columns("G:H").NumberFormat = "0.00"

    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, 10).Value <> "Matched" Then
            ws.Cells(r, 10).Font.Bold = True
        End If
    Next r

    ws.Range("A1:K" & lastRow).AutoFilter

End Sub
