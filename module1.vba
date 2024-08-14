
Function IdentifyColumn(ValueToFind As String) As Integer
Dim i As Integer, columnValue As String
i = 2
Do While i < 1000
    columnValue = Worksheets(1).Cells(1, i).Value
    If columnValue = "" Then
        Worksheets(1).Cells(1, i).Value = ValueToFind 'label not found, we're creating a new one
        Exit Do
    End If
    If columnValue = ValueToFind Then
     Exit Do 'column found, we're leaving
    End If
    i = i + 1
Loop

IdentifyColumn = i
End Function
Sub TestSplit()

Dim WebString As String, currentRow As Integer, i As Integer, j As Integer
currentRow = 2
WebString = "start"

Dim Label As String, Value As String, ColumnN As Integer

Do While WebString <> ""
    WebString = Worksheets(1).Cells(currentRow, 1).Value
    WebString = Replace(WebString, "&#181;", "µ")
    WebString = Replace(WebString, "&#8239;", " ")
    WebString = Replace(WebString, "&#62;", ">")
    WebString = Replace(WebString, "&gt;", ">")
    WebString = Replace(WebString, "&#8805;", ">")

    Debug.Print WebString
    Dim WebResult() As String, WebResultSplit() As String
    WebResult = Split(WebString, "<")
    i = 0
    Do While i <= UBound(WebResult)
        If i > 500 Then
            Exit Do
        End If
        WebResultSplit = Split(WebResult(i), ">")
        Label = WebResultSplit(0)
        Value = ""
        
        For j = 1 To UBound(WebResultSplit)
            Value = Value + WebResultSplit(j)
        Next j
        ColumnN = IdentifyColumn(Label)
        Worksheets(1).Cells(currentRow, ColumnN).Value = Value
        Debug.Print Label
        Debug.Print Value
        i = i + 1
    Loop
currentRow = currentRow + 1
Loop

End Sub

Sub HyperAdd()
'Converts each text hyperlink selected into a working hyperlink

Dim xCell As Range
    
    For Each xCell In Selection
        ActiveSheet.Hyperlinks.Add Anchor:=xCell, Address:=xCell.Formula
    Next xCell
End Sub
