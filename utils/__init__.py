
Months = {
    "Jan": "01",
    "Fev": "02",
    "Mar": "03",
    "Abr": "04",
    "Mai": "05",
    "Jun": "06",
    "Jul": "07",
    "Ago": "08",
    "Set": "09",
    "Out": "10",
    "Nov": "11",
    "Dez": "12"
}

Ufs = {
'ACRE':'AC',
'ALAGOAS':'AL',
'AMAPÁ':'AP',
'AMAZONAS':'AM',
'BAHIA':'BA',
'CEARÁ':'CE',
'DISTRITO FEDERAL':'DF',
'ESPÍRITO SANTO':'ES',
'GOIÁS':'GO',
'MARANHÃO':'MA',
'MATO GROSSO':'MT',
'MATO GROSSO DO SUL':'MS',
'MINAS GERAIS':'MG',
'PARÁ':'PA',
'PARAÍBA':'PB',
'PARANÁ':'PR',
'PERNAMBUCO':'PE',
'PIAUÍ':'PI',
'RIO DE JANEIRO':'RJ',
'RIO GRANDE DO NORTE':'RN',
'RIO GRANDE DO SUL':'RS',
'RONDÔNIA':'RO',
'RORAIMA':'RR',
'SANTA CATARINA':'SC',
'SÃO PAULO':'SP',
'SERGIPE':'SE',
'TOCANTINS':'TO'
}

Sub_Detail_Pivot_tables ="""
Sub {}()
'lists all pivot tables in
' active workbook
'use the Notes column to
' add comments about fields
Dim lRow As Long
Dim Ws As Worksheet
Dim wsList As Worksheet
Dim pt As PivotTable
Dim pf As PivotField
Dim df As PivotField
Dim pi As PivotItem
Dim lLoc As Long
Dim lPos As Long
Dim pfCount As Long
Dim myList As ListObject
Dim bOLAP As Boolean
Application.DisplayAlerts = False

On Error GoTo errHandler

Set wsList = Sheets.Add
wsList.Name = "{}"

lRow = 2

With wsList
  .Cells(1, 1).Value = "Sheet"
  .Cells(1, 2).Value = "PT Name"
  .Cells(1, 3).Value = "PT Address"
  .Cells(1, 4).Value = "Caption"
  .Cells(1, 5).Value = "Heading"
  .Cells(1, 6).Value = "Source Name"
  .Cells(1, 7).Value = "Location"
  .Cells(1, 8).Value = "Position"
  .Cells(1, 9).Value = "Sample Item"
  .Cells(1, 10).Value = "Formula"
  .Cells(1, 11).Value = "OLAP"
  .Rows(1).Font.Bold = True
  
  For Each Ws In ActiveWorkbook.Worksheets
    For Each pt In Ws.PivotTables
      bOLAP = pt.PivotCache.OLAP
      
      For pfCount = 1 To pt.RowFields.Count
        Set pf = pt.RowFields(pfCount)
        lLoc = pf.Orientation
        If pf.Caption <> "Values" Then
        .Cells(lRow, 1).Value = Ws.Name
        .Cells(lRow, 2).Value = pt.Name
        .Cells(lRow, 3).Value = pt.TableRange2.Address
        .Cells(lRow, 4).Value = pf.Caption
        .Cells(lRow, 5).Value = pf.LabelRange.Address
        '.Cells(lRow, 6).Value = pf.SourceName
        .Cells(lRow, 7).Value = lLoc & " - Row"
        .Cells(lRow, 8).Value = pfCount
          On Error Resume Next
          If pf.PivotItems.Count > 0 _
            And bOLAP = False Then
            .Cells(lRow, 9).Value _
                = pf.PivotItems(1).Value
          End If
          On Error GoTo errHandler
        .Cells(lRow, 11).Value = bOLAP
          lRow = lRow + 1
          lLoc = 0
        End If
      Next pfCount
      
      For pfCount = 1 To pt.ColumnFields.Count
        Set pf = pt.ColumnFields(pfCount)
        lLoc = pf.Orientation
        If pf.Caption <> "Values" Then
        .Cells(lRow, 1).Value = Ws.Name
        .Cells(lRow, 2).Value = pt.Name
        .Cells(lRow, 3).Value = pt.TableRange2.Address
        .Cells(lRow, 4).Value = pf.Caption
        .Cells(lRow, 5).Value = pf.LabelRange.Address
        .Cells(lRow, 6).Value = pf.SourceName
        .Cells(lRow, 7).Value = lLoc & " - Column"
        .Cells(lRow, 8).Value = pfCount
          On Error Resume Next
          If pf.PivotItems.Count > 0 _
            And bOLAP = False Then
            .Cells(lRow, 9).Value _
                = pf.PivotItems(1).Value
          End If
          On Error GoTo errHandler
        .Cells(lRow, 11).Value = bOLAP
          lRow = lRow + 1
          lLoc = 0
        End If
      Next pfCount
      
      For pfCount = 1 To pt.PageFields.Count
        Set pf = pt.PageFields(pfCount)
        lLoc = pf.Orientation
        .Cells(lRow, 1).Value = Ws.Name
        .Cells(lRow, 2).Value = pt.Name
        .Cells(lRow, 3).Value = pt.TableRange2.Address
        .Cells(lRow, 4).Value = pf.Caption
        .Cells(lRow, 5).Value = pf.LabelRange.Address
        .Cells(lRow, 6).Value = pf.SourceName
        .Cells(lRow, 7).Value = lLoc & " - Filter"
        .Cells(lRow, 8).Value = pfCount
        On Error Resume Next
          If pf.PivotItems.Count > 0 _
            And bOLAP = False Then
          .Cells(lRow, 9).Value _
              = pf.PivotItems(1).Value
        End If
        On Error GoTo errHandler
        .Cells(lRow, 11).Value = bOLAP
        lRow = lRow + 1
        lLoc = 0
      Next pfCount
      
      For pfCount = 1 To pt.DataFields.Count
        Set pf = pt.DataFields(pfCount)
        lLoc = pf.Orientation
        
        Set df = pt.PivotFields(pf.SourceName)
        .Cells(lRow, 1).Value = Ws.Name
        .Cells(lRow, 2).Value = pt.Name
        .Cells(lRow, 3).Value = pt.TableRange2.Address
        .Cells(lRow, 4).Value = df.Caption
        .Cells(lRow, 5).Value = _
              pf.LabelRange.Cells(1).Address
       .Cells(lRow, 6).Value = df.SourceName
        .Cells(lRow, 7).Value = lLoc & " - Data"

        .Cells(lRow, 8).Value = pfCount
        'sample data not shown for value fields
        On Error Resume Next
          'print formula for calculated fields
          '.Cells(lRow, 6).Value = " " & pf.Formula
            If df.IsCalculated = True Then
              .Cells(lRow, 10).Value = _
                  Right(df.Formula, Len(df.Formula) - 1)
            End If
        On Error GoTo errHandler
         .Cells(lRow, 11).Value = bOLAP
       lRow = lRow + 1
        lLoc = 0
        Set df = Nothing
      Next pfCount
            
    Next pt
  Next Ws
  .Columns("A:K").EntireColumn.AutoFit
  Set myList = .ListObjects.Add(xlSrcRange, _
      Range("A1").CurrentRegion)

End With

exitHandler:
    Application.DisplayAlerts = True
    Exit Sub
errHandler:
    Resume exitHandler

End Sub

"""

sub_expand_sales ="""
Sub {subname}(SheetName As String, position As String)
    Dim Year As Integer
    Dim Ws As Worksheet    

    Sheets("{mainsheet}").Select
    Sheets("{mainsheet}").range(position).Select

    Selection.ShowDetail = True
    Set Ws = ActiveSheet
    Ws.name = SheetName

End Sub
"""