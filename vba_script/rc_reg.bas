Attribute VB_Name = "Module1"
Sub PopulateUnitFromSQLReport_ByCode()
    Dim wsData As Worksheet
    Dim wsSQL As Worksheet
    Dim lastRowData As Long
    Dim lastRowSQL As Long
    Dim dataCode As String
    Dim tenantCode As String
    Dim roommateCode As String
    Dim unitValue As String
    Dim propertyNum As String
    Dim i As Long, j As Long
    Dim matchFound As Boolean

    Set wsData = ThisWorkbook.Sheets("Rent Cafe Data")
    Set wsSQL = ThisWorkbook.Sheets("SQL Report")

    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastRowSQL = wsSQL.Cells(wsSQL.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRowData
        dataCode = Trim(wsData.Cells(i, 4).value)
        matchFound = False

        If Len(dataCode) > 0 Then
            For j = 2 To lastRowSQL
                tenantCode = Trim(wsSQL.Cells(j, 3).value)
                roommateCode = Trim(wsSQL.Cells(j, 118).value)

                If Left(dataCode, 1) = "t" Then
                    If StrComp(dataCode, tenantCode, vbTextCompare) = 0 Then
                        unitValue = wsSQL.Cells(j, 2).value          ' Column B = Unit
                        propertyNum = wsSQL.Cells(j, 1).value        ' Column A = Property Number
                        wsData.Cells(i, 12).value = "'" & unitValue  ' Column L
                        wsData.Cells(i, 13).value = propertyNum      ' Column M
                        matchFound = True
                        Exit For
                    End If
                ElseIf Left(dataCode, 1) = "r" Then
                    If Len(roommateCode) > 0 Then
                        If StrComp(dataCode, roommateCode, vbTextCompare) = 0 Then
                            unitValue = wsSQL.Cells(j, 2).value          ' Column B = Unit
                            propertyNum = wsSQL.Cells(j, 1).value        ' Column A = Property Number
                            wsData.Cells(i, 12).value = "'" & unitValue  ' Column L
                            wsData.Cells(i, 13).value = propertyNum      ' Column M
                            matchFound = True
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If

        If Not matchFound Then
            wsData.Cells(i, 12).value = "NO MATCH FOUND"
            wsData.Cells(i, 13).value = "" ' Clear property column if no match
        End If
    Next i

    MsgBox "Unit and property data population complete!", vbInformation
End Sub


Sub FilterDuplicates_KeepRegisteredInvitedUnregistered()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim unitDict As Object
    Dim unitVal As String
    Dim regStatus As Variant
    Dim rowsToDelete As Collection
    Dim group As Variant
    Dim index As Long
    Dim keepIndex As Long
    Dim priorityStatus As Variant
    Dim foundStatus As Boolean

    Set ws = ThisWorkbook.Sheets("Rent Cafe Data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set unitDict = CreateObject("Scripting.Dictionary")
    Set rowsToDelete = New Collection

    ' Step 1: Group rows by Unit value (Column L)
    For i = 2 To lastRow
        unitVal = Trim(ws.Cells(i, 12).value)
        If unitVal <> "" Then
            If Not unitDict.exists(unitVal) Then
                unitDict.Add unitVal, Array(i)
            Else
                unitDict(unitVal) = AppendToArray(unitDict(unitVal), i)
            End If
        End If
    Next i

    ' Step 2: Process each group with duplicates
    For Each group In unitDict.Keys
        If UBound(unitDict(group)) > 0 Then ' Only if it's a duplicate group
            keepIndex = -1
            foundStatus = False

            ' Priority list to check in order
            priorityStatus = Array("Registered", "Invited", "Unregistered")

            ' Check each priority in order
            For Each regStatus In priorityStatus
                For index = LBound(unitDict(group)) To UBound(unitDict(group))
                    i = unitDict(group)(index)
                    If StrComp(Trim(ws.Cells(i, 7).value), regStatus, vbTextCompare) = 0 Then
                        keepIndex = i
                        foundStatus = True
                        Exit For
                    End If
                Next index
                If foundStatus Then Exit For
            Next regStatus

            ' Mark rows to delete: all except the chosen keepIndex
            If foundStatus Then
                For index = LBound(unitDict(group)) To UBound(unitDict(group))
                    i = unitDict(group)(index)
                    If i <> keepIndex Then rowsToDelete.Add i
                Next index
            End If
        End If
    Next group

    ' Step 3: Delete rows from bottom to top
    For i = rowsToDelete.Count To 1 Step -1
        ws.Rows(rowsToDelete(i)).Delete
    Next i

    MsgBox "Filtered and kept best available status (Registered to Invited to Unregistered) for duplicates.", vbInformation
End Sub

Function AppendToArray(arr As Variant, value As Variant) As Variant
    Dim newArr() As Variant
    Dim i As Long
    ReDim newArr(0 To UBound(arr) + 1)
    For i = 0 To UBound(arr)
        newArr(i) = arr(i)
    Next i
    newArr(UBound(arr) + 1) = value
    AppendToArray = newArr
End Function

