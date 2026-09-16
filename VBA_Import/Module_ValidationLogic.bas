Attribute VB_Name = "Module_ValidationLogic"

Option Explicit

' ================================================================
' Module_ValidationLogic
' Validation update, temporary lists with column support
' ================================================================

' Updates validation for a single column
Public Sub RefreshValidationForColumn(ByVal ws As Worksheet, ByVal filterCol As Long)
    On Error GoTo CleanExit

    Debug.Print "RefreshValidationForColumn: filterCol = " & filterCol

    If filterCol < START_COL Or filterCol > MAX_COL Then Exit Sub

    Dim wasProtected As Boolean
    wasProtected = ws.ProtectContents
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect Password:="ptm"
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
    End If

    Dim curMaterial As Variant
    Dim curMountType As Variant
    Dim curInstallType As Variant
    Dim curDelivery As Variant
    Dim curModel As Variant

    curMaterial = ws.Cells(FILTER_START_ROW, filterCol).Value
    curMountType = ws.Cells(FILTER_START_ROW + 1, filterCol).Value
    curInstallType = ws.Cells(FILTER_START_ROW + 2, filterCol).Value
    curDelivery = ws.Cells(FILTER_START_ROW + 3, filterCol).Value
    curModel = ws.Cells(FILTER_END_ROW, filterCol).Value

    Debug.Print "RefreshValidationForColumn: curMaterial = " & curMaterial & ", curMountType = " & curMountType & ", curModel = " & curModel

    If Not IsEmpty(curModel) And curModel <> "" Then
        If (IsEmpty(curMaterial) Or curMaterial = "") And _
           (IsEmpty(curMountType) Or curMountType = "") And _
           (IsEmpty(curInstallType) Or curInstallType = "") And _
           (IsEmpty(curDelivery) Or curDelivery = "") Then
            
            ' refreshvalidationforcolumn: фильтры пустые, автозаполняем по модели
            ' refreshvalidationforcolumn: filters are empty, autofill from model
            Debug.Print Ru("0052 0065 0066 0072 0065 0073 0068 0056 0061 006C 0069 0064 0061 0074 0069 006F 006E 0046 006F 0072 0043 006F 006C 0075 006D 006E 003A 0020 0444 0438 043B 044C 0442 0440 044B") & Ru("0020 043F 0443 0441 0442 044B 0435 002C 0020 0430 0432 0442 043E 0437 0430 043F 043E 043B 043D 044F 0435 043C 0020 043F 043E 0020 043C 043E 0434 0435 043B 0438")
            Call AutoFillFiltersFromModel(ws, filterCol, curModel)
            
            curMaterial = ws.Cells(FILTER_START_ROW, filterCol).Value
            curMountType = ws.Cells(FILTER_START_ROW + 1, filterCol).Value
            curInstallType = ws.Cells(FILTER_START_ROW + 2, filterCol).Value
            curDelivery = ws.Cells(FILTER_START_ROW + 3, filterCol).Value
        End If
    End If

    ' CP validation
    Dim cpListExists As Boolean
    Dim cpTestRng As Range
    On Error Resume Next
    Set cpTestRng = ws.names("cp_List").RefersToRange
    If Err.Number = 0 Then
        cpListExists = True
    Else
        cpListExists = False
        Err.Clear
        On Error Resume Next
        Call UpdateCPValidationList
        If Err.Number = 0 Then
            Set cpTestRng = ws.names("cp_List").RefersToRange
            If Err.Number = 0 Then cpListExists = True
        End If
        Err.Clear
    End If
    On Error GoTo 0

    If cpListExists Then
        On Error Resume Next
        With ws.Cells(CP_MODEL_ROW, filterCol).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=cp_List"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

    ' Filter lists (rows 34-38)
    Dim materialList As String
    materialList = Module_ValidationLists.GetMaterialListCross(curMountType, curInstallType, curDelivery, curModel)
    If materialList = "" Then materialList = Module_ValidationLists.GetAllMaterials()
    Debug.Print "RefreshValidationForColumn: materialList = " & materialList
    Call UpdateTempList(ws, "TempMaterialList", materialList, filterCol)
    Call SetValidationWithRange(ws.Cells(FILTER_START_ROW, filterCol), "TempMaterialList", filterCol)

    Dim mountTypeList As String
    mountTypeList = Module_ValidationLists.GetMountTypeListCross(curMaterial, curInstallType, curDelivery, curModel)
    If mountTypeList = "" Then mountTypeList = Module_ValidationLists.GetAllMountTypes()
    Debug.Print "RefreshValidationForColumn: mountTypeList = " & mountTypeList
    Call UpdateTempList(ws, "TempMountList", mountTypeList, filterCol)
    Call SetValidationWithRange(ws.Cells(FILTER_START_ROW + 1, filterCol), "TempMountList", filterCol)

    Dim installTypeList As String
    installTypeList = Module_ValidationLists.GetInstallTypeListFromMountType(curMountType)
    If installTypeList = "" Then installTypeList = Module_ValidationLists.GetAllInstallTypes()
    Debug.Print "RefreshValidationForColumn: installTypeList = " & installTypeList
    Call UpdateTempList(ws, "TempInstallList", installTypeList, filterCol)
    Call SetValidationWithRange(ws.Cells(FILTER_START_ROW + 2, filterCol), "TempInstallList", filterCol)

    Dim deliveryList As String
    deliveryList = Module_ValidationLists.GetDeliveryListCross(curMaterial, curMountType, curInstallType, curModel)
    If deliveryList = "" Then deliveryList = Module_ValidationLists.GetAllDeliveries()
    Debug.Print "RefreshValidationForColumn: deliveryList = " & deliveryList
    Call UpdateTempList(ws, "TempDeliveryList", deliveryList, filterCol)
    Call SetValidationWithRange(ws.Cells(FILTER_START_ROW + 3, filterCol), "TempDeliveryList", filterCol)

    Dim modelList As String
    modelList = Module_ValidationLists.GetModelListCross(curMaterial, curMountType, curInstallType, curDelivery)
    If modelList = "" Then modelList = Module_ValidationLists.GetAllAGModels()
    Debug.Print "RefreshValidationForColumn: modelList = " & modelList
    Call UpdateTempList(ws, "TempModelList", modelList, filterCol)
    Call SetValidationWithRange(ws.Cells(FILTER_END_ROW, filterCol), "TempModelList", filterCol)

    Call Module_Visual.ApplyColorToInstallationType(ws, filterCol)

    If wasProtected Then
        On Error Resume Next
        ws.Protect Password:="ptm", UserInterfaceOnly:=True
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

CleanExit:
End Sub

' Updates temporary list with column-specific name
Private Sub UpdateTempList(ByVal ws As Worksheet, ByVal rangeName As String, _
                           ByVal listString As String, ByVal columnIndex As Long)
    On Error GoTo CleanExit

    Dim fullName As String
    fullName = rangeName & "_" & columnIndex

    If listString = "" Then
        On Error Resume Next
        ws.names(fullName).Delete
        On Error GoTo 0
        Exit Sub
    End If

    Dim items As Variant
    items = Split(listString, ",")
    Dim cleanItems() As String
    Dim cnt As Long
    Dim i As Long
    cnt = 0
    For i = 0 To UBound(items)
        If Trim(items(i)) <> "" Then
            ReDim Preserve cleanItems(cnt)
            cleanItems(cnt) = Trim(items(i))
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        On Error Resume Next
        ws.names(fullName).Delete
        On Error GoTo 0
        Exit Sub
    End If

    ' Determine the start row for this list type
    Dim startRow As Long
    Select Case rangeName
        Case "TempMaterialList": startRow = 200
        Case "TempMountList": startRow = 300
        Case "TempInstallList": startRow = 400
        Case "TempDeliveryList": startRow = 500
        Case "TempModelList": startRow = 600
        Case Else: Exit Sub
    End Select

    ' Clear the buffer (400 rows)
    ws.Range(ws.Cells(startRow, columnIndex), ws.Cells(startRow + 400, columnIndex)).ClearContents

    ' Write list items
    For i = 0 To cnt - 1
        ws.Cells(startRow + i, columnIndex).Value = cleanItems(i)
    Next i

    ' Create a named range
    Dim rngNew As Range
    Set rngNew = ws.Range(ws.Cells(startRow, columnIndex), ws.Cells(startRow + cnt - 1, columnIndex))
    On Error Resume Next
    ws.names(fullName).Delete
    Err.Clear
    On Error GoTo 0
    ws.names.Add name:=fullName, RefersTo:=rngNew

CleanExit:
End Sub

' Sets validation with range for a specific column
Private Sub SetValidationWithRange(ByVal targetCell As Range, ByVal rangeName As String, ByVal columnIndex As Long)
    On Error GoTo CleanExit
    If targetCell Is Nothing Then Exit Sub

    Dim fullName As String
    fullName = rangeName & "_" & columnIndex
    Debug.Print "SetValidationWithRange: " & targetCell.address & " -> " & fullName

    Dim nm As name
    On Error Resume Next
    Set nm = targetCell.Worksheet.names(fullName)
    If Err.Number <> 0 Then
        On Error GoTo 0
        On Error Resume Next
        targetCell.Validation.Delete
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Dim rng As Range
    Set rng = nm.RefersToRange
    If rng Is Nothing Or Application.WorksheetFunction.CountA(rng) = 0 Then
        On Error Resume Next
        targetCell.Validation.Delete
        On Error GoTo 0
        Exit Sub
    End If

    On Error Resume Next
    targetCell.Validation.Delete
    On Error GoTo 0

    On Error Resume Next
    With targetCell.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & fullName
        If Err.Number = 0 Then
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        Else
            Err.Clear
        End If
    End With
    On Error GoTo 0

CleanExit:
End Sub

' Updates all cross filters for a single column
Public Sub UpdateAllFiltersCross(ByVal ws As Worksheet, ByVal targetCol As Long)
    On Error GoTo CleanExit

    Debug.Print "UpdateAllFiltersCross: targetCol = " & targetCol

    Dim curMaterial As Variant, curMountType As Variant, curInstallType As Variant
    Dim curDelivery As Variant, curModel As Variant

    curMaterial = ws.Cells(FILTER_START_ROW, targetCol).Value
    curMountType = ws.Cells(FILTER_START_ROW + 1, targetCol).Value
    curInstallType = ws.Cells(FILTER_START_ROW + 2, targetCol).Value
    curDelivery = ws.Cells(FILTER_START_ROW + 3, targetCol).Value
    curModel = ws.Cells(FILTER_END_ROW, targetCol).Value

    Dim savedModel As Variant
    savedModel = curModel

    Dim materialList As String
    materialList = Module_ValidationLists.GetMaterialListCross(curMountType, curInstallType, curDelivery, curModel)
    If materialList <> "" Then
        Call UpdateTempList(ws, "TempMaterialList", materialList, targetCol)
        Call SetValidationWithRange(ws.Cells(FILTER_START_ROW, targetCol), "TempMaterialList", targetCol)
    Else
        Call ClearValidation(ws.Cells(FILTER_START_ROW, targetCol))
    End If

    Dim mountTypeList As String
    mountTypeList = Module_ValidationLists.GetMountTypeListCross(curMaterial, curInstallType, curDelivery, curModel)
    If mountTypeList <> "" Then
        Call UpdateTempList(ws, "TempMountList", mountTypeList, targetCol)
        Call SetValidationWithRange(ws.Cells(FILTER_START_ROW + 1, targetCol), "TempMountList", targetCol)
    Else
        Call ClearValidation(ws.Cells(FILTER_START_ROW + 1, targetCol))
    End If

    Dim installTypeList As String
    installTypeList = Module_ValidationLists.GetInstallTypeListFromMountType(curMountType)
    If installTypeList <> "" Then
        Call UpdateTempList(ws, "TempInstallList", installTypeList, targetCol)
        Call SetValidationWithRange(ws.Cells(FILTER_START_ROW + 2, targetCol), "TempInstallList", targetCol)
    Else
        Call ClearValidation(ws.Cells(FILTER_START_ROW + 2, targetCol))
    End If

    Dim deliveryList As String
    deliveryList = Module_ValidationLists.GetDeliveryListCross(curMaterial, curMountType, curInstallType, curModel)
    If deliveryList <> "" Then
        Call UpdateTempList(ws, "TempDeliveryList", deliveryList, targetCol)
        Call SetValidationWithRange(ws.Cells(FILTER_START_ROW + 3, targetCol), "TempDeliveryList", targetCol)
    Else
        Call ClearValidation(ws.Cells(FILTER_START_ROW + 3, targetCol))
    End If

    Dim modelList As String
    modelList = Module_ValidationLists.GetModelListCross(curMaterial, curMountType, curInstallType, curDelivery)
    If modelList <> "" Then
        Call UpdateTempList(ws, "TempModelList", modelList, targetCol)
        Call SetValidationWithRange(ws.Cells(FILTER_END_ROW, targetCol), "TempModelList", targetCol)
    Else
        Call ClearValidation(ws.Cells(FILTER_END_ROW, targetCol))
    End If

CleanExit:
End Sub

' Sets simple validation list (comma-separated string)
Private Sub SetValidationList(ByVal targetCell As Range, ByVal listSource As String)
    On Error Resume Next
    If targetCell Is Nothing Or listSource = "" Then
        Call ClearValidation(targetCell)
        Exit Sub
    End If
    
    targetCell.Validation.Delete
    With targetCell.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=listSource
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
End Sub

' Clears validation from a cell
Private Sub ClearValidation(ByVal targetCell As Range)
    On Error Resume Next
    If Not targetCell Is Nothing Then
        targetCell.Validation.Delete
    End If
End Sub

' Checks if value exists in comma-separated list
Private Function ValueExistsInList(ByVal Value As Variant, ByVal list As String) As Boolean
    If IsEmpty(Value) Or Value = "" Or list = "" Then
        ValueExistsInList = False
        Exit Function
    End If
    
    Dim items As Variant
    items = Split(list, ",")
    Dim i As Long
    Dim valStr As String
    valStr = CleanStringForTable(CStr(Value))
    
    For i = LBound(items) To UBound(items)
        If CleanStringForTable(items(i)) = valStr Then
            ValueExistsInList = True
            Exit Function
        End If
    Next i
    
    ValueExistsInList = False
End Function

' Forces refresh of validation for all columns
Public Sub ForceRefreshAllValidation(ByVal ws As Worksheet)
    On Error GoTo CleanExit
    
    Dim colBr As Long
    Dim pipeCountValue As Variant
    
    pipeCountValue = SafeRangeValue(ws.Range("pipeCountCP"))
    colBr = SafeNumericValue(pipeCountValue, 1)
    
    Dim colIdx As Long
    For colIdx = START_COL To START_COL + colBr - 1
        Call RefreshValidationForColumn(ws, colIdx)
        If colIdx Mod 5 = 0 Then DoEvents
    Next colIdx

CleanExit:
End Sub

' Auto-fills filters from model
Public Sub AutoFillFiltersFromModel(ByVal ws As Worksheet, ByVal col As Long, ByVal modelValue As String)
    On Error GoTo CleanExit
    
    Debug.Print "AutoFillFiltersFromModel: col = " & col & ", modelValue = " & modelValue
    
    If modelValue = "" Or IsEmpty(modelValue) Then Exit Sub
    
    Dim rngModel As Range
    On Error Resume Next
    Set rngModel = thisWorkbook.names("AG_Model").RefersToRange
    On Error GoTo 0
    
    If rngModel Is Nothing Then Exit Sub
    
    Dim modelVals As Variant
    modelVals = rngModel.Value
    If Not IsArray(modelVals) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = modelVals
        modelVals = arr
    End If
    
    Dim modelClean As String
    modelClean = CleanStringForTable(CStr(modelValue))
    
    Dim foundRow As Long
    foundRow = 0
    Dim i As Long
    For i = 1 To UBound(modelVals, 1)
        Dim valClean As String
        valClean = CleanStringForTable(CStr(modelVals(i, 1)))
        If valClean = modelClean Then
            foundRow = i
            Exit For
        End If
    Next i
    
    If foundRow = 0 Then
        Call Module_AGData.ClearCalculatedDataForCol(ws, col)
        Exit Sub
    End If
    
    Dim material As Variant
    Dim mountType As Variant
    Dim completion As Variant
    
    Dim rngMat As Range
    Dim rngMount As Range
    Dim rngShip As Range
    
    On Error Resume Next
    Set rngMat = thisWorkbook.names("AG_Material").RefersToRange
    Set rngMount = thisWorkbook.names("AG_MountType").RefersToRange
    Set rngShip = thisWorkbook.names("AG_Completion").RefersToRange
    On Error GoTo 0
    
    If Not rngMat Is Nothing Then material = rngMat.Cells(foundRow, 1).Value
    If Not rngMount Is Nothing Then mountType = rngMount.Cells(foundRow, 1).Value
    If Not rngShip Is Nothing Then completion = rngShip.Cells(foundRow, 1).Value
    
    If Not IsEmpty(material) And material <> "" Then
        If IsEmpty(ws.Cells(FILTER_START_ROW, col).Value) Or ws.Cells(FILTER_START_ROW, col).Value = "" Then
            ws.Cells(FILTER_START_ROW, col).Value = material
        End If
    End If
    
    If Not IsEmpty(mountType) And mountType <> "" Then
        If IsEmpty(ws.Cells(FILTER_START_ROW + 1, col).Value) Or ws.Cells(FILTER_START_ROW + 1, col).Value = "" Then
            ws.Cells(FILTER_START_ROW + 1, col).Value = mountType
        End If
    End If
    
    If Not IsEmpty(mountType) And mountType <> "" Then
        Dim installType As String
        installType = Module_ValidationLists.GetInstallationTypeFromMountType(mountType)
        If installType <> "" Then
            If IsEmpty(ws.Cells(FILTER_START_ROW + 2, col).Value) Or ws.Cells(FILTER_START_ROW + 2, col).Value = "" Then
                ws.Cells(FILTER_START_ROW + 2, col).Value = installType
            End If
        End If
    End If
    
    If Not IsEmpty(completion) And completion <> "" Then
        If IsEmpty(ws.Cells(FILTER_START_ROW + 3, col).Value) Or ws.Cells(FILTER_START_ROW + 3, col).Value = "" Then
            ws.Cells(FILTER_START_ROW + 3, col).Value = completion
        End If
    End If
    
    If IsEmpty(ws.Cells(FILTER_END_ROW, col).Value) Or ws.Cells(FILTER_END_ROW, col).Value = "" Then
        ws.Cells(FILTER_END_ROW, col).Value = modelValue
    End If
    
    Call RefreshValidationForColumn(ws, col)
    Call Module_AGData.InsertAGDataForColumn(ws, col)

CleanExit:
End Sub
