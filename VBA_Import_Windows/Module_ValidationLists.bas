Attribute VB_Name = "Module_ValidationLists"

Option Explicit

' ================================================================
' Module_ValidationLists
' Functions for obtaining unique lists and cross-filtering
' ================================================================

' Returns list of unique materials
Public Function GetAllMaterials() As String
    On Error Resume Next
    Dim rng As Range
    Set rng = thisWorkbook.names("AG_Material").RefersToRange
    If rng Is Nothing Then Exit Function
    
    Dim vals As Variant
    vals = rng.Value
    If Not IsArray(vals) Then
        GetAllMaterials = CStr(vals)
        Exit Function
    End If
    
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            exists = False
            For Each item In col
                If item = val Then exists = True: Exit For
            Next item
            If Not exists Then col.Add val
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetAllMaterials = result
End Function

' Returns list of unique mounting types
Public Function GetAllMountTypes() As String
    On Error Resume Next
    Dim rng As Range
    Set rng = thisWorkbook.names("AG_MountType").RefersToRange
    If rng Is Nothing Then Exit Function
    
    Dim vals As Variant
    vals = rng.Value
    If Not IsArray(vals) Then
        GetAllMountTypes = CStr(vals)
        Exit Function
    End If
    
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            exists = False
            For Each item In col
                If item = val Then exists = True: Exit For
            Next item
            If Not exists Then col.Add val
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetAllMountTypes = result
End Function

' Returns list of unique installation types
Public Function GetAllInstallTypes() As String
    On Error Resume Next
    Dim wsMounting As Worksheet
    Dim tblMounting As ListObject
    Dim installVals As Variant
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    Set wsMounting = thisWorkbook.Worksheets("ListAG")
    If wsMounting Is Nothing Then
        Set wsMounting = thisWorkbook.Worksheets(SHEET_LIST_AG)
    End If
    If wsMounting Is Nothing Then Exit Function
    
    Set tblMounting = wsMounting.ListObjects("TableMounting")
    If tblMounting Is Nothing Then Exit Function
    
    Dim installColNum As Long
    installColNum = 3   ' installation method
    installVals = tblMounting.ListColumns(installColNum).DataBodyRange.Value
    
    If Not IsArray(installVals) Then
        GetAllInstallTypes = CStr(installVals)
        Exit Function
    End If
    
    For i = 1 To UBound(installVals, 1)
        val = installVals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            exists = False
            For Each item In col
                If item = val Then exists = True: Exit For
            Next item
            If Not exists Then col.Add val
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetAllInstallTypes = result
End Function

' Returns list of unique deliveries
Public Function GetAllDeliveries() As String
    On Error Resume Next
    Dim rng As Range
    Set rng = thisWorkbook.names("AG_Completion").RefersToRange
    If rng Is Nothing Then Exit Function
    
    Dim vals As Variant
    vals = rng.Value
    If Not IsArray(vals) Then
        GetAllDeliveries = CStr(vals)
        Exit Function
    End If
    
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            exists = False
            For Each item In col
                If item = val Then exists = True: Exit For
            Next item
            If Not exists Then col.Add val
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetAllDeliveries = result
End Function

' Returns list of unique AG models
Public Function GetAllAGModels() As String
    On Error Resume Next
    Dim rng As Range
    Set rng = thisWorkbook.names("AG_Model").RefersToRange
    If rng Is Nothing Then Exit Function
    
    Dim vals As Variant
    vals = rng.Value
    If Not IsArray(vals) Then
        GetAllAGModels = CStr(vals)
        Exit Function
    End If
    
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            exists = False
            For Each item In col
                If item = val Then exists = True: Exit For
            Next item
            If Not exists Then col.Add val
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetAllAGModels = result
End Function

' Returns list of installation types that depend only on mounting type
Public Function GetInstallTypeListFromMountType(ByVal mountType As String) As String
    On Error GoTo CleanExit
    
    Dim wsMounting As Worksheet
    Dim tblMounting As ListObject
    Dim typeColNum As Long, installColNum As Long
    Dim mountVals As Variant, installVals As Variant
    Dim i As Long
    Dim col As New Collection
    Dim result As String
    Dim item As Variant
    
    Set wsMounting = thisWorkbook.Worksheets("ListAG")
    If wsMounting Is Nothing Then
        Set wsMounting = thisWorkbook.Worksheets(SHEET_LIST_AG)
    End If
    If wsMounting Is Nothing Then Exit Function
    
    Set tblMounting = wsMounting.ListObjects("TableMounting")
    If tblMounting Is Nothing Then Exit Function
    
    Dim headerRow As Range
    Set headerRow = tblMounting.HeaderRowRange
    Dim headerCell As Range
    
    typeColNum = 4   ' mounting type
    installColNum = 3 ' installation method
    
    mountVals = tblMounting.ListColumns(typeColNum).DataBodyRange.Value
    installVals = tblMounting.ListColumns(installColNum).DataBodyRange.Value
    
    If Not IsArray(mountVals) Then
        ReDim arr(1 To 1, 1 To 1): arr(1, 1) = mountVals: mountVals = arr
    End If
    If Not IsArray(installVals) Then
        ReDim arr(1 To 1, 1 To 1): arr(1, 1) = installVals: installVals = arr
    End If
    
    Dim cleanMount As String
    cleanMount = CleanStringForTable(CStr(mountType))
    
    For i = 1 To UBound(mountVals, 1)
        Dim mVal As Variant, iVal As Variant
        mVal = mountVals(i, 1)
        iVal = installVals(i, 1)
        If Not IsError(mVal) And Not IsError(iVal) And Not IsEmpty(iVal) Then
            If cleanMount = "" Or InStr(1, CleanStringForTable(CStr(mVal)), cleanMount, vbTextCompare) > 0 Then
                Dim installClean As String
                installClean = CleanStringForTable(CStr(iVal))
                If installClean <> "" Then
                    Dim exists As Boolean
                    exists = False
                    Dim it As Variant
                    For Each it In col
                        If it = installClean Then exists = True: Exit For
                    Next it
                    If Not exists Then col.Add installClean
                End If
            End If
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetInstallTypeListFromMountType = result
    Exit Function
CleanExit:
    GetInstallTypeListFromMountType = ""
End Function

' Checks if combination of mounting type and installation type exists
Public Function CheckInstallationExists(ByVal mountType As String, ByVal installType As String) As Boolean
    On Error GoTo CleanExit

    If mountType = "" Or installType = "" Then
        CheckInstallationExists = True
        Exit Function
    End If
    
    Dim wsMounting As Worksheet
    Dim tblMounting As ListObject
    Dim typeColNum As Long, installColNum As Long
    Dim mountVals As Variant, installVals As Variant
    Dim i As Long
    
    On Error Resume Next
    Set wsMounting = thisWorkbook.Worksheets("ListAG")
    If wsMounting Is Nothing Then
        Set wsMounting = thisWorkbook.Worksheets(SHEET_LIST_AG)
    End If
    On Error GoTo 0
    If wsMounting Is Nothing Then
        CheckInstallationExists = False
        Exit Function
    End If
    
    On Error Resume Next
    Set tblMounting = wsMounting.ListObjects("TableMounting")
    On Error GoTo 0
    If tblMounting Is Nothing Then
        CheckInstallationExists = False
        Exit Function
    End If
    
    Dim headerRow As Range
    Set headerRow = tblMounting.HeaderRowRange
    Dim headerCell As Range
    
typeColNum = 4   ' mounting type
installColNum = 3 ' installation method
    
    mountVals = tblMounting.ListColumns(typeColNum).DataBodyRange.Value
    installVals = tblMounting.ListColumns(installColNum).DataBodyRange.Value
    
    If Not IsArray(mountVals) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = mountVals
        mountVals = arr
    End If
    If Not IsArray(installVals) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = installVals
        installVals = arr
    End If
    
    Dim cleanMount As String, cleanInstall As String
    cleanMount = NormalizeFilterText(mountType)
    cleanInstall = NormalizeFilterText(installType)
    
    For i = 1 To UBound(mountVals, 1)
        Dim mVal As Variant, iVal As Variant
        mVal = mountVals(i, 1)
        iVal = installVals(i, 1)
        If Not IsError(mVal) And Not IsError(iVal) Then
            If NormalizeFilterText(CStr(mVal)) = cleanMount And _
                NormalizeFilterText(CStr(iVal)) = cleanInstall Then
                CheckInstallationExists = True
                Exit Function
            End If
        End If
    Next i
    
    CheckInstallationExists = False
    Exit Function
    
CleanExit:
    CheckInstallationExists = False
End Function

' Returns installation type for given mounting type
Public Function GetInstallationTypeFromMountType(ByVal mountType As String) As String
    On Error GoTo CleanExit
    
    If mountType = "" Or IsEmpty(mountType) Then
        GetInstallationTypeFromMountType = ""
        Exit Function
    End If
    
    Dim wsMounting As Worksheet
    Dim tblMounting As ListObject
    Dim typeColNum As Long, installColNum As Long
    Dim mountVals As Variant, installVals As Variant
    Dim i As Long
    
    On Error Resume Next
    Set wsMounting = thisWorkbook.Worksheets("ListAG")
    If wsMounting Is Nothing Then
        Set wsMounting = thisWorkbook.Worksheets(SHEET_LIST_AG)
    End If
    On Error GoTo 0
    If wsMounting Is Nothing Then Exit Function
    
    On Error Resume Next
    Set tblMounting = wsMounting.ListObjects("TableMounting")
    On Error GoTo 0
    If tblMounting Is Nothing Then Exit Function
    
    Dim headerRow As Range
    Set headerRow = tblMounting.HeaderRowRange
    Dim headerCell As Range
    
    typeColNum = 4   ' mounting type
    installColNum = 3 ' installation method
    
    mountVals = tblMounting.ListColumns(typeColNum).DataBodyRange.Value
    installVals = tblMounting.ListColumns(installColNum).DataBodyRange.Value
    
    If Not IsArray(mountVals) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = mountVals
        mountVals = arr
    End If
    If Not IsArray(installVals) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = installVals
        installVals = arr
    End If
    
    Dim cleanMountType As String
    cleanMountType = CleanStringForTable(mountType)
    
    For i = 1 To UBound(mountVals, 1)
        Dim mVal As Variant
        mVal = mountVals(i, 1)
        If Not IsError(mVal) And Not IsEmpty(mVal) And mVal <> "" Then
            mVal = CleanStringForTable(CStr(mVal))
            If mVal = cleanMountType Then
                Dim iVal As Variant
                iVal = installVals(i, 1)
                If Not IsError(iVal) And Not IsEmpty(iVal) And iVal <> "" Then
                    GetInstallationTypeFromMountType = CStr(iVal)
                    Exit Function
                End If
            End If
        End If
    Next i
    
CleanExit:
    If Err.Number <> 0 Then
        GetInstallationTypeFromMountType = ""
    End If
End Function

' Returns list of materials filtered by other criteria
Public Function GetMaterialListCross(ByVal mountType As Variant, ByVal installType As Variant, _
                                     ByVal delivery As Variant, ByVal model As Variant) As String
    On Error GoTo CleanExit
    
    Dim rngMat As Range, rngMount As Range, rngShip As Range, rngModel As Range
    On Error Resume Next
    Set rngMat = thisWorkbook.names("AG_Material").RefersToRange
    Set rngMount = thisWorkbook.names("AG_MountType").RefersToRange
    Set rngShip = thisWorkbook.names("AG_Completion").RefersToRange
    Set rngModel = thisWorkbook.names("AG_Model").RefersToRange
    On Error GoTo 0
    
    If rngMat Is Nothing Then Exit Function
    
    Dim matVals As Variant, mountVals As Variant, shipVals As Variant, modelVals As Variant
    matVals = rngMat.Value
    mountVals = rngMount.Value
    shipVals = rngShip.Value
    modelVals = rngModel.Value
    
    If Not IsArray(matVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = matVals: matVals = arr
    If Not IsArray(mountVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = mountVals: mountVals = arr
    If Not IsArray(shipVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = shipVals: shipVals = arr
    If Not IsArray(modelVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = modelVals: modelVals = arr
    
    Dim col As New Collection
    Dim i As Long
    Dim cleanMount As String, cleanInstall As String, cleanShip As String, cleanModel As String
    
    cleanMount = CleanStringForTable(CStr(mountType))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanShip = CleanStringForTable(CStr(delivery))
    cleanModel = CleanStringForTable(CStr(model))
    
    For i = 1 To UBound(matVals, 1)
        Dim mVal As Variant, mtVal As Variant, sVal As Variant, modVal As Variant
        mVal = matVals(i, 1): mtVal = mountVals(i, 1): sVal = shipVals(i, 1): modVal = modelVals(i, 1)
        
        If Not IsError(mVal) And Not IsEmpty(mVal) Then
            Dim matchMount As Boolean, matchShip As Boolean, matchModel As Boolean, matchInstall As Boolean
            matchMount = (cleanMount = "" Or CleanStringForTable(CStr(mtVal)) = cleanMount)
            matchShip = (cleanShip = "" Or CleanStringForTable(CStr(sVal)) = cleanShip)
            matchModel = (cleanModel = "" Or CleanStringForTable(CStr(modVal)) = cleanModel)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(CleanStringForTable(CStr(mtVal)), cleanInstall)
            
            If matchMount And matchShip And matchModel And matchInstall Then
                AddUniqueToCollection col, mVal
            End If
        End If
    Next i
    
    GetMaterialListCross = CollectionToString(col)
    Exit Function
CleanExit:
    GetMaterialListCross = ""
End Function

' Returns list of mounting types filtered by other criteria
Public Function GetMountTypeListCross(ByVal material As Variant, ByVal installType As Variant, _
                                      ByVal delivery As Variant, ByVal model As Variant) As String
    On Error GoTo CleanExit
    
    Dim rngMat As Range, rngMount As Range, rngShip As Range, rngModel As Range
    On Error Resume Next
    Set rngMat = thisWorkbook.names("AG_Material").RefersToRange
    Set rngMount = thisWorkbook.names("AG_MountType").RefersToRange
    Set rngShip = thisWorkbook.names("AG_Completion").RefersToRange
    Set rngModel = thisWorkbook.names("AG_Model").RefersToRange
    On Error GoTo 0
    
    If rngMount Is Nothing Then Exit Function
    
    Dim matVals As Variant, mountVals As Variant, shipVals As Variant, modelVals As Variant
    matVals = rngMat.Value
    mountVals = rngMount.Value
    shipVals = rngShip.Value
    modelVals = rngModel.Value
    
    If Not IsArray(matVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = matVals: matVals = arr
    If Not IsArray(mountVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = mountVals: mountVals = arr
    If Not IsArray(shipVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = shipVals: shipVals = arr
    If Not IsArray(modelVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = modelVals: modelVals = arr
    
    Dim col As New Collection
    Dim i As Long
    Dim cleanMat As String, cleanInstall As String, cleanShip As String, cleanModel As String
    
    cleanMat = CleanStringForTable(CStr(material))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanShip = CleanStringForTable(CStr(delivery))
    cleanModel = CleanStringForTable(CStr(model))
    
    For i = 1 To UBound(matVals, 1)
        Dim mVal As Variant, mtVal As Variant, sVal As Variant, modVal As Variant
        mVal = matVals(i, 1): mtVal = mountVals(i, 1): sVal = shipVals(i, 1): modVal = modelVals(i, 1)
        
        If Not IsError(mtVal) And Not IsEmpty(mtVal) Then
            Dim cleanMtVal As String
            cleanMtVal = CleanStringForTable(CStr(mtVal))
            
            Dim matchMat As Boolean, matchShip As Boolean, matchModel As Boolean, matchInstall As Boolean
            matchMat = (cleanMat = "" Or CleanStringForTable(CStr(mVal)) = cleanMat)
            matchShip = (cleanShip = "" Or CleanStringForTable(CStr(sVal)) = cleanShip)
            matchModel = (cleanModel = "" Or CleanStringForTable(CStr(modVal)) = cleanModel)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cleanMtVal, cleanInstall)
            
            If matchMat And matchShip And matchModel And matchInstall Then
                AddUniqueToCollection col, mtVal
            End If
        End If
    Next i
    
    GetMountTypeListCross = CollectionToString(col)
    Exit Function
CleanExit:
    GetMountTypeListCross = ""
End Function

' Returns list of deliveries filtered by other criteria
Public Function GetDeliveryListCross(ByVal material As Variant, ByVal mountType As Variant, _
                                     ByVal installType As Variant, ByVal model As Variant) As String
    On Error GoTo CleanExit
    
    Dim rngMat As Range, rngMount As Range, rngShip As Range, rngModel As Range
    On Error Resume Next
    Set rngMat = thisWorkbook.names("AG_Material").RefersToRange
    Set rngMount = thisWorkbook.names("AG_MountType").RefersToRange
    Set rngShip = thisWorkbook.names("AG_Completion").RefersToRange
    Set rngModel = thisWorkbook.names("AG_Model").RefersToRange
    On Error GoTo 0
    
    If rngShip Is Nothing Then Exit Function
    
    Dim matVals As Variant, mountVals As Variant, shipVals As Variant, modelVals As Variant
    matVals = rngMat.Value
    mountVals = rngMount.Value
    shipVals = rngShip.Value
    modelVals = rngModel.Value
    
    If Not IsArray(matVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = matVals: matVals = arr
    If Not IsArray(mountVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = mountVals: mountVals = arr
    If Not IsArray(shipVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = shipVals: shipVals = arr
    If Not IsArray(modelVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = modelVals: modelVals = arr
    
    Dim col As New Collection
    Dim i As Long
    Dim cleanMat As String, cleanMount As String, cleanInstall As String, cleanModel As String
    
    cleanMat = CleanStringForTable(CStr(material))
    cleanMount = CleanStringForTable(CStr(mountType))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanModel = CleanStringForTable(CStr(model))
    
    For i = 1 To UBound(matVals, 1)
        Dim mVal As Variant, mtVal As Variant, sVal As Variant, modVal As Variant
        mVal = matVals(i, 1): mtVal = mountVals(i, 1): sVal = shipVals(i, 1): modVal = modelVals(i, 1)
        
        If Not IsError(sVal) And Not IsEmpty(sVal) Then
            Dim cleanSVal As String
            cleanSVal = CleanStringForTable(CStr(sVal))
            
            Dim matchMat As Boolean, matchMount As Boolean, matchModel As Boolean, matchInstall As Boolean
            matchMat = (cleanMat = "" Or CleanStringForTable(CStr(mVal)) = cleanMat)
            matchMount = (cleanMount = "" Or CleanStringForTable(CStr(mtVal)) = cleanMount)
            matchModel = (cleanModel = "" Or CleanStringForTable(CStr(modVal)) = cleanModel)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(CleanStringForTable(CStr(mtVal)), cleanInstall)
            
            If matchMat And matchMount And matchModel And matchInstall Then
                AddUniqueToCollection col, sVal
            End If
        End If
    Next i
    
    GetDeliveryListCross = CollectionToString(col)
    Exit Function
CleanExit:
    GetDeliveryListCross = ""
End Function

' Returns list of models filtered by other criteria
Public Function GetModelListCross(ByVal material As Variant, ByVal mountType As Variant, _
                                  ByVal installType As Variant, ByVal delivery As Variant) As String
    On Error GoTo CleanExit
    
    Dim rngMat As Range, rngMount As Range, rngShip As Range, rngModel As Range
    On Error Resume Next
    Set rngMat = thisWorkbook.names("AG_Material").RefersToRange
    Set rngMount = thisWorkbook.names("AG_MountType").RefersToRange
    Set rngShip = thisWorkbook.names("AG_Completion").RefersToRange
    Set rngModel = thisWorkbook.names("AG_Model").RefersToRange
    On Error GoTo 0
    
    If rngModel Is Nothing Then Exit Function
    
    Dim matVals As Variant, mountVals As Variant, shipVals As Variant, modelVals As Variant
    matVals = rngMat.Value
    mountVals = rngMount.Value
    shipVals = rngShip.Value
    modelVals = rngModel.Value
    
    If Not IsArray(matVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = matVals: matVals = arr
    If Not IsArray(mountVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = mountVals: mountVals = arr
    If Not IsArray(shipVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = shipVals: shipVals = arr
    If Not IsArray(modelVals) Then ReDim arr(1 To 1, 1 To 1): arr(1, 1) = modelVals: modelVals = arr
    
    Dim col As New Collection
    Dim i As Long
    Dim cleanMat As String, cleanMount As String, cleanInstall As String, cleanShip As String
    
    cleanMat = CleanStringForTable(CStr(material))
    cleanMount = CleanStringForTable(CStr(mountType))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanShip = CleanStringForTable(CStr(delivery))
    
    For i = 1 To UBound(matVals, 1)
        Dim mVal As Variant, mtVal As Variant, sVal As Variant, modVal As Variant
        mVal = matVals(i, 1): mtVal = mountVals(i, 1): sVal = shipVals(i, 1): modVal = modelVals(i, 1)
        
        If Not IsError(modVal) And Not IsEmpty(modVal) Then
            Dim cleanModVal As String
            cleanModVal = CleanStringForTable(CStr(modVal))
            
            Dim matchMat As Boolean, matchMount As Boolean, matchShip As Boolean, matchInstall As Boolean
            matchMat = (cleanMat = "" Or CleanStringForTable(CStr(mVal)) = cleanMat)
            matchMount = (cleanMount = "" Or CleanStringForTable(CStr(mtVal)) = cleanMount)
            matchShip = (cleanShip = "" Or CleanStringForTable(CStr(sVal)) = cleanShip)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(CleanStringForTable(CStr(mtVal)), cleanInstall)
            
            If matchMat And matchMount And matchShip And matchInstall Then
                AddUniqueToCollection col, modVal
            End If
        End If
    Next i
    
    GetModelListCross = CollectionToString(col)
    Exit Function
CleanExit:
    GetModelListCross = ""
End Function

' Helper to add unique item to collection
Private Sub AddUniqueToCollection(ByVal col As Collection, ByVal Value As Variant)
    On Error Resume Next
    Dim item As Variant
    For Each item In col
        If item = Value Then Exit Sub
    Next item
    col.Add Value
    On Error GoTo 0
End Sub

' Helper to convert collection to comma-separated string
Private Function CollectionToString(ByVal col As Collection) As String
    Dim result As String
    Dim item As Variant
    For Each item In col
        If result <> "" Then result = result & ","
        result = result & item
    Next item
    CollectionToString = result
End Function

Public Function NormalizeFilterText(ByVal Value As String) As String
    Dim normalized As String

    normalized = CStr(Value)
    normalized = Replace(normalized, ChrW(160), " ")
    normalized = Replace(normalized, ChrW(9), " ")
    normalized = Replace(normalized, ChrW(10), " ")
    normalized = Replace(normalized, ChrW(13), " ")
    normalized = Replace(normalized, ChrW(1105), ChrW(1077))
    normalized = Replace(normalized, ChrW(1025), ChrW(1045))
    normalized = CleanStringForTable(normalized)

    Do While InStr(normalized, "  ") > 0
        normalized = Replace(normalized, "  ", " ")
    Loop

    NormalizeFilterText = Trim$(normalized)
End Function
