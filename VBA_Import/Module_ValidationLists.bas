Attribute VB_Name = "Module_ValidationLists"

Option Explicit

' ================================================================
' Module_ValidationLists
' Unique lists and cross-filtering. ListAG / TableMounting are cached
' so RefreshValidationForColumn does not re-read Excel on every column.
' ================================================================

Private cacheReady As Boolean
Private cacheN As Long
Private cacheMatRaw() As Variant
Private cacheMountRaw() As Variant
Private cacheShipRaw() As Variant
Private cacheModelRaw() As Variant
Private cacheMatClean() As String
Private cacheMountClean() As String
Private cacheShipClean() As String
Private cacheModelClean() As String

Private mountN As Long
Private mountRaw() As Variant
Private installRaw() As Variant
Private mountClean() As String
Private installClean() As String
Private mountPairs As Collection

Private strAllMaterials As String
Private strAllMountTypes As String
Private strAllInstallTypes As String
Private strAllDeliveries As String
Private strAllModels As String

Public Sub InvalidateAGListCache()
    cacheReady = False
    cacheN = 0
    mountN = 0
    Set mountPairs = Nothing
    strAllMaterials = ""
    strAllMountTypes = ""
    strAllInstallTypes = ""
    strAllDeliveries = ""
    strAllModels = ""
End Sub

' Returns list of unique materials
Public Function GetAllMaterials() As String
    If Not EnsureAGListCache() Then Exit Function
    GetAllMaterials = strAllMaterials
End Function

' Returns list of unique mounting types
Public Function GetAllMountTypes() As String
    If Not EnsureAGListCache() Then Exit Function
    GetAllMountTypes = strAllMountTypes
End Function

' Returns list of unique installation types
Public Function GetAllInstallTypes() As String
    If Not EnsureAGListCache() Then Exit Function
    GetAllInstallTypes = strAllInstallTypes
End Function

' Returns list of unique deliveries
Public Function GetAllDeliveries() As String
    If Not EnsureAGListCache() Then Exit Function
    GetAllDeliveries = strAllDeliveries
End Function

' Returns list of unique AG models
Public Function GetAllAGModels() As String
    If Not EnsureAGListCache() Then Exit Function
    GetAllAGModels = strAllModels
End Function

' Returns list of installation types that depend only on mounting type
Public Function GetInstallTypeListFromMountType(ByVal mountType As String) As String
    On Error GoTo CleanExit
    If Not EnsureAGListCache() Then Exit Function

    Dim col As New Collection
    Dim i As Long
    Dim cleanMount As String
    cleanMount = CleanStringForTable(CStr(mountType))

    For i = 1 To mountN
        If Not IsError(mountRaw(i)) And Not IsError(installRaw(i)) And Not IsEmpty(installRaw(i)) Then
            If cleanMount = "" Or InStr(1, mountClean(i), cleanMount, vbTextCompare) > 0 Then
                If installClean(i) <> "" Then AddUniqueToCollection col, installClean(i)
            End If
        End If
    Next i

    GetInstallTypeListFromMountType = CollectionToString(col)
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

    If Not EnsureAGListCache() Then
        CheckInstallationExists = False
        Exit Function
    End If
    If mountPairs Is Nothing Then
        CheckInstallationExists = False
        Exit Function
    End If

    Dim key As String
    key = PairKey(NormalizeFilterText(mountType), NormalizeFilterText(installType))
    If key = "" Then
        CheckInstallationExists = False
        Exit Function
    End If

    On Error Resume Next
    Dim dummy As Variant
    dummy = mountPairs.item(key)
    CheckInstallationExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
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
    If Not EnsureAGListCache() Then Exit Function

    Dim cleanMountType As String
    Dim i As Long
    cleanMountType = CleanStringForTable(mountType)

    For i = 1 To mountN
        If mountClean(i) = cleanMountType Then
            If Not IsError(installRaw(i)) And Not IsEmpty(installRaw(i)) And CStr(installRaw(i)) <> "" Then
                GetInstallationTypeFromMountType = CStr(installRaw(i))
                Exit Function
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
    If Not EnsureAGListCache() Then Exit Function

    Dim col As New Collection
    Dim i As Long
    Dim cleanMount As String, cleanInstall As String, cleanShip As String, cleanModel As String
    Dim matchMount As Boolean, matchShip As Boolean, matchModel As Boolean, matchInstall As Boolean

    cleanMount = CleanStringForTable(CStr(mountType))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanShip = CleanStringForTable(CStr(delivery))
    cleanModel = CleanStringForTable(CStr(model))

    For i = 1 To cacheN
        If Not IsError(cacheMatRaw(i)) And Not IsEmpty(cacheMatRaw(i)) Then
            matchMount = (cleanMount = "" Or cacheMountClean(i) = cleanMount)
            matchShip = (cleanShip = "" Or cacheShipClean(i) = cleanShip)
            matchModel = (cleanModel = "" Or cacheModelClean(i) = cleanModel)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall)
            If matchMount And matchShip And matchModel And matchInstall Then
                AddUniqueToCollection col, cacheMatRaw(i)
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
    If Not EnsureAGListCache() Then Exit Function

    Dim col As New Collection
    Dim i As Long
    Dim cleanMat As String, cleanInstall As String, cleanShip As String, cleanModel As String
    Dim matchMat As Boolean, matchShip As Boolean, matchModel As Boolean, matchInstall As Boolean

    cleanMat = CleanStringForTable(CStr(material))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanShip = CleanStringForTable(CStr(delivery))
    cleanModel = CleanStringForTable(CStr(model))

    For i = 1 To cacheN
        If Not IsError(cacheMountRaw(i)) And Not IsEmpty(cacheMountRaw(i)) Then
            matchMat = (cleanMat = "" Or cacheMatClean(i) = cleanMat)
            matchShip = (cleanShip = "" Or cacheShipClean(i) = cleanShip)
            matchModel = (cleanModel = "" Or cacheModelClean(i) = cleanModel)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall)
            If matchMat And matchShip And matchModel And matchInstall Then
                AddUniqueToCollection col, cacheMountRaw(i)
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
    If Not EnsureAGListCache() Then Exit Function

    Dim col As New Collection
    Dim i As Long
    Dim cleanMat As String, cleanMount As String, cleanInstall As String, cleanModel As String
    Dim matchMat As Boolean, matchMount As Boolean, matchModel As Boolean, matchInstall As Boolean

    cleanMat = CleanStringForTable(CStr(material))
    cleanMount = CleanStringForTable(CStr(mountType))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanModel = CleanStringForTable(CStr(model))

    For i = 1 To cacheN
        If Not IsError(cacheShipRaw(i)) And Not IsEmpty(cacheShipRaw(i)) Then
            matchMat = (cleanMat = "" Or cacheMatClean(i) = cleanMat)
            matchMount = (cleanMount = "" Or cacheMountClean(i) = cleanMount)
            matchModel = (cleanModel = "" Or cacheModelClean(i) = cleanModel)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall)
            If matchMat And matchMount And matchModel And matchInstall Then
                AddUniqueToCollection col, cacheShipRaw(i)
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
    If Not EnsureAGListCache() Then Exit Function

    Dim col As New Collection
    Dim i As Long
    Dim cleanMat As String, cleanMount As String, cleanInstall As String, cleanShip As String
    Dim matchMat As Boolean, matchMount As Boolean, matchShip As Boolean, matchInstall As Boolean

    cleanMat = CleanStringForTable(CStr(material))
    cleanMount = CleanStringForTable(CStr(mountType))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanShip = CleanStringForTable(CStr(delivery))

    For i = 1 To cacheN
        If Not IsError(cacheModelRaw(i)) And Not IsEmpty(cacheModelRaw(i)) Then
            matchMat = (cleanMat = "" Or cacheMatClean(i) = cleanMat)
            matchMount = (cleanMount = "" Or cacheMountClean(i) = cleanMount)
            matchShip = (cleanShip = "" Or cacheShipClean(i) = cleanShip)
            matchInstall = True
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall)
            If matchMat And matchMount And matchShip And matchInstall Then
                AddUniqueToCollection col, cacheModelRaw(i)
            End If
        End If
    Next i

    GetModelListCross = CollectionToString(col)
    Exit Function
CleanExit:
    GetModelListCross = ""
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

' ================================================================
' cache
' ================================================================

Private Function EnsureAGListCache() As Boolean
    If cacheReady Then
        EnsureAGListCache = True
        Exit Function
    End If

    On Error GoTo Fail

    Dim matVals As Variant, mountVals As Variant, shipVals As Variant, modelVals As Variant
    matVals = LoadName2D("AG_Material")
    mountVals = LoadName2D("AG_MountType")
    shipVals = LoadName2D("AG_Completion")
    modelVals = LoadName2D("AG_Model")
    If IsEmpty(matVals) Or IsEmpty(mountVals) Or IsEmpty(shipVals) Or IsEmpty(modelVals) Then GoTo Fail

    cacheN = UBound(matVals, 1)
    ReDim cacheMatRaw(1 To cacheN)
    ReDim cacheMountRaw(1 To cacheN)
    ReDim cacheShipRaw(1 To cacheN)
    ReDim cacheModelRaw(1 To cacheN)
    ReDim cacheMatClean(1 To cacheN)
    ReDim cacheMountClean(1 To cacheN)
    ReDim cacheShipClean(1 To cacheN)
    ReDim cacheModelClean(1 To cacheN)

    Dim colMat As New Collection
    Dim colMount As New Collection
    Dim colShip As New Collection
    Dim colModel As New Collection
    Dim i As Long
    For i = 1 To cacheN
        cacheMatRaw(i) = matVals(i, 1)
        cacheMountRaw(i) = mountVals(i, 1)
        cacheShipRaw(i) = shipVals(i, 1)
        cacheModelRaw(i) = modelVals(i, 1)
        cacheMatClean(i) = CleanCell(cacheMatRaw(i))
        cacheMountClean(i) = CleanCell(cacheMountRaw(i))
        cacheShipClean(i) = CleanCell(cacheShipRaw(i))
        cacheModelClean(i) = CleanCell(cacheModelRaw(i))
        AddUniqueToCollection colMat, cacheMatRaw(i)
        AddUniqueToCollection colMount, cacheMountRaw(i)
        AddUniqueToCollection colShip, cacheShipRaw(i)
        AddUniqueToCollection colModel, cacheModelRaw(i)
    Next i
    strAllMaterials = CollectionToString(colMat)
    strAllMountTypes = CollectionToString(colMount)
    strAllDeliveries = CollectionToString(colShip)
    strAllModels = CollectionToString(colModel)

    Dim wsMounting As Worksheet
    Dim tblMounting As ListObject
    On Error Resume Next
    Set wsMounting = thisWorkbook.Worksheets("ListAG")
    If wsMounting Is Nothing Then Set wsMounting = thisWorkbook.Worksheets(SHEET_LIST_AG)
    If Not wsMounting Is Nothing Then Set tblMounting = wsMounting.ListObjects("TableMounting")
    Err.Clear
    On Error GoTo Fail

    Call LoadMountingCache(tblMounting)

    cacheReady = True
    EnsureAGListCache = True
    Exit Function

Fail:
    cacheReady = False
    EnsureAGListCache = False
End Function

' TableMounting is optional: a missing/broken table must not fail ListAG lists.
Private Sub LoadMountingCache(ByVal tblMounting As ListObject)
    Set mountPairs = New Collection
    mountN = 0
    strAllInstallTypes = ""

    On Error GoTo FailMount
    If tblMounting Is Nothing Then Exit Sub
    If tblMounting.DataBodyRange Is Nothing Then Exit Sub

    Dim mVals As Variant, iVals As Variant
    Dim colInst As Collection
    Dim key As String
    Dim i As Long

    mVals = As2D(tblMounting.ListColumns(4).DataBodyRange.Value)
    iVals = As2D(tblMounting.ListColumns(3).DataBodyRange.Value)
    mountN = UBound(mVals, 1)
    ReDim mountRaw(1 To mountN)
    ReDim installRaw(1 To mountN)
    ReDim mountClean(1 To mountN)
    ReDim installClean(1 To mountN)

    Set colInst = New Collection
    For i = 1 To mountN
        mountRaw(i) = mVals(i, 1)
        installRaw(i) = iVals(i, 1)
        mountClean(i) = CleanCell(mountRaw(i))
        installClean(i) = CleanCell(installRaw(i))
        AddUniqueToCollection colInst, installRaw(i)
        If Not IsError(mountRaw(i)) And Not IsError(installRaw(i)) Then
            key = PairKey(NormalizeFilterText(CStr(mountRaw(i))), NormalizeFilterText(CStr(installRaw(i))))
            If key <> "" Then
                On Error Resume Next
                mountPairs.Add True, key
                Err.Clear
                On Error GoTo FailMount
            End If
        End If
    Next i
    strAllInstallTypes = CollectionToString(colInst)
    Exit Sub

FailMount:
    Err.Clear
    mountN = 0
    Set mountPairs = New Collection
    strAllInstallTypes = ""
End Sub

Private Function LoadName2D(ByVal nm As String) As Variant
    Dim rng As Range
    On Error Resume Next
    Set rng = thisWorkbook.names(nm).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    LoadName2D = As2D(rng.Value)
End Function

Private Function As2D(ByVal vals As Variant) As Variant
    If IsArray(vals) Then
        As2D = vals
    Else
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = vals
        As2D = arr
    End If
End Function

Private Function CleanCell(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        CleanCell = ""
    Else
        CleanCell = CleanStringForTable(CStr(v))
    End If
End Function

Private Function PairKey(ByVal mountNorm As String, ByVal installNorm As String) As String
    If mountNorm = "" Or installNorm = "" Then
        PairKey = ""
    Else
        PairKey = "k" & mountNorm & Chr$(1) & installNorm
    End If
End Function

Private Sub AddUniqueToCollection(ByVal col As Collection, ByVal Value As Variant)
    If IsError(Value) Then Exit Sub
    If IsEmpty(Value) Then Exit Sub
    If CStr(Value) = "" Then Exit Sub
    On Error Resume Next
    col.Add Value, CStr(Value)
    On Error GoTo 0
End Sub

Private Function CollectionToString(ByVal col As Collection) As String
    Dim result As String
    Dim item As Variant
    For Each item In col
        If result <> "" Then result = result & ","
        result = result & item
    Next item
    CollectionToString = result
End Function
