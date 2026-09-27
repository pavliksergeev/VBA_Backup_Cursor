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
Private shipRaw() As Variant
Private mountClean() As String
Private installClean() As String
Private shipClean() As String
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

' Returns list of installation types for TableAG mount type AND комплектация.
' комплектный -> TableMounting rows marked комплектный (with activator).
' некомплектный -> rows marked некомплектный or with empty комплектация (no activator).
Public Function GetInstallTypeListFromMountType(ByVal mountType As String, Optional ByVal delivery As Variant) As String
    On Error GoTo CleanExit
    If Not EnsureAGListCache() Then Exit Function

    Dim col As New Collection
    Dim i As Long
    Dim cleanMount As String
    Dim cleanShip As String
    cleanMount = CleanStringForTable(CStr(mountType))
    cleanShip = OptionalClean(delivery)

    For i = 1 To mountN
        If Not IsError(installRaw(i)) And Not IsEmpty(installRaw(i)) Then
            If MountShipMatch(i, cleanMount, cleanShip) Then
                If installClean(i) <> "" Then AddUniqueToCollection col, installClean(i)
            End If
        End If
    Next i

    GetInstallTypeListFromMountType = CollectionToString(col)
    Exit Function
CleanExit:
    GetInstallTypeListFromMountType = ""
End Function

' Checks if TableMounting has this mount + install, optionally for комплектация.
' комплектный only matches комплектный rows; некомплектный matches empty/некомплектный.
Public Function CheckInstallationExists(ByVal mountType As String, ByVal installType As String, Optional ByVal delivery As Variant) As Boolean
    On Error GoTo CleanExit

    If mountType = "" Or installType = "" Then
        CheckInstallationExists = True
        Exit Function
    End If

    If Not EnsureAGListCache() Then
        CheckInstallationExists = False
        Exit Function
    End If

    Dim cleanMount As String
    Dim cleanInstall As String
    Dim cleanShip As String
    Dim i As Long
    cleanMount = CleanStringForTable(CStr(mountType))
    cleanInstall = CleanStringForTable(CStr(installType))
    cleanShip = OptionalClean(delivery)
    If cleanMount = "" Or cleanInstall = "" Then
        CheckInstallationExists = False
        Exit Function
    End If

    For i = 1 To mountN
        If installClean(i) = cleanInstall Then
            If MountShipMatch(i, cleanMount, cleanShip) Then
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

' Returns first installation type for mounting type and optional комплектация
Public Function GetInstallationTypeFromMountType(ByVal mountType As String, Optional ByVal delivery As Variant) As String
    On Error GoTo CleanExit

    If mountType = "" Or IsEmpty(mountType) Then
        GetInstallationTypeFromMountType = ""
        Exit Function
    End If
    If Not EnsureAGListCache() Then Exit Function

    Dim cleanMountType As String
    Dim cleanShip As String
    Dim i As Long
    cleanMountType = CleanStringForTable(mountType)
    cleanShip = OptionalClean(delivery)

    For i = 1 To mountN
        If MountShipMatch(i, cleanMountType, cleanShip) Then
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
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall, cacheShipClean(i))
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
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall, cacheShipClean(i))
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
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall, cacheShipClean(i))
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
            If cleanInstall <> "" Then matchInstall = CheckInstallationExists(cacheMountClean(i), cleanInstall, cacheShipClean(i))
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

    Dim mVals As Variant, iVals As Variant, sVals As Variant
    Dim colInst As Collection
    Dim key As String
    Dim i As Long
    Dim colMountIdx As Long, colInstIdx As Long, colShipIdx As Long

    colInstIdx = MountColIndex(tblMounting, Ru("0441 043F 043E 0441 043E 0431 0020 043C 043E 043D 0442 0430 0436 0430"), 3)
    colMountIdx = MountColIndex(tblMounting, Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430"), 4)
    colShipIdx = MountColIndex(tblMounting, Ru("043A 043E 043C 043F 043B 0435 043A 0442 0430 0446 0438 044F"), 5)
    If colMountIdx = 0 Or colInstIdx = 0 Then Exit Sub

    mVals = As2D(tblMounting.ListColumns(colMountIdx).DataBodyRange.Value)
    iVals = As2D(tblMounting.ListColumns(colInstIdx).DataBodyRange.Value)
    If colShipIdx > 0 Then
        sVals = As2D(tblMounting.ListColumns(colShipIdx).DataBodyRange.Value)
    Else
        sVals = Empty
    End If
    mountN = UBound(mVals, 1)
    ReDim mountRaw(1 To mountN)
    ReDim installRaw(1 To mountN)
    ReDim shipRaw(1 To mountN)
    ReDim mountClean(1 To mountN)
    ReDim installClean(1 To mountN)
    ReDim shipClean(1 To mountN)

    Set colInst = New Collection
    For i = 1 To mountN
        mountRaw(i) = mVals(i, 1)
        installRaw(i) = iVals(i, 1)
        If IsArray(sVals) Then
            shipRaw(i) = sVals(i, 1)
        Else
            shipRaw(i) = ""
        End If
        mountClean(i) = CleanCell(mountRaw(i))
        installClean(i) = CleanCell(installRaw(i))
        shipClean(i) = CleanCell(shipRaw(i))
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

Private Function OptionalClean(ByVal delivery As Variant) As String
    If IsMissing(delivery) Then
        OptionalClean = ""
    ElseIf IsError(delivery) Or IsEmpty(delivery) Then
        OptionalClean = ""
    Else
        OptionalClean = CleanStringForTable(CStr(delivery))
    End If
End Function

' TableAG тип монтажа = TableMounting тип монтажа.
' комплектный: only TableMounting комплектация = комплектный (always with activator).
' некомплектный: TableMounting комплектация empty or некомплектный (always without activator).
Private Function MountShipMatch(ByVal i As Long, ByVal cleanMount As String, ByVal cleanShip As String) As Boolean
    If cleanMount <> "" And mountClean(i) <> cleanMount Then Exit Function
    If cleanShip <> "" Then
        If KitIsComplete(cleanShip) Then
            If Not KitIsComplete(shipClean(i)) Then Exit Function
        ElseIf KitIsIncomplete(cleanShip) Then
            If KitIsComplete(shipClean(i)) Then Exit Function
        Else
            If shipClean(i) <> "" And shipClean(i) <> cleanShip Then Exit Function
        End If
    End If
    MountShipMatch = True
End Function

Private Function KitIsIncomplete(ByVal cleanShip As String) As Boolean
    Dim needle As String
    If Len(cleanShip) = 0 Then Exit Function
    needle = CleanStringForTable(Ru("043D 0435 043A 043E 043C 043F 043B 0435 043A 0442 043D 044B 0439"))
    KitIsIncomplete = (InStr(1, cleanShip, needle, vbTextCompare) > 0)
End Function

Private Function KitIsComplete(ByVal cleanShip As String) As Boolean
    Dim needle As String
    If Len(cleanShip) = 0 Then Exit Function
    If KitIsIncomplete(cleanShip) Then Exit Function
    needle = CleanStringForTable(Ru("043A 043E 043C 043F 043B 0435 043A 0442 043D 044B 0439"))
    KitIsComplete = (InStr(1, cleanShip, needle, vbTextCompare) > 0)
End Function

Private Function MountColIndex(ByVal tbl As ListObject, ByVal headerRu As String, ByVal fallback As Long) As Long
    Dim c As Long
    Dim hdr As String
    If tbl Is Nothing Then Exit Function
    hdr = CleanStringForTable(headerRu)
    For c = 1 To tbl.ListColumns.Count
        If StrComp(CleanCell(tbl.ListColumns(c).Name), hdr, vbTextCompare) = 0 Then
            MountColIndex = c
            Exit Function
        End If
    Next c
    If fallback >= 1 And fallback <= tbl.ListColumns.Count Then MountColIndex = fallback
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
