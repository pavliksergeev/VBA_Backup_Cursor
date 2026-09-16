Attribute VB_Name = "Module_NamedRanges"

' ================================================================
' create named ranges for TableAG columns
' (full version with aliases)
' ================================================================
Sub CreateAGColumnNames()
    On Error GoTo CleanExit
    
    Dim tbl As ListObject
    Set tbl = thisWorkbook.Worksheets(SHEET_LIST_AG).ListObjects(TABLE_AG)
    If tbl Is Nothing Then
        ' таблица не найдена!
        ' table not found!
        MsgBox Ru("0442 0430 0431 043B 0438 0446 0430 0020") & TABLE_AG & Ru("0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 0021"), vbExclamation
        Exit Sub
    End If
    
    ' mapping: header -> range name
    Dim headers As Variant
    Dim names As Variant
    Dim i As Long
    
    ' материал анода
    ' anode material
    ' тип монтажа
    ' mounting type
    ' комплектация
    ' delivery set
    ' диаметр электрода заземлителя (de, m)
    ' grounding electrode diameter (de, m)
    ' длина рабочей части электрода заземлителя (le, m)
    ' working length of grounding electrode (le, m)
    ' диаметр электрода заземлителя в коксовой засыпке (dz, m)
    ' grounding electrode diameter in coke breeze backfill (dz, m)
    ' длина электрода заземлителя в коксовой засыпке (la, m)
    ' grounding electrode length in coke breeze backfill (la, m)
    ' масса материала электрода заземлителя без коксовой засыпки (ge, kg)
    ' electrode material mass without coke breeze (ge, kg)
    ' скорость растворения материала заземлителя (qz, kg/(a*year))
    ' anode dissolution rate (qz, kg/(a*year))
    ' номинальная токовая нагрузка, a, не более (ia.n, a)
    ' rated current load, a, not more than (ia.n, a)
    ' удельное электрическое сопротивление материала рабочего элемента заземлителя (ra, om*m)
    ' electrical resistivity anode working-element material (ra, om*m)
    ' удельное электрическое сопротивление коксовой засыпки заземлителя (rkz, om*m)
    ' electrical resistivity coke breeze backfill of anode (rkz, om*m)
    ' удельная номинальная токовая нагрузка на погонный метр изделия (iaz.n, a/m)
    ' specific rated current load on linear meter of the product (iaz.n, a/m)
    ' удельная масса рабочего элемента протяженного заземлителя (ge.pr, kg/m)
    ' specific mass of working element extended anode (ge.pr, kg/m)
    headers = Array( _
        Ru("043C 0430 0442 0435 0440 0438 0430 043B 0020 0430 043D 043E 0434 0430"), _
        Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430"), _
        Ru("043A 043E 043C 043F 043B 0435 043A 0442 0430 0446 0438 044F"), _
        "model", _
        Ru("0434 0438 0430 043C 0435 0442 0440 0020 044D 043B 0435 043A 0442 0440 043E 0434 0430 0020 0437 0430 0437 0435 043C 043B 0438 0442 0435 043B 044F 0020 0028 0064 0065 002C 0020") & Ru("006D 0029"), _
        Ru("0434 043B 0438 043D 0430 0020 0440 0430 0431 043E 0447 0435 0439 0020 0447 0430 0441 0442 0438 0020 044D 043B 0435 043A 0442 0440 043E 0434 0430 0020 0437 0430 0437 0435 043C") & Ru("043B 0438 0442 0435 043B 044F 0020 0028 006C 0065 002C 0020 006D 0029"), _
        Ru("0434 0438 0430 043C 0435 0442 0440 0020 044D 043B 0435 043A 0442 0440 043E 0434 0430 0020 0437 0430 0437 0435 043C 043B 0438 0442 0435 043B 044F 0020 0432 0020 043A 043E 043A") & Ru("0441 043E 0432 043E 0439 0020 0437 0430 0441 044B 043F 043A 0435 0020 0028 0064 007A 002C 0020 006D 0029"), _
        Ru("0434 043B 0438 043D 0430 0020 044D 043B 0435 043A 0442 0440 043E 0434 0430 0020 0437 0430 0437 0435 043C 043B 0438 0442 0435 043B 044F 0020 0432 0020 043A 043E 043A 0441 043E") & Ru("0432 043E 0439 0020 0437 0430 0441 044B 043F 043A 0435 0020 0028 006C 0061 002C 0020 006D 0029"), _
        Ru("043C 0430 0441 0441 0430 0020 043C 0430 0442 0435 0440 0438 0430 043B 0430 0020 044D 043B 0435 043A 0442 0440 043E 0434 0430 0020 0437 0430 0437 0435 043C 043B 0438 0442 0435") & Ru("043B 044F 0020 0431 0435 0437 0020 043A 043E 043A 0441 043E 0432 043E 0439 0020 0437 0430 0441 044B 043F 043A 0438 0020 0028 0047 0065 002C 0020 006B 0067 0029"), _
        Ru("0441 043A 043E 0440 043E 0441 0442 044C 0020 0440 0430 0441 0442 0432 043E 0440 0435 043D 0438 044F 0020 043C 0430 0442 0435 0440 0438 0430 043B 0430 0020 0437 0430 0437 0435") & Ru("043C 043B 0438 0442 0435 043B 044F 0020 0028 0071 007A 002C 0020 006B 0067 002F 0028 0041 002A 0079 0065 0061 0072 0029 0029"), _
        Ru("043D 043E 043C 0438 043D 0430 043B 044C 043D 0430 044F 0020 0442 043E 043A 043E 0432 0430 044F 0020 043D 0430 0433 0440 0443 0437 043A 0430 002C 0020 0041 002C 0020 043D 0435") & Ru("0020 0431 043E 043B 0435 0435 0020 0028 0049 0061 002E 006E 002C 0020 0041 0029"), _
        Ru("0443 0434 0435 043B 044C 043D 043E 0435 0020 044D 043B 0435 043A 0442 0440 0438 0447 0435 0441 043A 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438") & Ru("0435 0020 043C 0430 0442 0435 0440 0438 0430 043B 0430 0020 0440 0430 0431 043E 0447 0435 0433 043E 0020 044D 043B 0435 043C 0435 043D 0442 0430 0020 0437 0430 0437 0435 043C") & Ru("043B 0438 0442 0435 043B 044F 0020 0028 0072 0061 002C 0020 004F 006D 002A 006D 0029"), _
        Ru("0443 0434 0435 043B 044C 043D 043E 0435 0020 044D 043B 0435 043A 0442 0440 0438 0447 0435 0441 043A 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438") & Ru("0435 0020 043A 043E 043A 0441 043E 0432 043E 0439 0020 0437 0430 0441 044B 043F 043A 0438 0020 0437 0430 0437 0435 043C 043B 0438 0442 0435 043B 044F 0020 0028 0072 006B 007A") & Ru("002C 0020 004F 006D 002A 006D 0029"), _
        Ru("0443 0434 0435 043B 044C 043D 0430 044F 0020 043D 043E 043C 0438 043D 0430 043B 044C 043D 0430 044F 0020 0442 043E 043A 043E 0432 0430 044F 0020 043D 0430 0433 0440 0443 0437") & Ru("043A 0430 0020 043D 0430 0020 043F 043E 0433 043E 043D 043D 044B 0439 0020 043C 0435 0442 0440 0020 0438 0437 0434 0435 043B 0438 044F 0020 0028 0049 0061 007A 002E 006E 002C") & Ru("0020 0041 002F 006D 0029"), _
        Ru("0443 0434 0435 043B 044C 043D 0430 044F 0020 043C 0430 0441 0441 0430 0020 0440 0430 0431 043E 0447 0435 0433 043E 0020 044D 043B 0435 043C 0435 043D 0442 0430 0020 043F 0440") & Ru("043E 0442 044F 0436 0435 043D 043D 043E 0433 043E 0020 0437 0430 0437 0435 043C 043B 0438 0442 0435 043B 044F 0020 0028 0047 0065 002E 0070 0072 002C 0020 006B 0067 002F 006D") & Ru("0029") _
    )
    
    names = Array( _
        "AG_Material", _
        "AG_MountType", _
        "AG_Completion", _
        "AG_Model", _
        "AG_Diameter", _
        "AG_Length", _
        "AG_CokeDiam", _
        "AG_CokeLength", _
        "AG_Mass", _
        "AG_DissolutionRate", _
        "AG_RatedCurrent", _
        "AG_ResistivityMaterial", _
        "AG_CokeResistivity", _
        "AG_SpecificRatedCurrent", _
        "AG_SpecificMass" _
    )
    
    ' check the counts
    If UBound(headers) <> UBound(names) Then
        ' количество заголовков не совпадает с количеством имён!
        ' header count does not match the count of names!
        MsgBox Ru("043A 043E 043B 0438 0447 0435 0441 0442 0432 043E 0020 0437 0430 0433 043E 043B 043E 0432 043A 043E 0432 0020 043D 0435 0020 0441 043E 0432 043F 0430 0434 0430 0435 0442 0020") & Ru("0441 0020 043A 043E 043B 0438 0447 0435 0441 0442 0432 043E 043C 0020 0438 043C 0451 043D 0021"), vbCritical
        Exit Sub
    End If
    
    Dim col As ListColumn
    Dim headerVal As String
    Dim nameToCreate As String
    Dim createdCount As Long
    Dim found As Boolean
    
    ' walk all table columns
    For Each col In tbl.ListColumns
        headerVal = Trim(col.Range.Cells(1, 1).Value)
        found = False
        
        For i = LBound(headers) To UBound(headers)
            If StrComp(headerVal, headers(i), vbTextCompare) = 0 Then
                nameToCreate = names(i)
                found = True
                Exit For
            End If
        Next i
        
        If found Then
            On Error Resume Next
            thisWorkbook.names(nameToCreate).Delete
            On Error GoTo 0
            
            thisWorkbook.names.Add name:=nameToCreate, RefersTo:=col.DataBodyRange
            createdCount = createdCount + 1
            ' создано имя: 
            ' name created: 
            Debug.Print Ru("0441 043E 0437 0434 0430 043D 043E 0020 0438 043C 044F 003A 0020") & nameToCreate
        End If
    Next col
    
    ' ================================================================
    ' create aliases for compatibility with wsAnod
    ' ================================================================
    Dim aliasNames As Variant
    aliasNames = Array( _
        "typeMaterial", "AG_Material", _
        "typeMountingAG", "AG_MountType", _
        "typeDeliveryAG", "AG_Completion", _
        "typeAG", "AG_Model", _
        "diameterAG", "AG_Diameter", _
        "lengthElectrodeAG", "AG_Length", _
        "cokeBreezeDiameterAG", "AG_CokeDiam", _
        "cokeBreezelengthElectrodeAG", "AG_CokeLength", _
        "massOneElectrodeAG", "AG_Mass", _
        "dissolutionRateAG", "AG_DissolutionRate", _
        "ratedCurrent", "AG_RatedCurrent", _
        "resistivityMaterialAG", "AG_ResistivityMaterial", _
        "cokeBreezeResistivityAG", "AG_CokeResistivity", _
        "specificRatedCurrent", "AG_SpecificRatedCurrent", _
        "specificMaccOneMeterAG", "AG_SpecificMass" _
    )
    
    For i = LBound(aliasNames) To UBound(aliasNames) Step 2
        Dim aliasName As String
        Dim targetName As String
        aliasName = aliasNames(i)
        targetName = aliasNames(i + 1)
        
        On Error Resume Next
        Dim testRng As Range
        Set testRng = thisWorkbook.names(aliasName).RefersToRange
        If Err.Number <> 0 Then
            Err.Clear
            Dim targetRange As Range
            Set targetRange = thisWorkbook.names(targetName).RefersToRange
            If Not targetRange Is Nothing Then
                thisWorkbook.names.Add name:=aliasName, RefersTo:=targetRange
                ' создан синоним: 
                ' alias created: 
                Debug.Print Ru("0441 043E 0437 0434 0430 043D 0020 0441 0438 043D 043E 043D 0438 043C 003A 0020") & aliasName & " -> " & targetName
                createdCount = createdCount + 1
            End If
        End If
        On Error GoTo 0
    Next i
    
    ' создано именованных диапазонов: 
    ' named ranges created: 
    MsgBox Ru("0441 043E 0437 0434 0430 043D 043E 0020 0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0445 0020 0434 0438 0430 043F 0430 0437 043E 043D 043E 0432 003A 0020") & createdCount, vbInformation
    Exit Sub

CleanExit:
    If Err.Number <> 0 Then
        ' ошибка: 
        ' error: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020") & Err.Description, vbCritical
    End If
End Sub
