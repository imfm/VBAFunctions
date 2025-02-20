
Function MergeArrays(arr1, arr2, Optional remdup = 1, Optional alti = 0, Optional altj = 0)
    If NumberOfArrayDimensions(arr1) > 1 Then
        arr1 = Flatten2DArray(arr1)
    End If
    If NumberOfArrayDimensions(arr2) > 1 Then
        arr2 = Flatten2DArray(arr2)
    End If
    Dim arr3() As Variant
    ReDim arr3(UBound(arr1) + UBound(arr2) + 1)
    If alti = 0 Then
        alti = LBound(arr1)
    End If
    If altj = 0 Then
        altj = LBound(arr2)
    End If
    
    Dim i As Integer
    For i = alti To UBound(arr1)
        arr3(i) = arr1(i)
    Next i
    Dim j As Integer
    For j = altj To UBound(arr2)
        arr3(i) = arr2(j)
        i = i + 1
    Next j
    
    If remdup Then
        arr3 = RemoveArrayDuplicates(arr3)
    End If
    
    MergeArrays = arr3

End Function

Function RemoveArrayDuplicates(arr, Optional RemoveNull = 1, Optional SplitRemoveDelimiter, Optional SplitRemoveDirection)
    initialsize = UBound(arr) - LBound(arr)
    NewSize = initialsize
    Dim newarr() As Variant
    ReDim newarr(0)
    
    undupcount = 0
    For i = LBound(arr) To UBound(arr)
        inarray = False
        For j = LBound(newarr) To UBound(newarr)
            arrval = arr(i)
            newarrval = newarr(j)
            If Not IsMissing(SplitRemoveDelimiter) Then
                arrval = SplitString(arrval, SplitRemoveDelimiter, SplitRemoveDirection)
                newarrval = SplitString(newarrval, SplitRemoveDelimiter, SplitRemoveDirection)
            End If
            If arrval = newarrval Then
                inarray = True
            End If
        Next j
        If Not inarray Then
            If RemoveNull = 0 Or (Not IsEmpty(arr(i)) And Not arr(i) = "") Then
                ReDim Preserve newarr(undupcount)
                newarr(undupcount) = arr(i)
                undupcount = undupcount + 1
            End If
        End If
    Next i
    'ReDim Preserve newarr(undupcount - 1)
    RemoveArrayDuplicates = newarr
End Function

Function CopyUnique(cell, searchcol)
    copysheet = cell.Parent.Name
    searchsheet = searchcol.Parent.Name
    myRow = cell.row
    mycol = cell.Column
    index = myRow - 1
    previouscell = Worksheets(copysheet).Cells(index, mycol)
    searchindex = myRow - 1
    previousfound = False
    Do
        searchcell = Worksheets(searchsheet).Cells(searchindex, searchcol.Column)
        If previouscell = searchcell Then
            previousfound = True
        ElseIf previousfound And Not Previous = searchcell Then
            CopyUnique = searchcell
            Exit Do
        End If
        searchindex = searchindex + 1
    Loop While Not IsEmpty(searchcell)
    If IsEmpty(searchcell) Then
        CopyUnique = ""
    End If
    'MsgBox (myrow & " " & mycol & " " & searchcell)
End Function


Function ArrayCompare(cell As String, cell2 As String, Optional ReturnType = 0)
    Temp = Split(cell, ";")
    temp2 = Split(cell2, ";")
    
    Dim CopyMatchTempA(15) As Variant
    Dim CopyMatchTemp
    Dim inarray
    CopyMatchTempA(0) = ""
    i = 0
    MatchFound = False
    
    For Each cs In Temp
        If (IsEmpty(cs)) Or cs = "" Then
            Exit For
        End If
        For Each cs2 In temp2
            If CStr(c2) = CStr(cs) Or cs2 = cs And Not IsNullNull(cs2) And Not cs2 = "" Then
                'MsgBox (" match " & cs & " " & match & " cs2 " & cs2)
                MatchFound = True
                CopyMatchTemp = cs2
                inarray = False
                For Each av In CopyMatchTempA
                    If (CopyMatchTemp = av) Then
                        inarray = True
                        Exit For
                    End If
                Next
                If Not (inarray) And Not IsNullNull(CopyMatchTemp) And Not CopyMatchTemp = "" Then
                    CopyMatchTempA(i) = CopyMatchTemp
                    i = i + 1
                End If
            End If
        Next
    Next

    If IsNullNull(CopyMatchTempA(0)) Or CopyMatchTempA(0) = "" Then
        ArrayCompare = ""
    Else
        CopyMatchTemp = ""
        index = 0
        Do
            tempstr = CStr(CopyMatchTempA(index))
            CopyMatchTemp = CopyMatchTemp + tempstr + ";"
            index = index + 1
        Loop Until index = i
        If ReturnType = 1 Then
            ArrayCompare = Split(CopyMatchTemp, ";")
        ElseIf ReturnType = 2 Then
            ArrayCompare = ArraySize(CopyMatchTemp)
        Else
            ArrayCompare = CopyMatchTemp
        End If
    End If
End Function

Function ArraySize(a)
    If Not IsArray(a) Then
        a = Split(a, ";")
    End If
    Temp = 0
    For Each i In a
        Temp = Temp + 1
    Next
    ArraySize = Temp - 1
End Function

Function ArrayGet(cell As String, index As Integer, Optional dateconv = 0)
    Temp = Split(cell, ";")
    If index = -1 Then
        aSize = ArraySize(Temp)
        index = aSize - 1
    End If
    If dateconv Then
        ArrayGet = CDate(Temp(index))
    Else
        ArrayGet = Temp(index)
    End If
End Function


Function ArrayGetMatch(cell As String, searchr)
    Temp = Split(cell, ";")
    searchsheet = searchr.Parent.Name
    Dim CopyMatchTempA(15) As Variant
    Dim CopyMatchTemp
    Dim inarray
    CopyMatchTempA(0) = ""
    i = 0
    MatchFound = False
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            Exit For
        End If
        For Each match In Temp
            'MsgBox (" match " & match & " cs " & cs)
            If CStr(match) = CStr(cs) Or match = cs Then
                'MsgBox (" match " & cs & " " & match)
                MatchFound = True
                CopyMatchTemp = match
                inarray = False
                For Each av In CopyMatchTempA
                    If (CopyMatchTemp = av) Then
                        inarray = True
                        Exit For
                    End If
                Next
                If Not (inarray) And Not IsNullNull(CopyMatchTemp) Then
                    CopyMatchTempA(i) = CopyMatchTemp
                    i = i + 1
                End If
            End If
        Next
    Next

    If IsNullNull(CopyMatchTempA(0)) Or CopyMatchTempA(0) = "" Then
        ArrayGetMatch = ""
    Else
        CopyMatchTemp = ""
        index = 0
        Do
            tempstr = CStr(CopyMatchTempA(index))
            CopyMatchTemp = CopyMatchTemp + tempstr + ";"
            index = index + 1
        Loop Until index = i
        ArrayGetMatch = CopyMatchTemp
    End If

End Function

Function ArrayGetMatchValues(cell As String, searchr, valuesr, Optional sizer)
    On Error Resume Next
    Temp = Split(cell, ";")
    searchsheet = searchr.Parent.Name
    Dim CopyMatchTempA(5) As Variant
    Dim CopyMatchTemp
    Dim inarray
    CopyMatchTempA(0) = ""
    MatchFound = False
    i = 0
    j = 1
    If sizer Then
        rangesize = sizer.End(xlDown).row
    Else
        rangesize = 0
    End If
    
    For Each cs In searchr
        If rangesize = 0 Then
            If IsEmpty(cs) Then
                Exit For
            End If
        ElseIf j > rangesize Then
            Exit For
        End If
        j = j + 1
        For Each match In Temp
            If CStr(match) = CStr(cs) Or match = cs Then
                MatchFound = True
                CopyMatchTemp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
                inarray = False
                For Each av In CopyMatchTempA
                    If (CopyMatchTemp = av) Then
                        inarray = True
                        Exit For
                    End If
                Next
                If Not (inarray) And Not IsNullNull(CopyMatchTemp) Then
                    CopyMatchTempA(i) = CopyMatchTemp
                    i = i + 1
                End If
            End If
        Next
    Next

    If IsNullNull(CopyMatchTempA(0)) Or CopyMatchTempA(0) = "" Then
        ArrayGetMatchValues = ""
    Else
        CopyMatchTemp = ""
        index = 0
        Do
            tempstr = CStr(CopyMatchTempA(index))
            CopyMatchTemp = CopyMatchTemp + tempstr + ";"
            index = index + 1
        Loop Until index = i
        ArrayGetMatchValues = CopyMatchTemp
    End If

End Function


Function ArrayCountMatch(cell As String, searchr, Optional token = ";")
    Temp = Split(cell, token)
    searchsheet = searchr.Parent.Name
    MatchCount = 0
    
    For Each csas In searchr
        If (IsEmpty(csas)) Then
            Exit For
        End If
        csa = Split(csas, token)
        For Each cs In csa
            For Each match In Temp
                If Trim(CStr(match)) = Trim(CStr(cs)) Or match = cs Then
                    MatchCount = MatchCount + 1
                End If
            Next
        Next
    Next
    
    ArrayCountMatch = MatchCount

End Function

Function ArraySumMatch(cell As String, searchr, valuesr)
    Temp = Split(cell, ";")
    searchsheet = searchr.Parent.Name
    matchsum = 0
    
    For Each csas In searchr
        If (IsEmpty(csas)) Then
            Exit For
        End If
        csa = Split(csas, ";")
        For Each cs In csa
            For Each match In Temp
                If CStr(match) = CStr(cs) Or match = cs Then
                    matchvalue = Worksheets(searchsheet).Cells(csas.row, valuesr.Column)
                    matchsum = matchsum + matchvalue
                End If
            Next
        Next
    Next
    
    ArraySumMatch = matchsum

End Function

Function ArrayCountMatchMulti(cell As String, searchr, match2, matchr)
    
    On Error Resume Next
    
    Temp = Split(cell, ";")
    searchsheet = searchr.Parent.Name
    MatchCount = 0
    match2found = False
    
    For Each match In matchr
        If (IsEmpty(match)) Then
            Exit For
        ElseIf CStr(match) = CStr(match2) Or match = match2 Then
            match2found = True
            csas = Worksheets(searchsheet).Cells(match.row, searchr.Column)
            csa = Split(csas, ";")
            For Each tv In Temp
                For Each cs In csa
                    If (CStr(tv) = CStr(cs) Or tv = cs) Then
                        MatchCount = MatchCount + 1
                    End If
                Next
            Next
        End If
    Next
    
    ArrayCountMatchMulti = MatchCount

End Function


Function ArraySequenceCalc(cell As String, interval As Integer, Optional endval)
    Temp = Split(cell, ";")
    tempsize = 0
    
    For Each tempval In Temp
        tempsize = tempsize + 1
    Next
    tempsize = tempsize - 1
    
    index = 0
    temptotal = 0
    Do
        val1 = Temp(index)
        val2 = Temp(index + 1)
        If (val2 = "") And Not IsMissing(endval) And index = (tempsize - 1) Then
            val2 = endval
            interval = 999
        ElseIf index = (tempsize - 1) Then
            Exit Do
        End If
        If IsDate(val1) And IsDate(val2) Then
            diff = DateDiff("d", val1, val2)
        Else
            diff = val2 - val1
        End If
        'MsgBox (val2 & " " & val1 & " " & diff)
        If diff <= interval Then
            temptotal = temptotal + diff
        Else
            temptotal = 0
        End If
        index = index + 1
    Loop Until index = tempsize
    
    ArraySequenceCalc = temptotal
    
End Function

Function CopyMatchArray(match, searchr, valuesr, Optional speedoption = 0, Optional zero = 0)
    
    searchsheet = searchr.Parent.Name
    Dim CopyMatchTempA(15) As Variant
    Dim CopyMatchTemp
    Dim inarray
    
    CopyMatchTempA(0) = ""
    i = 0
    MatchFound = False
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            'MsgBox (" exiting loop ")
            Exit For
        End If
        If match = cs Then
            MatchFound = True
            CopyMatchTemp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            inarray = False
            'MsgBox (" Checking if Value: " & CopyMatchTemp & "  - is in array ")
            For Each av In CopyMatchTempA
                'MsgBox (" not crashed ")
                If (CopyMatchTemp = av) Then
                    'MsgBox (" Value: " & CopyMatchTemp & "  - is in array ")
                    inarray = True
                    Exit For
                End If
            Next
            If Not (inarray) And Not IsNullNull(CopyMatchTemp) Then
                CopyMatchTempA(i) = CopyMatchTemp
                i = i + 1
                'MsgBox (" not crashed inside loop. added value: " & i & CopyMatchTempA(i))
            End If
        ElseIf (speedoption = 1 And MatchFound) Then
            Exit For
        ElseIf (speedoption = 2 And (cs > match) And (vartype(cs) = vartype(match))) Then
            'MsgBox (" cs: " & cs & VarType(cs) & "  match: " & match)
            Exit For
        End If
    Next

    If IsNullNull(CopyMatchTempA(0)) Or CopyMatchTempA(0) = "" Then
        CopyMatchArray = ""
    Else
        CopyMatchTemp = ""
        index = 0
        Do
            tempstr = CStr(CopyMatchTempA(index))
            CopyMatchTemp = CopyMatchTemp + tempstr + ";"
            index = index + 1
        Loop Until index = i
        CopyMatchArray = CopyMatchTemp
    End If

End Function

Function CopyMatchArrayMulti(match, searchr, match2, searchr2, valuesr, Optional speedoption = 0, Optional zero = 0)
    
    searchsheet = searchr.Parent.Name
    Dim CopyMatchTempA(15) As Variant
    Dim CopyMatchTemp
    Dim inarray
    
    CopyMatchTempA(0) = ""
    i = 0
    MatchFound = False
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            Exit For
        End If
        If match = cs Then
            cs2 = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
            If match2 = cs2 Then
                MatchFound = True
                CopyMatchTemp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
                inarray = False
                'MsgBox (" Checking if Value: " & CopyMatchTemp & "  - is in array ")
                For Each av In CopyMatchTempA
                    'MsgBox (" not crashed ")
                    If (CopyMatchTemp = av) Then
                        'MsgBox (" Value: " & CopyMatchTemp & "  - is in array ")
                        inarray = True
                        Exit For
                    End If
                Next
                If Not (inarray) And Not IsNullNull(CopyMatchTemp) And Not CopyMatchTemp = "" Then
                    CopyMatchTempA(i) = CopyMatchTemp
                    i = i + 1
                    'MsgBox (" not crashed inside loop. added value: " & i & CopyMatchTempA(i))
                End If
            End If
        ElseIf (speedoption = 1 And MatchFound) Then
            Exit For
        ElseIf (speedoption = 2 And (cs > match) And (vartype(cs) = vartype(match))) Then
            'MsgBox (" cs: " & cs & VarType(cs) & "  match: " & match)
            Exit For
        End If
    Next

    If IsNullNull(CopyMatchTempA(0)) Or CopyMatchTempA(0) = "" Then
        CopyMatchArrayMulti = ""
    Else
        CopyMatchTemp = ""
        index = 0
        Do
            tempstr = CStr(CopyMatchTempA(index))
            CopyMatchTemp = CopyMatchTemp + tempstr + ";"
            index = index + 1
        Loop Until index = i
        CopyMatchArrayMulti = CopyMatchTemp
    End If

End Function

Function CopyMatchCompare(match, searchr, valuesr, Optional operation = 0, Optional converttypes = 0)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    CopyMatchTemp1 = ""
    CopyMatchTemp2 = ""
    MatchFound = False
    
    'Temp = searchr.Sort(searchr, xlAscending)
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            CopyMatchCompare = CopyMatchTemp1
            If (CopyMatchCompare = "") Then
                If (zero) Then
                    CopyMatchCompare = 0
                Else
                    CopyMatchCompare = ""
                End If
            End If
            Exit For
        End If
        If (converttypes) Then
            matchv = ConvertType(match, converttypes)
            csv = ConvertType(cs, converttypes)
        Else
            matchv = match
            csv = cs
        End If
        If matchv = csv Then
            MatchFound = True
            If (IsNull(CopyMatchTemp1)) Then
                'MsgBox (" setting temp1 ")
                CopyMatchTemp1 = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            End If
            CopyMatchTemp2 = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            If ((CopyMatchTemp1 = "" Or CopyMatchTemp2 = "") And (operation = 3 Or operation = 4)) Then
                'MsgBox (" NULL END DATE FOUND ")
                CopyMatchCompare = ""
                Exit Function
            End If
            If (operation = 1) Then
                If (CopyMatchTemp2 < CopyMatchTemp1 And Not CopyMatchTemp2 = "") Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 2) Then
                'MsgBox (" comparing " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2)
                If (CopyMatchTemp2 > CopyMatchTemp1 And Not CopyMatchTemp2 = "") Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 3) Then
                If ((CopyMatchTemp2 < CopyMatchTemp1) Or (CopyMatchTemp1 = "") Or IsNull(CopyMatchTemp1)) Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 4) Then
                'MsgBox (" comparing temp1: " & CopyMatchTemp1 & " temp2: " & CopyMatchTemp2)
                If ((CopyMatchTemp2 > CopyMatchTemp1) Or (CopyMatchTemp1 = "") Or IsNull(CopyMatchTemp1)) Then
                    'MsgBox (" Switching values - temp1 " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2 & "CMTSET: " & CMTSET)
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            End If
            'MsgBox (" temp1 " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2 & "CMTSET: " & CMTSET)
        ElseIf MatchFound Then
            Exit For
        End If
    Next

    If IsNull(CopyMatchTemp1) Or CopyMatchTemp1 = "" Then
        CopyMatchCompare = ""
    Else
        CopyMatchCompare = CopyMatchTemp1
    End If

End Function

Function CopyMatchCompareMulti(match, searchr, match2, searchr2, valuesr, Optional operation = 0, Optional speedoption = 0, Optional zero = 0)
    
    searchsheet = searchr.Parent.Name
    CopyMatchTemp1 = ""
    CopyMatchTemp2 = ""
    MatchFound = False
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            CopyMatchCompareMulti = CopyMatchTemp1
            If (CopyMatchCompareMulti = "") Then
                If (zero) Then
                    CopyMatchCompareMulti = 0
                Else
                    CopyMatchCompareMulti = ""
                End If
            End If
            Exit For
        End If
        If match = cs Then
            MatchFound = True
            If (IsNull(CopyMatchTemp1)) Then
                CopyMatchTemp1 = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            End If
            CopyMatchTemp2 = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            If ((CopyMatchTemp1 = "" Or CopyMatchTemp2 = "") And (operation = 3 Or operation = 4)) Then
                'MsgBox (" NULL END DATE FOUND ")
                CopyMatchCompareMulti = ""
                Exit Function
            End If
            cs2 = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
            If (operation = 1) Then
                If (CopyMatchTemp2 < CopyMatchTemp1 And Not CopyMatchTemp2 = "") And cs2 = match2 Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 2) Then
                'MsgBox (" comparing " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2)
                If (CopyMatchTemp2 > CopyMatchTemp1 And Not CopyMatchTemp2 = "") And cs2 = match2 Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 3) Then
                If ((CopyMatchTemp2 < CopyMatchTemp1) Or (CopyMatchTemp1 = "") Or IsNull(CopyMatchTemp1)) And cs2 = match2 Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 4) Then
                'MsgBox (" comparing temp1: " & CopyMatchTemp1 & " temp2: " & CopyMatchTemp2)
                If ((CopyMatchTemp2 > CopyMatchTemp1) Or (CopyMatchTemp1 = "") Or IsNull(CopyMatchTemp1)) And cs2 = match2 Then
                    'MsgBox (" Switching values - temp1 " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2 & "CMTSET: " & CMTSET)
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            End If
            'MsgBox (" temp1 " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2 & "CMTSET: " & CMTSET)
        ElseIf MatchFound And speedoption Then
            Exit For
        End If
    Next

    If IsNullNull(CopyMatchTemp1) Or CopyMatchTemp1 = "" Then
        CopyMatchCompareMulti = ""
    Else
        CopyMatchCompareMulti = CopyMatchTemp1
    End If

End Function

Function CopyMatchCompareMatchMulti(match, searchr, match2, searchr2, valuesr, comparer, Optional operation = 0, Optional speedoption = 0, Optional zero = 0)
    
    searchsheet = searchr.Parent.Name
    CopyMatchTemp1 = ""
    CopyMatchTemp2 = ""
    MatchFound = False
    compare1 = ""
    compare2 = ""
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            CopyMatchCompareMatchMulti = CopyMatchTemp1
            If (CopyMatchCompareMatchMulti = "") Then
                If (zero) Then
                    CopyMatchCompareMatchMulti = 0
                Else
                    CopyMatchCompareMatchMulti = ""
                End If
            End If
            Exit For
        End If
        If match = cs Then
            MatchFound = True
            If (IsNull(CopyMatchTemp1)) Then
                CopyMatchTemp1 = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            End If
            CopyMatchTemp2 = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            If ((CopyMatchTemp1 = "" Or CopyMatchTemp2 = "") And (operation = 3 Or operation = 4)) Then
                'MsgBox (" NULL END DATE FOUND ")
                CopyMatchCompareMatchMulti = ""
                Exit Function
            End If
            cs2 = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
            compare1 = compare2
            compare2 = Worksheets(searchsheet).Cells(cs.row, comparer.Column)
            If (operation = 1) Then
                If (compare2 < compare1 And Not compare2 = "") And cs2 = match2 Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 2) Then
                'MsgBox (" comparing " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2)
                If (compare2 > compare1 And Not compare2 = "") And cs2 = match2 Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 3) Then
                If ((compare2 < compare1) Or (compare1 = "") Or IsNull(compare1)) And cs2 = match2 Then
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            ElseIf (operation = 4) Then
                'MsgBox (" comparing temp1: " & CopyMatchTemp1 & " temp2: " & CopyMatchTemp2)
                If ((compare2 > compare1) Or (compare2 = "") Or IsNull(compare2)) And cs2 = match2 Then
                    'MsgBox (" Switching values - temp1 " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2 & "CMTSET: " & CMTSET)
                    CopyMatchTemp1 = CopyMatchTemp2
                End If
            End If
            'MsgBox (" temp1 " & CopyMatchTemp1 & " temp2 " & CopyMatchTemp2 & "CMTSET: " & CMTSET)
        ElseIf MatchFound And speedoption Then
            Exit For
        End If
    Next

    If IsNullNull(CopyMatchTemp1) Or CopyMatchTemp1 = "" Then
        CopyMatchCompareMatchMulti = ""
    Else
        CopyMatchCompareMatchMulti = CopyMatchTemp1
    End If

End Function

Function CopyMatchDateInRange(match, searchr, dater, startdt As Date, enddt As Date) As Date

    searchsheet = searchr.Parent.Name
    Dim date1
        
    For Each cs In searchr
        If IsEmpty(cs) Then
            CopyMatchDateInRange = Null
            Exit For
        End If
        If match = cs Then
            date1 = Worksheets(searchsheet).Cells(cs.row, dater.Column)
            If (date1 >= startdt And date1 <= enddt) Then
                'MsgBox ("Date IS in range " & startdt & " - " & date1 & " - " & enddt)
                CopyMatchDateInRange = date1
                Exit For
            Else
                'MsgBox ("Date not in range " & startdt & " - " & date1 & " - " & enddt)
            End If
        End If
    Next
    
End Function

Function CopyMatchLast(match, searchr, valuesr, Optional zero = 0)
    
    searchsheet = searchr.Parent.Name
    CopyMatchTemp = ""
    
    For Each cs In searchr
        If IsEmpty(cs) Then
            CopyMatchLast = CopyMatchTemp
            Exit For
        End If
        'MsgBox (" Match-Search " & match & " - " & cs)
        If match = cs Then
            CopyMatchTemp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            'MsgBox ("Found Match " & match & cs & " Row " & cs.Row & " Col " & cs.Column & " Val " & CopyMatchLast)
        End If
    Next
    CopyMatchLast = CopyMatchTemp
    If (CopyMatchLast = "") Then
        If (zero) Then
            CopyMatchLast = 0
        Else
            CopyMatchLast = ""
        End If
    End If
End Function

Public Function CopyMatchPercent(match, searchr, valuesr, percent, Optional zero = 0)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    If vartype(match) = vbString Then
        match = StripString(UCase(match), 2)
    End If
    
    For Each cs In searchr
        If Filtered(cs) = 0 Then
            If IsNullAlt(cs, 1) = 2 Then
                Exit For
            End If
            csr = cs.row
            If vartype(match) = vbString Then
                cs = StripString(UCase(cs.Text), 2)
            End If
            If StringPercentMatch(match, cs) >= percent Then
                CopyMatchPercent = Worksheets(searchsheet).Cells(csr, valuesr.Column)
                If (Not CopyMatchPercent = "") Then
                    Exit For
                End If
            End If
        End If
    Next
    
    If (IsNullAlt(CopyMatchPercent) And zero = 1) Then
        CopyMatchPercent = 0
    ElseIf (IsNullAlt(CopyMatchPercent) And zero = 0) Then
        CopyMatchPercent = ""
    End If
    
End Function

Public Function CopyMatchIFsP(match, searchr, percent, or1, match2, searchr2, percent2, or2, match3, searchr3, percent3, valuesr, Optional xWord, Optional zero = 0)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    matchfound1 = False
    matchfound2 = False
    matchfound3 = False
    
    For Each cs In searchr
        css = cs.Value
        If Filtered(cs) = 0 Then
            If IsNullAlt(cs, 1) = 2 Then
                Exit For
            End If
            If vartype(match) = vbString Then
                match = StripString(UCase(match), 2)
                css = StripString(UCase(cs.Text), 2)
            End If
            If StringPercentMatch(match, css, 0, xWord) >= percent Then
                matchfound1 = True
            End If
            If matchfound1 Or or1 Then
                cs2s = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
                If vartype(match2) = vbString Then
                    match2 = StripString(UCase(match2), 2)
                    cs2s = StripString(UCase(cs2s), 2)
                End If
                If StringPercentMatch(match2, cs2s, 0, xWord) >= percent2 Then
                    matchfound2 = True
                End If
                If (matchfound1 And matchfound2) Or (matchfound2 And or1) Then
                    cs3s = Worksheets(searchsheet).Cells(cs.row, searchr3.Column)
                    If vartype(match3) = vbString Then
                        match3 = StripString(UCase(match3), 2)
                        cs3s = StripString(UCase(cs3s), 2)
                    End If
                    If StringPercentMatch(match3, cs3s, 0, xWord) >= percent3 Then
                        matchfound3 = True
                    End If
                    If (matchfound2 And matchfound3) Or (matchfound3 And or2) Then
                        CopyMatchIFsP = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
                        If (Not CopyMatchIFsP = "") Then
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If IsNullAlt(CopyMatchIFsP) Then
        If zero Then
            CopyMatchIFsP = 0
        Else
            CopyMatchIFsP = ""
        End If
    End If
    
End Function

Public Function CopyMatch(match, searchr As Range, valuesr As Range, Optional speedoption = 0, Optional zero = 0)
    'On Error Resume Next
    'On Error GoTo errhandler
    If Application.Ready Then
        Application.ScreenUpdating = False
    End If
    
    searchsheet = searchr.Parent.Name
    MatchFound = False
    
    If vartype(match) = vbString Then
        matcht = StripString(UCase(match.Text), 1)
    Else
        matcht = match.Text
    End If
    matchv = ConvertType(match, 5)
    match = match
    
    For Each cs In searchr
        If Filtered(cs) = 0 Then
            If IsNullAlt(cs, 1) > 1 Then
                Exit For
            End If
            csr = cs.row
            If vartype(match) = vbString Then
                cst = StripString(UCase(cs.Text), 1)
            Else
                cst = cs.Text
            End If
            csv = ConvertType(cs, 5)
            cs = cs
            If (match = cs) Or (matcht = cst) Or ((matchv = csv) And Not IsEmpty(matchv) And Not IsEmpty(csv)) Then
                CopyMatch = Worksheets(searchsheet).Cells(csr, valuesr.Column)
                If (Not CopyMatch = "") Then
                    Exit For
                End If
            End If
        End If
    Next
    If (CopyMatch = "" And zero = 1) Then
        CopyMatch = 0
    ElseIf (CopyMatch = "" And zero = 0) Then
        CopyMatch = ""
    End If
    
    If Application.Ready Then
        Application.ScreenUpdating = True
    End If

'errhandler:
'    MsgBox ("Error " & Err.Number & " - " & Err.Description)

End Function

Function RawValue(var)
    On Error Resume Next
    vtext = var.Text
    vint = CDbl(var)
    vvar = var
    vval = var.Value
    If Not IsEmpty(vint) Then
        RawValue = vint
    ElseIf Not IsEmpty(vval) Then
        RawValue = vval
    ElseIf Not IsEmpty(vvar) Then
        RawValue = vvar
    Else
        RawValue = vtext
    End If
End Function

Public Function EstimateCount(searchr)
    searchsheet = searchr.Parent.Name
    searchcolumn = searchr.Column
    tempcount = searchr.count
    cell = Worksheets(searchsheet).Cells(tempcount, searchcolumn)
    If IsEmpty(cell) Then
        Do While IsEmpty(cell)
            tempcount = Round(tempcount / 2)
            cell = Worksheets(searchsheet).Cells(tempcount, searchcolumn)
        Loop
        EstimateCount = tempcount * 2
    Else
        EstimateCount = tempcount
    End If
End Function

Function CopyMatchMultiFast(match, match2, searchr, searchr2, valuesr, Optional speedoption = 0, Optional nulls = 0, Optional zero = "")
    
    searchsheet = searchr.Parent.Name
    MatchFound = False
    If vartype(match) = 8 Then
        match = LCase(match)
    End If
    If vartype(match2) = 8 Then
        match2 = LCase(match2)
    End If
    
    For Each cs In searchr
        If IsEmpty(cs) Then
            Exit For
        End If
        If vartype(cs) = 8 Then
            cs = LCase(cs)
        End If
        If match = cs Then
            MatchFound = True
            cs2 = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
            If vartype(cs2) = 8 Then
                cs2 = LCase(cs2)
            End If
            If (cs2 = match2) Then
                Temp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
                If Not (nulls) And Not (IsNull(Temp)) Then
                    'MsgBox (" TRUE ")
                    'MsgBox (temp)
                    CopyMatchMultiFast = Temp
                    Exit For
                End If
            End If
        ElseIf (MatchFound And speedoption = 1) Then
            Exit For
        ElseIf (speedoption = 2 And (cs > match) And (vartype(cs) = vartype(match))) Then
            Exit For
        End If
    Next
    'MsgBox (temp)
    'MsgBox (CopyMatchMultiFast)
    'MsgBox (" Not Crashed ")
    If IsNull(CopyMatchMultiFast) And (zero = 0) Then
        'MsgBox (" Null Zero ")
        CopyMatchMultiFast = 0
    ElseIf IsNull(CopyMatchMultiFast) Then
        'MsgBox (" Null ")
        CopyMatchMultiFast = ""
    End If
    
End Function

Function CopyMatchMultiRangeFast(match, match2, match3, searchr, searchr2, searchr3, valuesr, Optional speedoption = 0, Optional nulls = 0, Optional zero = "")
    
    searchsheet = searchr.Parent.Name
    MatchFound = False
    
    For Each cs In searchr
        If IsEmpty(cs) Then
            Exit For
        End If
        If match = cs Or Trim(UCase(match)) = Trim(UCase(cs)) Then
            MatchFound = True
            cs2 = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
            If IsNull(match3) Then
                cs3 = ""
            Else
                cs3 = Worksheets(searchsheet).Cells(cs.row, searchr3.Column)
            End If
            If (cs2 = match2 And cs3 = match3) Or (Trim(UCase(match2)) = Trim(UCase(cs2)) And Trim(UCase(match3)) = Trim(UCase(cs3))) Then
                Temp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
                If Not (nulls) And Not (IsNull(Temp)) Then
                    'MsgBox (" TRUE ")
                    'MsgBox (temp)
                    CopyMatchMultiRangeFast = Temp
                    Exit For
                End If
            End If
        ElseIf (MatchFound And speedoption = 1) Then
            Exit For
        ElseIf (speedoption = 2 And (cs > match) And (vartype(cs) = vartype(match))) Then
            Exit For
        End If
    Next
    'MsgBox (temp)
    'MsgBox (CopyMatchMultiRangeFast)
    'MsgBox (" Not Crashed ")
    If IsNullAlt(CopyMatchMultiRangeFast) And (zero = 0) Then
        'MsgBox (" Null Zero ")
        CopyMatchMultiRangeFast = 0
    ElseIf IsNullAlt(CopyMatchMultiRangeFast) Then
        'MsgBox (" Null ")
        CopyMatchMultiRangeFast = ""
    End If
    
End Function

Function DateInRange(date1, date2, date3) As Integer
        If date1 = "" Or date2 = "" Or date3 = "" Then
            DateInRange = 0
            Exit Function
        End If
        date1 = CDate(date1)
        date2 = CDate(date2)
        date3 = CDate(date3)
        If (date1 >= date2 And date1 <= date3) Then
            DateInRange = 1
        Else
            DateInRange = 0
        End If
End Function

Function DateRangeInRange(dateI1, dateI2, dateO1, dateO2) As Integer
        If dateI1 = "" Or dateI2 = "" Or dateO1 = "" Or dateO2 = "" Then
            DateInRange = 0
            Exit Function
        End If
        dateI1 = CDate(dateI1)
        dateI2 = CDate(dateI2)
        dateO1 = CDate(dateO1)
        dateO2 = CDate(dateO2)
        If (dateI1 >= dateO1 And dateI1 <= dateO2 And dateI2 >= dateO1 And dateI2 <= dateO2) Then
            DateRangeInRange = 1
        Else
            DateRangeInRange = 0
        End If
End Function

Function DateInRangeCount(dater, date2, date3) As Integer
    
    searchsheet = dater.Parent.Name
    date2 = CDate(date2)
    date3 = CDate(date3)
    DateInRangeCount = 0
    On Error Resume Next
    
    For Each cs In dater
        If IsEmpty(cs) Then
            Exit For
        End If
        If Not IsNull(cs) And Not cs = "" Then
            date1 = CDate(cs)
            If (date1 >= date2 And date1 <= date3) Then
                DateInRangeCount = DateInRangeCount + 1
            End If
        End If
    Next
    
End Function

Function DateInRangeSum(dater, date2, date3, sumr) As Integer
    
    searchsheet = dater.Parent.Name
    date2 = CDate(date2)
    date3 = CDate(date3)
    DateInRangeSum = 0
    On Error Resume Next
    
    For Each cs In dater
        If IsEmpty(cs) Then
            Exit For
        End If
        If Not IsNull(cs) And Not cs = "" Then
            date1 = CDate(cs)
            If (date1 >= date2 And date1 <= date3) Then
                sumv = Worksheets(searchsheet).Cells(cs.row, sumr.Column)
                DateInRangeSum = DateInRangeSum + sumv
            End If
        End If
    Next
    
End Function

Function DateInRangeMatchCount(dater As Range, date2, date3, match, searchr) As Integer
    
    On Error Resume Next
    
    searchsheet = dater.Parent.Name
    date2 = CDate(date2)
    date3 = CDate(date3)
    DateInRangeMatchCount = 0
    nullcount = 0
    
    For Each cs In dater
        If IsNull(cs) Then
            nullcount = nullcount + 1
            If nullcount > 10 Then
                Exit For
            End If
        End If
        If Not IsNull(cs) And Not cs = "" Then
            nullcount = 0
            date1 = CDate(cs)
            If (date1 >= date2 And date1 <= date3) Then
                cs2 = Worksheets(searchsheet).Cells(cs.row, searchr.Column)
                If cs2 = match Then
                    DateInRangeMatchCount = DateInRangeMatchCount + 1
                End If
            End If
        End If
    Next
    
End Function

Function DateInRangeMatch(match, searchr, dater, startdt As Date, enddt As Date) As Integer

    searchsheet = searchr.Parent.Name
    
    'Dim date1 As Date
    Dim date1
        
            For Each cs In searchr
                If IsEmpty(cs) Then
                    DateInRangeMatch = 0
                    Exit For
                End If
                If match = cs Then
                    'date1 = CDate(Worksheets(searchsheet).Cells(cs.Row, dater.Column))
                    date1 = Worksheets(searchsheet).Cells(cs.row, dater.Column)
                    'MsgBox ("vars: .row " & date1r.Row & " .column " & date1r.Column & " date1r(x) " & date1r(cs.Row))
                    'MsgBox ("Found Match " & cm & cs & " Row " & cs.Row & " Col " & cs.Column)
                    If (date1 >= startdt And date1 <= enddt) Then
                        'MsgBox ("Date IS in range " & startdt & " - " & date1 & " - " & enddt)
                        DateInRangeMatch = 1
                        Exit For
                    Else
                        'MsgBox ("Date not in range " & startdt & " - " & date1 & " - " & enddt)
                    End If
                End If
            Next
        
End Function

Function DateInRangeMatchGet(match, searchr, dater, startdt As Date, enddt As Date, Optional compare = 0, Optional sorted = 0, Optional zero = 1)

    searchsheet = searchr.Parent.Name
    
    Dim date1
    Dim prevdate
    MatchFound = 0
        
            For Each cs In searchr
                If IsEmpty(cs) Then
                    If zero And IsEmpty(DateInRangeMatchGet) Then
                        DateInRangeMatchGet = 0
                    End If
                    Exit For
                ElseIf match = cs Then
                    MatchFound = 1
                    date1 = Worksheets(searchsheet).Cells(cs.row, dater.Column)
                    If (date1 >= startdt And date1 <= enddt) Then
                        If Not IsEmpty(DateInRangeMatchGet) Then
                            prevdate = DateInRangeMatchGet
                        End If
                        DateInRangeMatchGet = date1
                        If compare = 0 Then
                            Exit For
                        ElseIf compare < 0 Then
                            If prevdate < DateInRangeMatchGet And Not IsEmpty(prevdate) Then
                                DateInRangeMatchGet = prevdate
                            End If
                        ElseIf compare > 0 Then
                            If prevdate > DateInRangeMatchGet And Not IsEmpty(prevdate) Then
                                DateInRangeMatchGet = prevdate
                            End If
                        End If
                    End If
                ElseIf sorted And MatchFound Then
                    Exit For
                End If
            Next
        
End Function

Function DateOverlap(start1, end1, start2, end2) As Integer
    
    start1 = CDate(start1)
    end1 = CDate(end1)
    start2 = CDate(start2)
    end2 = CDate(end2)
    
    If start2 >= start1 And start2 <= end1 Then
        DateOverlap = 1
    ElseIf end2 >= start1 And end2 <= end1 Then
        DateOverlap = 1
    ElseIf start2 <= start1 And end2 >= start1 Then
        DateOverlap = 1
    ElseIf start2 <= end1 And end2 >= end1 Then
        DateOverlap = 1
    Else
        DateOverlap = 0
    End If
End Function




Function DateOverlapMultiMatchOR(match, searchr, startdater, enddater, startdt, enddt, Optional valuesr As Range = "", Optional arg1 As Variant = "", Optional arg2 As Variant = "", Optional arg3 As Variant = "", Optional arg4 As Variant = "", Optional arg5 As Variant = "") As Integer

    searchsheet = searchr.Parent.Name
    arg1f = 0
    arg2f = 0
    arg3f = 0
    arg4f = 0
    arg5f = 0
        
    For Each cs In searchr
        If IsEmpty(cs) Then
            DateOverlapMultiMatchOR = 0
            Exit For
        End If
        If match = cs Then
            vstart = Worksheets(searchsheet).Cells(cs.row, startdater.Column)
            vend = Worksheets(searchsheet).Cells(cs.row, enddater.Column)
            csvalue = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            'MsgBox (" Match - vstart: " & vstart & " vend: " & vend & " csvalue: " & csvalue)
            If (vend = "") Then
                vend = DateValue(Now)
            End If
            If ((vstart >= startdt And vstart <= enddt) Or (vend >= startdt And vend <= enddt) Or (vstart < startdt And vend > enddt) Or (vstart > startdt And vend < enddt)) Then
                If (csvalue = arg1 Or arg1 = "") Then
                    arg1f = 1
                End If
                If (csvalue = arg2) Then
                    arg2f = 1
                End If
                If (csvalue = arg3) Then
                    arg3f = 1
                End If
                If (csvalue = arg4) Then
                    arg4f = 1
                End If
                If (csvalue = arg5) Then
                    arg5f = 1
                End If
                If (arg1f Or arg2f Or arg3f Or arg4f Or arg5f) Then
                    'MsgBox (" Match Found - vstart: " & vstart & " vend: " & vend & " csvalue: " & csvalue)
                    DateOverlapMultiMatchOR = 1
                    Exit For
                End If
            Else
                'MsgBox (" Match NOT Found - vstart: " & vstart & " vend: " & vend & " csvalue: " & csvalue)
            End If
        End If
    Next
        
End Function


Function Filtered(cell) As Integer
    If (cell.EntireRow.Hidden = True) Then
        Filtered = 1
    Else
        Filtered = 0
    End If
End Function


Function IfGreater(arg, alt)
    If TypeName(arg) = "Range" Then
        arg = arg.Value
    End If
    If IsError(arg) Then
        IfGreater = alt
    ElseIf arg = "" Then
        IfGreater = alt
    ElseIf alt = "" Then
        IfGreater = arg
    ElseIf arg > alt Then
        IfGreater = arg
    Else
        IfGreater = alt
    End If
End Function

Function IfLess(arg, alt)
    If arg = "" Then
        IfLess = alt
    ElseIf alt = "" Then
        IfLess = arg
    ElseIf arg < alt Then
        IfLess = arg
    Else
        IfLess = alt
    End If
End Function

Function IfNot(arg, alt, Optional checknull = True)
On Error Resume Next
    'MsgBox (IsError(arg))
    If checknull Then
        If IsNull(arg) Or IsError(arg) Then
            IfNot = alt
            Exit Function
        End If
    End If
    If Not Application.WorksheetFunction.IsText(arg) Then
        If (arg) Then
            IfNot = arg
        Else
            IfNot = alt
        End If
    Else
        If arg <> "" Then
            IfNot = arg
        Else
            IfNot = alt
        End If
    End If
End Function

Function IfNull(cell, alt, Optional checktype)
    On Error Resume Next
    'vartype(cell) = vartype(alt)
    'MsgBox (VarType(cell))
    'MsgBox (IsError(cell))
    If Not IsMissing(checktype) Then
       If Not vartype(cell) = checktype Then
        cell = ""
       End If
    End If
    If IsNullAlt(cell) Then
        IfNull = alt
    Else
        IfNull = cell
    End If
End Function

Function IfSmart(TEST, iftrue, iffalse, Optional test2, Optional ReturnType = 0)
    'MsgBox (VarType(iffalse))
    On Error Resume Next
    If IsMissing(TEST) Or IsError(TEST) Or IsNullAlt(TEST) Or TEST = "" Then
        IfSmart = iffalse
    ElseIf TEST Then
        If (iftrue = "#THIS#") Then
            IfSmart = TEST
        Else
            IfSmart = iftrue
        End If
    Else
        IfSmart = iffalse
    End If
    If ReturnType = 8 Then
        IfSmart = CStr(IfSmart)
    ElseIf ReturnType = 7 Then
        IfSmart = CDate(IfSmart)
    End If
End Function

Function IfIf(eval, test1, test2, iftrue, iffalse, Optional ReturnType = 0)
    'MsgBox (VarType(iffalse))
    On Error Resume Next
    If IsMissing(eval) Or IsError(eval) Or IsNull(eval) Then
        IfIf = iffalse
    ElseIf eval Then
        If MyEvaluate(eval & test1) Then
            If (test2 = True) Or MyEvaluate(eval & test2) Then
                If (iftrue = "#THIS#") Then
                    IfIf = eval
                Else
                    IfIf = iftrue
                End If
            End If
        End If
    End If
    If IsEmpty(IfIf) Then
        IfIf = iffalse
    End If
    If ReturnType = 8 Then
        IfIf = CStr(IfIf)
    ElseIf ReturnType = 7 Then
        IfIf = CDate(IfIf)
    End If
End Function

Function MyEvaluate(exp As String)
    'On Error Resume Next
    Dim opa() As String, oplt() As String, ope() As String
    Dim opaa() As Variant
    
    opgt = Split(exp, ">")
    oplt = Split(exp, "<")
    ope = Split(exp, "=")
    
    If UBound(opgt) Then
        MyEvaluate = CDbl(opgt(0)) > CDbl(opgt(1))
    ElseIf UBound(oplt) Then
        MyEvaluate = CDbl(oplt(0)) > CDbl(oplt(1))
    ElseIf UBound(ope) Then
        MyEvaluate = CDbl(ope(0)) > CDbl(ope(1))
    Else
        MyEvaluate = 0
    End If
    
End Function

Function IsNullNull(cell)
    If IsNumeric(cell) Then
        IsNullNull = 0
    ElseIf (cell = "NULL" Or cell = "null" Or Len(cell) < 1 Or cell = Null Or IsEmpty(cell)) Then
        IsNullNull = 1
    Else
        IsNullNull = 0
    End If
End Function

Function IsNullAlt(cell, Optional nullType = 0)
    On Error Resume Next
    If IsError(cell) Then
        IsNullAlt = 0
        Exit Function
    End If
    If ctype = TypeName(cell) = "Range" Then
        ctext = cell.Text
        cvalue = cell.Value2
    Else
        ctext = cell
        cvalue = cell
    End If
    If (ctext = "") Then
        If (nullType) And IsEmpty(cvalue) Then
            IsNullAlt = 2
            If IsMissing(cell) Then
                IsNullAlt = 3
            End If
        Else
            IsNullAlt = 1
        End If
    End If
End Function

Function MultiMatchAND(match, searchr, valuesr, arg1, Optional arg2 As Variant = "", Optional arg3 As Variant = "", Optional arg4 As Variant = "", Optional arg5 As Variant = "")
    
    searchsheet = searchr.Parent.Name
    arg1f = 0
    arg2f = 0
    arg3f = 0
    arg4f = 0
    arg5f = 0
    
    For Each cs In searchr
        If IsEmpty(cs) Then
            MultiMatchAND = 0
            Exit For
        End If
        If match = cs Then
            csvalue = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            'MsgBox (" Match Found Searching For: " & arg1 & " " & arg2 & " value: " & csvalue)
            If (csvalue = arg1) Then
                arg1f = 1
            End If
            If (csvalue = arg2 Or arg2 = "") Then
                arg2f = 1
            End If
            If (csvalue = arg3 Or arg3 = "") Then
                arg3f = 1
            End If
            If (csvalue = arg4 Or arg4 = "") Then
                arg4f = 1
            End If
            If (csvalue = arg5 Or arg5 = "") Then
                arg5f = 1
            End If
            If (arg1f And arg2f And arg3f And arg4f And arg5f) Then
                'MsgBox (" Match Found ")
                MultiMatchAND = 1
                Exit For
            End If
        End If
    Next
    
End Function

Function MultiMatchMultiRangeAND(match, matchr, match2, matchr2, valuesr, arg1, Optional arg2 As Variant = "", Optional arg3 As Variant = "", Optional arg4 As Variant = "", Optional arg5 As Variant = "")
    
    searchsheet = matchr.Parent.Name
    arg1f = 0
    arg2f = 0
    arg3f = 0
    arg4f = 0
    arg5f = 0
    
    For Each cs In matchr
        If IsEmpty(cs) Then
            MultiMatchMultiRangeAND = 0
            Exit For
        End If
        If match = cs Then
            cs2 = Worksheets(searchsheet).Cells(cs.row, matchr2.Column)
            csvalue = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            'MsgBox (" Match Found Searching For: " & arg1 & " " & arg2 & " value: " & csvalue)
            If match2 = cs2 Then
                If (csvalue = arg1) Then
                    arg1f = 1
                End If
                If (csvalue = arg2 Or arg2 = "") Then
                    arg2f = 1
                End If
                If (csvalue = arg3 Or arg3 = "") Then
                    arg3f = 1
                End If
                If (csvalue = arg4 Or arg4 = "") Then
                    arg4f = 1
                End If
                If (csvalue = arg5 Or arg5 = "") Then
                    arg5f = 1
                End If
                If (arg1f And arg2f And arg3f And arg4f And arg5f) Then
                    'MsgBox (" Match Found ")
                    MultiMatchMultiRangeAND = 1
                    Exit For
                End If
            End If
        End If
    Next
    
End Function


Function MultiMatchOR(match, searchr, valuesr, arg1, Optional arg2 As Variant = "", Optional arg3 As Variant = "", Optional arg4 As Variant = "", Optional arg5 As Variant = "", Optional arg6 As Variant = "", Optional arg7 As Variant = "")
    
    searchsheet = searchr.Parent.Name
    
    For Each cs In searchr
        If IsEmpty(cs) Then
            MultiMatchOR = 0
            Exit For
        End If
        If match = cs Then
            
            csvalue = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            'MsgBox (" Match Found Searching For: " & arg1 & " " & arg2 & " value: " & csvalue)
            If (csvalue = arg1 Or csvalue = arg2 Or csvalue = arg3 Or csvalue = arg4 Or csvalue = arg5 Or csvalue = arg6 Or csvalue = arg7) Then
                'MsgBox (" Match Found ")
                MultiMatchOR = 1
                Exit For
            End If
        End If
    Next
    
End Function

Function ConvertType(cell, ReturnType)
    On Error Resume Next
    'MsgBox (vartype(cell))
    cellt = cell
    If IsMissing(cell) Or IsError(cell) Or IsNull(cell) Then
        ConvertType = ""
    ElseIf ReturnType = 2 Then
        ConvertType = CInt(cellt)
    ElseIf ReturnType = 5 Then
        ConvertType = CDbl(cellt)
    ElseIf ReturnType = 7 Then
        ConvertType = CDate(cellt)
    ElseIf ReturnType = 8 Then
        ConvertType = CStr(cellt)
    End If
End Function

Function FormatCell(cell, cformat)
    On Error Resume Next
    cell = cell.Text
    cell = Format(cell, cformat)
    FormatCell = cell
End Function

Function DateModified()
    'DateModified = Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time"), "short date")
    Temp = ThisWorkbook.BuiltinDocumentProperties("Last Save Time")
    DateModified = Temp
End Function

Function SplitString(str, token, pos, Optional addEndToken = 0)
    On Error Resume Next
    stra = Split(str, token)
    If stra(0) = str Then
        SplitString = str
        Exit Function
    End If
    'Return last value
    If (pos = -1) Then
        SplitString = stra(UBound(stra))
    'Return everything except last value
    ElseIf pos = -2 Then
        For i = 0 To UBound(stra) - 1
            SplitString = SplitString & stra(i) & token
        Next i
        SplitString = Left(SplitString, UBound(SplitString) - 1)
    'Return everything except first value
    ElseIf pos = -3 Then
       ' SplitString = stra(0)
        For i = 1 To UBound(stra)
            SplitString = SplitString & stra(i) & token
        Next i
        SplitString = Left(SplitString, Len(SplitString) - 1)
    'Return everything except last two values
    ElseIf pos = -4 Then
        If UBound(stra) < 2 Then
            SplitString = str
            Exit Function
        Else
            For i = 0 To UBound(stra) - 2
                SplitString = SplitString & stra(i) & token
            Next i
        End If
        SplitString = Left(SplitString, Len(SplitString) - 1)
    'Return everything except last three values
    ElseIf pos = -5 Then
        If UBound(stra) < 3 Then
            SplitString = str
            Exit Function
        Else
            For i = 0 To UBound(stra) - 3
                SplitString = SplitString & stra(i) & token
            Next i
        End If
        SplitString = Left(SplitString, Len(SplitString) - 1)
    Else
        SplitString = stra(pos)
    End If
    
    If addEndToken = 1 Then
        SplitString = SplitString & token
    End If
End Function



Function StripArray(arr)
    Dim temparr() As Variant
    
    On Error Resume Next
    
    index = 0
    
    For Each av In arr
        If Not av = "" And Not IsNull(av) Then
            temparr(index) = av
            index = index + 1
        End If
    Next
    
    StripArray = temparr

End Function

Function StripArrayCell(arr As String, Optional convert = 0)
    On Error Resume Next
    
    temparr = Split(arr, ";")
    
    index = 0
    For Each av In temparr
        temparr(index) = StripString(av)
        index = index + 1
    Next

    index = 0
    tempstr = ""
    For Each av In temparr
        If Not av = "" Then
            If convert > 0 Then
                av = ConvertType(av, convert)
                av = CStr(av)
            End If
            tempstr = tempstr + av + ";"
            index = index + 1
        End If
    Next
    
    StripArrayCell = tempstr

End Function

Function StripString(str, Optional spaces = 0, Optional zeroreplace = 0, Optional illegalChars = 0)
    On Error Resume Next
    
    If TypeName(illegalChars) <> "Range" Then
        illegalChars = Array("  ", "   ", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "{", "}", "[", "]", "_", "+", "<", ">", "?", "/", "-", ".", "'", ",", ":", ";", " ")
    Else
        illegalChars = RangetoArray(illegalChars)
    End If
    
    str = str.Text
    
    If zeroreplace Then
        For Each ch In illegalChars
            If spaces = 0 Then
                str = replace(str, ch, 0)
            ElseIf spaces = 1 Then
                str = replace(str, ch, 0)
                str = replace(str, " ", 0)
            ElseIf spaces = 2 Then
                str = replace(str, ch, " ")
            End If
        Next
    Else
        For Each ch In illegalChars
            If spaces = 0 Then
                str = replace(str, ch, "")
            ElseIf spaces = 1 Then
                str = replace(str, ch, "")
                str = replace(str, " ", "")
            ElseIf spaces = 2 Then
                str = replace(str, ch, " ")
            End If
        Next
    End If
    
    str = Trim(str)
    str = StripExtraSpaces(str)
    
    StripString = str

End Function

Function StripExtraSpaces(str)
    On Error Resume Next
    
    illegalChars = Array("  ", "   ", "    ")
    str = str.Text
    
    For Each ch In illegalChars
        str = replace(str, ch, " ")
    Next
    
    StripExtraSpaces = str

End Function

Function StripNumber(num)
    On Error Resume Next
    
    letters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    num = StripString(num.Text, 1)
    num = UCase(num)
    
    For Each ch In letters
        num = replace(num, ch, "")
        If spaces Then
            num = replace(num, " ", "")
        End If
    Next
    
    num = CInt(num)
    
    StripNumber = num

End Function

Function StripSentence(str)
    On Error Resume Next
    
    Words = Array("AND", "OR", "OF", "THE")
    str = StripString(str.Text, 1)
    str = UCase(str)
    
    For Each ch In letters
        str = replace(str, ch, "")
        If spaces Then
            str = replace(str, " ", "")
        End If
    Next
    
    str = CInt(str)
    
    StripSentence = str

End Function

Function StringToNumber(str, Optional cutspaces = 0, Optional length = 0, Optional twosided = 0, Optional zeroreplace = 0)
    
    On Error Resume Next
    
    If cutspaces = 0 Then
        letters = Array(" ", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
        str = StripString(str, 0, zeroreplace)
        index = 0
    Else
        letters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
        str = StripString(str, 1, zeroreplace)
        index = 1
    End If
    
    str = UCase(str)
    
    If length > 0 Then
        If twosided Then
            strL = Left(str, length)
            strR = Right(str, length)
            str = strL + strR
        Else
            str = Left(str, length)
        End If
    ElseIf length < 0 Then
        lengthabs = 0 - length
        strlen = Len(str)
        str = Mid(str, strlen / 2, lengthabs)
    End If
    
    For Each ch In letters
        str = replace(str, ch, index)
        index = index + 1
    Next
    
    'str = CDbl(str)
    str = Int(str)
    
    StringToNumber = str

End Function

Function StringPercentMatchWords(str1, str2, Optional row = 0, Optional thresh = 0)
    
    Dim str1a() As String
    Dim str2a() As String
    
    On Error Resume Next
    
    If IsError(str1) Or IsError(str2) Then
        StringPercentMatchWords = 0
        Exit Function
    End If
    
    If (str1 = str2) Then
        StringPercentMatchWords = 1
        Exit Function
    End If
    
    str1t = str1
    str1t = StripString(str1t, 2)
    str1t = UCase(str1t)
    str2t = str2
    str2t = StripString(str2t, 2)
    str2t = UCase(str2t)
    
    str1a = Split(str1t, " ")
    'str1a = StripArray(str1a)
    str1words = ArraySize(str1a) + 1
    str2a = Split(str2t, " ")
    'str2a = StripArray(str2a)
    str2words = ArraySize(str2a) + 1
    
    wordsmatch = 0
    
    For Each str1av In str1a
        For Each str2av In str2a
            If str1av = str2av Then
                wordsmatch = wordsmatch + 2
                Exit For
            End If
        Next
    Next
    
    percent = wordsmatch / (str1words + str2words)
    If (row) Then
        If percent >= thresh Then
            StringPercentMatchWords = row
        Else
            StringPercentMatchWords = 0
        End If
    Else
        StringPercentMatchWords = percent
    End If
End Function

Function StringPercentMatchLetters(str1, str2, Optional row = 0, Optional thresh = 0)
    
    On Error Resume Next
    
    If IsError(str1) Or IsError(str2) Then
        StringPercentMatchLetters = 0
        Exit Function
    End If
    
    If (str1 = str2) Then
        StringPercentMatchLetters = 1
        Exit Function
    End If
    
    letters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    str1t = str1
    str1t = StripString(str1t, 2)
    str1t = UCase(str1t)
    str2t = str2
    str2t = StripString(str2t, 2)
    str2t = UCase(str2t)
    
    str1letters = Len(str1t)
    str2letters = Len(str2t)
    
    lettersden = 0
    lettersnum = 0
    For Each letter In letters
        lettermatch1 = 0
        lettermatch2 = 0
        For index1 = 1 To Len(str1t)
            ch1 = Mid(str1t, index1, 1)
            If ch1 = letter Then
                lettermatch1 = lettermatch1 + 1
            End If
        Next
        For index2 = 1 To Len(str2t)
            If lettermatch2 = lettermatch1 Then
                Exit For
            End If
            ch2 = Mid(str2t, index2, 1)
            If ch2 = letter Then
                lettermatch2 = lettermatch2 + 1
            End If
        Next
        lettersden = lettersden + lettermatch1
        lettersnum = lettersnum + lettermatch2
    Next
    
    'lettersmatch = 0
    'For index1 = 1 To Len(str1)
    '    ch1 = Mid(str1, index1, 1)
    '    For index2 = index1 To Len(str2)
    '        ch2 = Mid(str2, index2, 1)
    '        If ch1 = ch2 Then
    '            lettersmatch = lettersmatch + 1
    '            Exit For
    '        End If
    '    Next
    'Next
    'StringPercentMatch = lettersmatch / (str2letters)
    
    percent = lettersnum / lettersden
    If (row) Then
        If percent >= thresh Then
            StringPercentMatchLetters = row
        Else
            StringPercentMatchLetters = 0
        End If
    Else
        StringPercentMatchLetters = percent
    End If

End Function

Function StringPercentMatch(str1, str2, Optional matchnull = 0, Optional xWord = 0, Optional matchoneword = 0)
    
    Dim str1a() As String
    Dim str2a() As String
    
    On Error Resume Next
    
    If IsError(str1) Or IsError(str2) Then
        StringPercentMatch = 0
        Exit Function
    End If
    
    If matchnull Then
        If Not (IsNullAlt(str1, 1) = IsNullAlt(str2, 1)) Then
            StringPercentMatch = 0
            Exit Function
        End If
    End If
    
    If (str1 = str2) Then
        StringPercentMatch = 1
        Exit Function
    End If
    
    str1t = str1
    str1t = StripString(str1t, 2)
    str1t = UCase(str1t)
    str2t = str2
    str2t = StripString(str2t, 2)
    str2t = UCase(str2t)
    
    If xWord Then
        xWord = StripString(xWord, 2)
        xWord = UCase(xWord)
        str1ta = Split(str1t, " ")
        str2ta = Split(str2t, " ")
        str1t = ""
        str2t = ""
        For Each str1twords In str1ta
            If Not str1twords = xWord Then
                str1t = str1t + str1twords + " "
            End If
        Next
        For Each str2twords In str2ta
            If Not str2twords = xWord Then
                str2t = str2t + str2twords + " "
            End If
        Next
        str1t = Trim(str1t)
        str2t = Trim(str2t)
    End If
    
    If (str1t = str2t) Then
        StringPercentMatch = 1
        Exit Function
    End If
    
    str1a = Split(str1t, " ")
    str2a = Split(str2t, " ")
    str1words = ArraySize(str1a) + 1
    str2words = ArraySize(str2a) + 1
    
    wordsmatch = 0
    
    For Each str1av In str1a
        For Each str2av In str2a
            If str1av = str2av Then
                wordsmatch = wordsmatch + 2
                Exit For
            End If
        Next
    Next
    
    wordsmatchpct = wordsmatch / (str1words + str2words)
    
    If (wordsmatchpct = 1) Then
        StringPercentMatch = wordsmatchpct
    ElseIf (wordsmatchpct > 0.5) Then
        wordsmatch = 0
        For Each str1av In str1a
            For Each str2av In str2a
                If (StringPercentMatchLetters(str1av, str2av)) >= 0.9 Then
                    wordsmatch = wordsmatch + 2
                    Exit For
                End If
            Next
        Next
        wordsmatchpct = wordsmatch / (str1words + str2words)
    End If
    
    StringPercentMatch = wordsmatchpct

End Function

Function ConcatenateArrays(arr1, arr2)
    On Error Resume Next
    
    arr1 = Split(arr1, ";")
    arr2 = Split(arr2, ";")
    tempstr = ""
    For Each str1 In arr1
        For Each str2 In arr2
            If Not str1 = "" And Not str2 = "" Then
                tempstr = tempstr + str1 + str2 + ";"
            End If
        Next
    Next
    
    ConcatenateArrays = tempstr

End Function


Public Function FindMatch(sorted As Boolean, searchr As Range, match, Optional searchr2 As Range, Optional match2)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    
    FindMatch = 0
    Max = searchr.count
    'For Each cs In searchr
    For i = 1 To Max
        cs = Worksheets(searchsheet).Cells(i, searchr.Column)
        If IsEmpty(cs) Then
            Exit Function
        End If
        matchv = match.Value
        csv = CDbl(cs)
        If IsEmpty(csv) Then
            csv = -1
        End If
        If sorted And csv > matchv Then
            Exit Function
        End If
        cst = CStr(cs)
        matcht = CStr(match)
        If StrComp(cs, match, vbTextCompare) = 0 Then
            cs2 = Worksheets(searchsheet).Cells(i, searchr2.Column)
            If StrComp(cs2, match2, vbTextCompare) = 0 Then
                FindMatch = 1
                Exit Function
            End If
        End If
    Next i
    
End Function

Public Function MyCountIFs(searchr As Range, match, Optional searchr2 As Range, Optional match2)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    
    MyCountIFs = 0
    Max = searchr.count
    'For Each cs In searchr
    For i = 1 To Max
        cs = Worksheets(searchsheet).Cells(i, searchr.Column)
        If IsEmpty(cs) Then
            Exit For
        End If
        cst = CStr(cs)
        matcht = CStr(match)
        If StrComp(cs, match, vbTextCompare) = 0 Then
            cs2 = Worksheets(searchsheet).Cells(i, searchr2.Column)
            If StrComp(cs2, match2, vbTextCompare) = 0 Then
                MyCountIFs = MyCountIFs + 1
            End If
        End If
    Next i
    
End Function

Public Function CountIFP(searchr, match, Optional percent = 1)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    MatchFound = False
    CountIFP = 0
    If vartype(match) = vbString Then
        match = StripString(UCase(match), 2)
    End If
    
    For Each cs In searchr
        If Filtered(cs) = 0 Then
            If IsNullAlt(cs, 1) = 2 Then
                Exit For
            End If
            csr = cs.row
            If vartype(match) = vbString Then
                cs = StripString(UCase(cs.Text), 2)
            End If
            If StringPercentMatch(match, cs) >= percent Then
                MatchFound = True
                CountIFP = CountIFP + 1
            End If
        End If
    Next
    
End Function

Public Function CountIFSP(searchr, match, Optional searchr2, Optional match2, Optional searchr3, Optional match3, Optional percent = 1)
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    MatchFound = False
    CountIFSP = 0
    If vartype(match) = vbString Then
        match = StripString(UCase(match), 2)
        match2 = StripString(UCase(match2), 2)
        match3 = StripString(UCase(match3), 2)
    End If
    
    For Each cs In searchr
        If Filtered(cs) = 0 Then
            If IsNullAlt(cs, 1) = 2 Then
                Exit For
            End If
            csr = cs.row
            cs2 = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
            cs3 = Worksheets(searchsheet).Cells(cs.row, searchr3.Column)
            If vartype(match) = vbString Then
                cs = StripString(UCase(cs.Text), 2)
                cs2 = StripString(UCase(cs2), 2)
                cs3 = StripString(UCase(cs3), 2)
            End If
            If StringPercentMatch(match, cs) >= percent Then
                If StringPercentMatch(match2, cs2) >= percent Then
                    If StringPercentMatch(match3, cs3) >= percent Then
                        MatchFound = True
                        CountIFSP = CountIFSP + 1
                    End If
                End If
            End If
        End If
    Next
    
End Function

Public Function CountIFSPx(searchr, match, Optional searchr2, Optional match2, Optional searchr3, Optional match3, Optional percent = 1, Optional xWord = 0)
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    MatchFound = False
    CountIFSPx = 0
    If vartype(match) = vbString Then
        match = StripString(UCase(match), 2)
        match2 = StripString(UCase(match2), 2)
        match3 = StripString(UCase(match3), 2)
    End If
    
    For Each cs In searchr
        If Filtered(cs) = 0 Then
            If IsNullAlt(cs, 1) = 2 Then
                Exit For
            End If
            csr = cs.row
            cs2 = Worksheets(searchsheet).Cells(cs.row, searchr2.Column)
            cs3 = Worksheets(searchsheet).Cells(cs.row, searchr3.Column)
            If vartype(match) = vbString Then
                cs = StripString(UCase(cs.Text), 2)
                cs2 = StripString(UCase(cs2), 2)
                cs3 = StripString(UCase(cs3), 2)
            End If
            If StringPercentMatch(match, cs, 0, xWord) >= percent Then
                If StringPercentMatch(match2, cs2, 0, xWord) >= percent Then
                    If StringPercentMatch(match3, cs3, 0, xWord) >= percent Then
                        MatchFound = True
                        CountIFSPx = CountIFSPx + 1
                    End If
                End If
            End If
        End If
    Next
    
End Function



Function ExtractZip(Address)
    On Error Resume Next
    
    addressa = Split(Address, " ")
    zipindex = UBound(addressa)
    zipraw = addressa(zipindex)
    zipa = Split(zipraw, "-")
    ExtractZip = zipa(0)

End Function

Public Function RangetoArray(sRange, Optional vartype = 2, Optional skip1 = 1, Optional allowduplicates = 0, Optional allowzero = 1, Optional allownull = 0, Optional rangesize, Optional startindex = 0)
    On Error Resume Next
    
    Dim tempa() As Variant
    ReDim tempa(startindex To startindex)
    
    If IsMissing(rangesize) Then
        'Try different approaches to get range size, check row size and count a
        rangesizea = Application.WorksheetFunction.CountA(sRange)
        rangesizeb = sRange.Cells(Rows.count, 1).End(xlUp).row
        rangesizec = sRange.Cells.SpecialCells(xlCellTypeLastCell).row
        rangesized = sRange.Cells.SpecialCells(xlCellTypeLastCell).Column
        'Use largest size
        rangesize = Application.WorksheetFunction.Max(rangesizea, rangesizeb, rangesizec, rangesized)
        'If size is excessively large, near limit, use counta
        If rangesize > 1000000 Then
            rangesize = rangesizea
        End If
    End If
    
    'i = 1
    n = startindex
    
    For i = 1 To rangesize
    'For Each cs In sRange
        If vartype = 2 Then cs = CStr(sRange.Cells(i, 1).Value2)
        If vartype = 1 Then cs = CDbl(sRange.Cells(i, 1).Value2)
        'If i > rangesize Then
        '    Exit For
        'End If
        inarray = False
        If ((skip1 = 1 And i > 1) Or (skip1 = 0)) And (allownull Or cs <> "") And (allowzero Or (cs <> 0 And cs <> "0")) Then
            If allowduplicates = 0 Then
                For Each inval In tempa
                    If inval = cs Then
                        inarray = True
                    End If
                Next
            End If
            If Not inarray Or allowduplicates Then
                ReDim Preserve tempa(startindex To n)
                tempa(n) = cs
                n = n + 1
            End If
        End If
        'i = i + 1
    Next i
    
    RangetoArray = tempa
    
End Function



Public Function RangetoString(sRange, Optional vartype = 2, Optional skip1 = 1, Optional rangesize = -1, Optional startrow = 1)
    'On Error Resume Next
    On Error GoTo ErrHandler
    Dim tempa() As Variant
    
    myDelimiter = ","
    
    If Not IsArray(sRange) Or TypeName(sRange) = "Range" Then
        tempa = RangetoArray(sRange, vartype, skip1)
    Else
        tempa = sRange
    End If
    
    startrow = startrow - 1
    If rangesize = -1 Then
        rangesize = UBound(tempa)
    Else
        rangesize = startrow + rangesize - 1
    End If
    
    Dim temps As String
    For i = startrow To rangesize
        If i > rangesize Then
            Exit For
        End If
        cs = tempa(i)
        If (removeZero And cs = 0) Or cs = "" Then
            
        Else
            If vartype = 2 Then
                temps = temps & "'" & cs & "'" & myDelimiter
            ElseIf vartype = 1 Then
                temps = temps & cs & myDelimiter
            End If
        End If
    Next
    
    temps = Left(temps, Len(temps) - 1)
    tempslen = Len(temps)
    
    RangetoString = temps
    
    Exit Function
    
ErrHandler:
    Temp = Err.Description
    If Temp <> "" Then
        MsgBox (Temp)
    End If
    
End Function

Public Function RangestoString(sRange, Optional vartype = 2, Optional skip1 = 1, Optional sRange2, Optional myDelimiter = ",", Optional removeZero = 0, Optional endDelimiter = 0)
    'On Error Resume Next
    On Error GoTo ErrHandler
    
    Dim tempa() As Variant
    i = 1
    n = 0
    allowzero = 1
    If removeZero = 1 Then
        allowzero = 0
    End If
    
    If Not IsArray(sRange) Or TypeName(sRange) = "Range" Then
        tempa = RangetoArray(sRange, vartype, skip1, 1)
    Else
        tempa = sRange
    End If
    
    If Not IsMissing(sRange2) Then
        If Not IsArray(sRange2) Or TypeName(sRange) = "Range" Then
            tempb = RangetoArray(sRange2, vartype, skip1, 1, allowzero)
        Else
            tempb = sRange2
        End If
        tempa = MergeArrays(tempa, tempb, 1)
    Else
        'If NumberOfArrayDimensions(tempa) > 1 Then
            tempa = Flatten2DArray(tempa)
        'End If
        tempa = RemoveArrayDuplicates(tempa)
    End If
    
    Dim temps As String
    For Each cs In tempa
        If removeZero And cs = 0 Then
            
        Else
            If vartype = 2 Then
                temps = temps & "'" & cs & "'" & myDelimiter
            ElseIf vartype = 1 Then
                temps = temps & cs & myDelimiter
            End If
        End If
    Next
    
    temps = Left(temps, Len(temps) - 1)
    tempslen = Len(temps)
    
    If endDelimiter Then
        temps = temps & myDelimiter
    End If
    
    RangestoString = temps
    
    Exit Function
    
ErrHandler:
    Temp = Err.Description
    If Temp <> "" Then
        MsgBox (Temp)
    End If
    
End Function

'TODO: Clean up or remove this function. Not sure it serves any purpose.
Public Function RangetoStringIf(sRange As Range, sRange2 As Range, condition, Optional vartype = 2, Optional skip1 = 1, Optional myDelimiter = ",", Optional removeZero = 0, Optional endDelimiter = 0)
    'On Error Resume Next
    
    'temp = sRange2.Cells(1, 1)
    
    Dim tempa() As Variant
    i = 1
    n = 0
    'rangesize = srange.End(xlDown).row
    rangesize = Application.WorksheetFunction.CountA(sRange)
    For Each cs In sRange
        If i > rangesize Then
            Exit For
        End If
        'temp = sRange2.Cells(i, 1)
        If ((skip1 = 1 And i > 1) Or (skip1 = 0)) And (sRange2.Cells(i, 1) = condition) Then
            ReDim Preserve tempa(n)
                If IsEmpty(cs.Text) Then
                    If cs.Value <> "" Then
                        tempa(n) = cs.Value
                    End If
                Else
                    If cs.Text <> "" Then
                        tempa(n) = cs.Text
                    End If
                End If
            n = n + 1
        End If
        i = i + 1
    Next
    
    tempa = RemoveArrayDuplicates(tempa)
    
    Dim temps As String
    For Each cs In tempa
        If removeZero And cs = 0 Then
            
        Else
            If vartype = 2 Then
                temps = temps & "'" & cs & "'" & myDelimiter
            ElseIf vartype = 1 Then
                temps = temps & cs & myDelimiter
            End If
        End If
    Next
    
    temps = Left(temps, Len(temps) - 1)
    tempslen = Len(temps)
    
    If endDelimiter Then
        temps = temps & myDelimiter
    End If
    
    'RangestoString = tempa
    
    RangetoStringIf = temps
    
End Function

Function MyCompare(arg1, arg2, Optional nullv = -99)
    If IsNullAlt(arg1) Or IsNullAlt(arg2) Then
        MyCompare = nullv
    ElseIf arg1 > arg2 Then
        MyCompare = -1
    ElseIf arg1 = arg2 Then
        MyCompare = 0
    ElseIf arg1 < arg2 Then
        MyCompare = 1
    End If
End Function

Public Function DateToQuarter(myDate As Date, Optional alt = 0)
    If IsError(myDate) Or (myDate = "6/28/1900") Then
        DateToQuarter = ""
        Exit Function
    End If
    myYear = Year(myDate)
    myMonth = Month(myDate)
    myQ = Round(myMonth / 3 + 0.3)
    
    If (alt) Then
        myQ = myQ + 1
        If (myQ = 5) Then
            myQ = 1
            myYear = myYear + 1
        End If
    End If
    DateToQuarter = CInt(myYear & myQ)
End Function

Public Function IsProvider(mycell As String)
    If (Left(mycell, 1) = "(") Then
        IsProvider = True
    Else
        IsProvider = False
    End If
End Function

Public Function TestString(mycell As String, mylimiter As String, location As Integer)
    'temp = Mid(mycell, location, 1)
    If (Mid(mycell, location, 1) = mylimiter) Then
        TestString = mycell
    Else
        TestString = False
    End If
End Function

Public Function CopyWhenTrueSave(savec)
    mySheet = savec.Parent.Name
    
    If savec = False Then
        i = 1
        Do While prev = False
            prev = Worksheets(mySheet).Cells(savec.row - i, savec.Column)
            i = i + 1
        Loop
        CopyWhenTrueSave = prev
    Else
        CopyWhenTrueSave = savec
    End If
End Function

Public Function ExtractFromString(str As String, lim1 As String, lim2 As String, Optional startposf = 1, Optional extractf = 1, Optional trimlim = 1)
    Dim starti As Integer
    Dim endi As Integer
    
    For i = 1 To Len(str)
        If startposf = 1 Then
            Temp = Mid(str, i, 1)
            chari = i
        ElseIf startposf = -1 Then
            Temp = Mid(str, Len(str) - i + 1, 1)
            chari = Len(str) - i + 1
        End If
        If (Temp = lim1) And starti = 0 Then
            starti = chari
        ElseIf (Temp = lim2) And endi = 0 Then
            endi = chari
        End If
        If starti > 0 And endi > 0 Then
            Exit For
        End If
    Next i
    
    extract = Mid(str, starti, endi - starti + 1)
    
    If extractf Then
        ExtractFromString = extract
        If trimlim Then
            Temp = Split(ExtractFromString, lim1)
            Temp = Split(Temp(1), lim2)
            ExtractFromString = Temp(0)
        End If
    Else
        Temp = Split(str, extract)
        If startposf = 1 Then
            ExtractFromString = Trim(Temp(1))
        Else
            ExtractFromString = Trim(Temp(0))
        End If
    End If
    
End Function


Function StringToArray(starr As String, Optional delim = ";", Optional TrimResults = 0)

    tempa = Split(starr, delim)
    If TrimResults Then
        For i = 0 To UBound(tempa)
            tempa(i) = Trim(tempa(i))
        Next i
    End If
    StringToArray = tempa
    
End Function

Function SearchArrayMatch(cell, searchr As Range, valuesr As Range)

    searchcells = Application.WorksheetFunction.CountA(searchr)

    For Each cs In searchr
        If i > searchcells Then
            Exit For
        End If
        i = i + 1
        Temp = Split(cs, ";")
        For Each aVal In Temp
            'If RawValue(cell) = RawValue(aval) Then
            If CStr(cell) = CStr(aVal) Or cell = aVal Then
                match = Worksheets(valuesr.Parent.Name).Cells(cs.row, valuesr.Column)
                Exit For
            End If
        Next
        If Not IsEmpty(match) Then
            Exit For
        End If
    Next
    
    SearchArrayMatch = match

End Function

Function MatchBool(range1 As Range, val1, Optional range2 As Range, Optional val2, Optional range3 As Range, Optional val3, Optional rangesize)
    On Error Resume Next
    If IsMissing(rangesize) Then
        rangesize = range1.count
    End If
    MatchBool = 0
    For i = 2 To rangesize
        range1val = Worksheets(range1.Parent.Name).Cells(i, range1.Column)
        If range1val > val1 Then
            Exit Function
        End If
        If range1val = val1 Then
            range2val = Worksheets(range2.Parent.Name).Cells(i, range2.Column)
            If range2val = val2 Then
                range3val = Worksheets(range3.Parent.Name).Cells(i, range3.Column)
                If range3val = val3 Then
                    MatchBool = 1
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Function RemoveSpecialCharacters(rng, Optional spaceReplace = 0, Optional rc = False, Optional rw)
    Dim a$, b$, c$, i As Integer
    a$ = rng
    For i = 1 To Len(a$)
        b$ = Mid(a$, i, 1)
        If b$ Like "[A-Z,a-z,0-9]" And b$ <> "," Then
            c$ = c$ & b$
        ElseIf rc And b$ = rc Then
            c$ = c$ & rw
        ElseIf spaceReplace Then
            c$ = c$ & " "
        End If
    Next i
    RemoveSpecialCharacters = c$
End Function

Function TrimArray(MyArr)
    ReDim newarr(LBound(MyArr) To UBound(MyArr))
    For i = LBound(MyArr) To UBound(MyArr)
        If MyArr(i) <> "" Then
            newarr(j) = MyArr(i)
            j = j + 1
        End If
    Next i
    ReDim Preserve newarr(LBound(MyArr) To j)
    TrimArray = newarr
End Function
    
Function CleanString(str, Optional allCaps = 0, Optional removeSC = 1, Optional spaceReplace = 1, Optional cleanSpaces = 1)
    If removeSC Then
        str = RemoveSpecialCharacters(str, spaceReplace)
    End If
    If cleanSpaces Then
        stra = Split(str, "  ")
        For Each part In stra
            If part <> "" Then
                Temp = Temp & " " & part
            End If
        Next
        str = Temp
    End If
    If allCaps Then
        str = UCase(str)
    End If
    CleanString = Trim(str)
End Function

Function StringInCell(str, cell, Optional Delimiter = "/")
    
    stra = Split(str, Delimiter)
    cella = Split(cell)
    
    For Each word In cella
        For Each match In stra
            If word = match Then
                StringInCell = True
                Exit Function
            End If
        Next
    Next
    
    StringInCell = False

End Function

Function ArraySearchGet(cell, searchr, valuesr)
    On Error Resume Next
    'Temp = Split(cell, ";")
    searchsheet = searchr.Parent.Name
    'ArraySearchGet = ""
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            Exit For
        End If
        Temp = Split(cs, ";")
        match = cell
        For Each cst In Temp
            If CStr(match) = CStr(cst) Or match = cst Then
                ArraySearchGet = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
                Exit Function
            End If
        Next
    Next

End Function

Function BoolString(stringr As Range, boolr As Range, Optional token = ",")
    Dim bs As String
    
    stringa = ConvertRangetoArray(stringr)
    boola = ConvertRangetoArray(boolr)
    
    For i = LBound(boola) To UBound(boola)
        If boola(i) Then
            bs = bs & token & stringa(i)
        End If
    Next i
    
    'For Each cell In boolr
    
    BoolString = bs
End Function

Public Function ConvertRangetoArray(myRange As Range)
    On Error Resume Next
    
    Dim tempa() As Variant
    i = 0
    n = 0
    For Each cs In myRange
        ReDim Preserve tempa(n)
        tempa(n) = cs.Value
        n = n + 1
    Next
    
    ConvertRangetoArray = tempa
    
End Function

Function NumericMatch(matchCell, searchRange As Range, flex As Integer)
    On Error Resume Next
    
    emptycount = 0
    
    For Each cs In searchRange
        If IsEmpty(cs) Then
            emptycount = emptycount + 1
            If emptycount >= 5 Then
                NumericMatch = Error
                Exit Function
            End If
        End If
        If IsNumeric(cs.Value) Then
            If matchCell <= cs.Value + flex And matchCell >= cs.Value - flex Then
                NumericMatch = cs.row
                Exit Function
            End If
        End If
    Next
    
    NumericMatch = Error

End Function

Function ExtractIntegers(str, token)

    stra = Split(str, token)
    Result = ""
    
    For Each part In stra
        If IsNumeric(part) Then
            Result = Result & part & ","
        End If
    Next
    
    ExtractIntegers = Left(Result, Len(Result) - 1)

End Function

Function IsBold(ByVal cell As Range) As Boolean
    IsBold = cell.Font.Bold
End Function

Function FontColor(ByVal cell As Range)
    FontColor = cell.Font.ColorIndex
End Function

Function InString(ByVal str, ByVal match, Optional nospaces = 1)
    On Error Resume Next
    If nospaces = 0 Then
        str = " " & str & " "
        match = " " & match & " "
    End If
    stra = Split(str, match)
    If (stra(0) = str Or str = 0) Then
        InString = 0
    Else
        InString = 1
    End If
End Function

Function IfZero(cell, alt)
    On Error Resume Next
    If (cell = 0) Then
        IfZero = alt
    Else
        IfZero = cell
    End If
End Function

Function RangePosition(cell, pRange As Range)
    matches = 0
    cRow = cell.row
    For Each rv In pRange
        rvRow = rv.row
        If rvRow > cRow Then
            Exit For
        End If
        If rv = cell Then
            matches = matches + 1
        End If
    Next
    RangePosition = matches
End Function

Function QuarterSubtract(q1, q2)
On Error Resume Next
    y1 = Left(q1, 4)
    y2 = Left(q2, 4)
    m1 = Right(q1, 1)
    m2 = Right(q2, 1)
    
    calc1 = (y1 - y2) * 4
    calc2 = m1 - m2
    QuarterSubtract = calc1 + calc2

End Function

Function QuarterAdd(q1, numq)
    
    If IsError(q1) Then
        Exit Function
    End If
    
    y1 = Left(q1, 4)
    m1 = Right(q1, 1)
    
    qstoadd = m1 + numq
    qsmod = qstoadd - (4 * Int(qstoadd / 4))
    If qsmod = 0 Then
        newq = 4
    Else
        newq = qsmod
    End If
    If qstoadd > 0 Then
        yearstoadd = Application.WorksheetFunction.RoundDown((qstoadd - 1) / 4, 0)
    Else
        yearstoadd = Application.WorksheetFunction.RoundDown((qstoadd - newq) / 4, 0)
    End If
    
    QuarterAdd = Int((y1 + yearstoadd) & newq)
    
End Function

Function ArraysMerge(arr1, arr2)
    On Error Resume Next
    
    Dim arr3() As Variant
    Dim arr4() As Variant
    Dim sortarr As Object
    Dim cell As Range
    
    arr1 = Split(arr1, ";")
    arr2 = Split(arr2, ";")
    arr3 = MergeArrays(arr1, arr2)
    arr4 = RemoveArrayDuplicates(arr3)
    
    Set sortarr = CreateObject("System.Collections.ArrayList")
    
    ' Initialise the ArrayList, for instance by taking values from a range:
    For Each aVal In arr4
        sortarr.Add aVal
    Next
    sortarr.Sort
    
    tempstr = ""
    For Each str1 In sortarr
        If Not str1 = "" Then
            tempstr = tempstr + str1 + ";"
        End If
    Next
    
    ArraysMerge = tempstr

End Function

Function ArraySearchGetAll(cell, searchr, valuesr)
    On Error Resume Next
    'Temp = Split(cell, ";")
    searchsheet = searchr.Parent.Name
    'ArraySearchGet = ""
    tempstr = ""
    
    For Each cs In searchr
        If (IsEmpty(cs)) Then
            Exit For
        End If
        Temp = Split(cs, ";")
        match = cell
        For Each cst In Temp
            If CStr(match) = CStr(cst) Or match = cst Then
                tempstr = tempstr & Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
            End If
        Next
    Next
    
    ArraySearchGetAll = tempstr

End Function

Function ArrayGetMatchValuesCriteria(cell As String, searchr, valuesr, Optional sizer As Range, Optional criterian = 0, Optional owBlank = 1)
    'Criteria
    '0 = Get first match
    '1 = Get last match
    '2 = Get smallest match
    '3 = Get largest match
    '4 = Get TRUE match
    aTemp = Split(cell, ";")
    searchsheet = searchr.Parent.Name
    MatchFound = False
    i = 0
    j = 1
    If Not sizer Is Nothing Then
        rangesize = sizer.End(xlDown).row
    Else
        rangesize = 0
    End If
    
    For Each cs In searchr
        If rangesize = 0 Then
            If IsEmpty(cs) Then
                Exit For
            End If
        ElseIf j > rangesize Then
            Exit For
        End If
        j = j + 1
        For Each match In aTemp
            If CStr(match) = CStr(cs) Or match = cs Then
                compareTemp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)

                If MatchFound Then
                    If owBlank And compareTemp = "" Then
                        'ArrayGetMatchValuesCriteria = "Blank Found"
                        'Exit Function
                    Else
                        Select Case criterian
                            Case 0
                                ArrayGetMatchValuesCriteria = compareTemp
                                Exit Function
                            Case 1
                                CopyMatchTemp = Worksheets(searchsheet).Cells(cs.row, valuesr.Column)
                            Case 2
                                If compareTemp < CopyMatchTemp Or IsEmpty(CopyMatchTemp) Then
                                    CopyMatchTemp = compareTemp
                                End If
                            Case 3
                                If compareTemp > CopyMatchTemp Or IsEmpty(CopyMatchTemp) Then
                                    CopyMatchTemp = compareTemp
                                End If
                            Case 4
                                If compareTemp Or IsEmpty(CopyMatchTemp) Then
                                    CopyMatchTemp = compareTemp
                                End If
                            Case Else
                                CopyMatchTemp = "ERROR - Criteria Number out of Range"
                        End Select
                    End If
                Else
                    MatchFound = True
                    CopyMatchTemp = compareTemp
                    If criterian = 0 Or (criterian = 4 And CopyMatchTemp = True) Then
                        If owBlank And CopyMatchTemp = "" Then
                        Else
                            ArrayGetMatchValuesCriteria = compareTemp
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    Next
    
    If IsEmpty(CopyMatchTemp) Then
        CopyMatchTemp = ""
    End If
    
    ArrayGetMatchValuesCriteria = CopyMatchTemp

End Function

Function LastTrue(searchr As Range, Optional getcolumn = 1)
    For Each cs In searchr
        If cs = True Then
            If getcolumn Then
                LastTrue = cs.Column
            Else
                LastTrue = cs.row
            End If
        End If
    Next
End Function

Function GetCell(R, c, src As String, Optional wbn As String)
    Dim wb As Workbook
    
    'If IsNull(wbn) Then
    If wbn = "" Then
        Set wb = ThisWorkbook
    Else
        Set wb = Workbooks(wbn)
    End If
    
    GetCell = wb.Worksheets(src).Cells(R, c)

End Function

Function eval(Ref As String)
    Application.Volatile
    eval = Evaluate(Ref)
End Function

Function fileName()
    fileName = GetCurrentFilename
End Function

Function GetCurrentFilename()
    GetCurrentFilename = ThisWorkbook.Name
End Function

Function IfNeg(arg, alt)
    If arg < 0 Then
        IfNeg = alt
    Else
        IfNeg = arg
    End If
End Function

Function IfNotPos(arg, alt)
    If arg <= 0 Then
        IfNotPos = alt
    Else
        IfNotPos = arg
    End If
End Function

Function IfLessThan(arg, thres, alt)
    If arg < thres Then
        IfLessThan = alt
    Else
        IfLessThan = arg
    End If
End Function

Function QuarterToDate(date1)
    year1 = Left(date1, 4)
    month1 = (Right(date1, 1) * 3) - 2
    day1 = 1
    QuarterToDate = CDate(month1 & "-" & day1 & "-" & year1)
End Function

'Deprecated Use QuarterAdd Instead
Function AddQuarters(Dont, Use)
    'Deprecated Use QuarterAdd Instead
    AddQuarters = QuarterAdd(Dont, Use)
End Function

Public Function MaxIF(maxr As Range, searchr As Range, match As Variant, Optional sorted = 0)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    
    MaxIF = 0
    Max = WorksheetFunction.CountA(searchr)
    For i = 1 To Max
        cs = Worksheets(searchsheet).Cells(i, searchr.Column)
        If IsEmpty(cs) Then
            Exit For
        End If
        If sorted And i > 1 And MatchFound Then
            If cs <> match Then
                Exit For
            End If
        End If
        If cs = match Then
            MatchFound = True
            cv = Worksheets(searchsheet).Cells(i, maxr.Column)
            If cv > MaxIF Then
                MaxIF = cv
            End If
        ElseIf MatchFound And sorted And i > 1 And cs <> match Then
            Exit For
        End If
    Next i
    
End Function

Public Function MaxIFs(maxr As Range, searchr As Range, match As Variant, Optional sorted = 0, Optional factor1 = 0, Optional factorr1, Optional factor2 = 0, Optional factorr2)
    
    On Error Resume Next
    
    searchsheet = searchr.Parent.Name
    
    MaxIFs = 0
    Max = WorksheetFunction.CountA(searchr)
    For i = 1 To Max
        cs = Worksheets(searchsheet).Cells(i, searchr.Column)
        If IsEmpty(cs) Then
            Exit For
        End If
        If sorted And i > 1 And MatchFound Then
            If cs <> match Then
                Exit For
            End If
        End If
        If cs = match Then
            MatchFound = True
            cv = Worksheets(searchsheet).Cells(i, maxr.Column)
            If cv > MaxIFs Then
                If factor1 Then
                    f1v = Worksheets(searchsheet).Cells(i, factorr1.Column)
                    'test1 = factor1 = f1v
                    'test2 = Evaluate(f1v & factor1)
                    If factor1 = f1v Or Evaluate(f1v & factor1) Then
                        If factor2 Then
                            f2v = Worksheets(searchsheet).Cells(i, factorr2.Column)
                            If factor2 = f2v Or Evaluate(f2v & factor2) Then
                                MaxIFs = cv
                            End If
                        Else
                            MaxIFs = cv
                        End If
                    End If
                Else
                    MaxIFs = cv
                End If
            End If
        ElseIf MatchFound And sorted And i > 1 And cs <> match Then
            Exit For
        End If
    Next i
    
End Function

Function MedianAt(MedianStart, MedianEnd, CheckRange, MedianRange)
    For Each cell In CheckRange
        If cell = MedianStart Then
            StartCol = cell.Column
        End If
        If cell = MedianEnd Then
            EndCol = cell.Column
            Exit For
        End If
    Next
    
    Dim MedianRangeSize As Integer
    MedianRangeSize = EndCol - StartCol
    Dim NewMedianRange() As Variant
    ReDim NewMedianRange(MedianRangeSize)
    
    For Each cell In MedianRange
        If cell.Column >= StartCol And cell.Column <= EndCol Then
            NewMedianRange(cell.Column - StartCol) = cell
        End If
        If cell.Column > EndCol Then
            Exit For
        End If
    Next
    
    MedianAt = Application.WorksheetFunction.Median(NewMedianRange)
End Function

Function MinNotZero(ParamArray rng() As Variant)
    Dim tmp As Variant
    tmp = rng(0)
    For Each cell In rng
        If cell < tmp And cell <> 0 Then
            tmp = cell
        End If
    Next
    MinNotZero = tmp
End Function

Function MinNotZeroAlt(MinRange As Range)

    Dim sh As Worksheet
    Dim rn As Range
    Set sh = MinRange.Worksheet

    minc = 2147483647
    ccount = 0
    Set rn = sh.UsedRange
    NumCells = rn.Rows.count + rn.row - 1
    For Each c In MinRange
        ccount = ccount + 1
        If Not IsError(c) Then
            If c.Value > 0 And c.Value < minc Then
                minc = c.Value
            End If
            If ccount > NumCells Then
                Exit For
            End If
        End If
    Next
    
    'MsgBox ("Completed in: " & ccount)
    
    MinNotZeroAlt = minc
End Function


Function MaxAlt(MaxRange As Range, Optional ForceNumeric = 0)

    Dim sh As Worksheet
    Dim rn As Range
    Set sh = MaxRange.Worksheet

    maxc = 0
    ccount = 0
    Set rn = sh.UsedRange
    NumCells = rn.Rows.count + rn.row - 1
    For Each c In MaxRange
        ccount = ccount + 1
        Temp = c.Value
        If Not IsError(c) Then
            If ForceNumeric = 0 Then
                If c.Value > maxc Then
                    maxc = c.Value
                End If
            Else
                If (IsNumeric(c.Value) Or IsDate(c.Value)) And c.Value > maxc Then
                    maxc = c.Value
                End If
            End If
        End If
        If ccount > NumCells Then
            Exit For
        End If
    Next
    
    'MsgBox ("Completed in: " & ccount)
    
    MaxAlt = maxc
End Function

Function IsInArray(val As Variant, arr As Variant, Optional uppercase = 0) As Boolean
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If uppercase Then
            If UCase(element) = UCase(val) Then
                IsInArray = True
                Exit Function
            End If
        Else
            If element = val Then
                IsInArray = True
                Exit Function
            End If
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Function IsStringInArray(tocheck As Variant, arr As Variant, Optional exact = 0) As Boolean
    If Not IsArray(arr) Then
        arr = StringToArray(CStr(arr))
    End If
    
    tempstr = CStr(tocheck)
    
    For Each thing In arr
        If exact Then
            If thing = tempstr Then
                IsStringInArray = True
                Exit Function
            End If
        Else
            If StripString(Trim(UCase(thing))) = StripString(Trim(UCase(tempstr))) Then
                IsStringInArray = True
                Exit Function
            End If
        End If
    Next
    IsStringInArray = False
End Function

Function IsStringInArrayIndex(tocheck As Variant, arr() As String, Optional exact = 0, Optional uppercase = 0) As Integer
    i = 0
    For Each thing In arr
        If exact Then
            If thing = tocheck Then
                IsStringInArrayIndex = i
                Exit Function
            End If
        ElseIf uppercase Then
            If UCase(thing) = UCase(tocheck) Then
                IsStringInArrayIndex = i
                Exit Function
            End If
        Else
            If StripString(Trim(UCase(thing))) = StripString(Trim(UCase(tocheck))) Then
                IsStringInArrayIndex = i
                Exit Function
            End If
        End If
        i = i + 1
    Next
    IsStringInArrayIndex = -1
End Function

Function ReplaceWord(ByVal str As String, ByVal word As String, Optional ByVal ReplaceWith As String = "", Optional DEPRECATED_USE_REPLACESTRING) As String
    ReplaceWord = ReplaceString(str, word, ReplaceWith)
End Function

Function ReplaceStringBySplit(ByVal str As String, ByVal word As String, Optional ByVal ReplaceWith As String = "") As String
    tempstra = Split(str, " ")
    NumResults = UBound(tempstra)
    If NumResults = 0 Then
        tempstr = str
    Else
        For i = 0 To UBound(tempstra)
            If UCase(Trim(tempstra(i))) = UCase(Trim(word)) Then
                tempstr = Trim(tempstr & " " & ReplaceWith)
            Else
                tempstr = Trim(tempstr & " " & Trim(tempstra(i)))
            End If
        Next
    End If
    ReplaceStringBySplit = tempstr
End Function

Function ReplaceString(ByVal str As String, ByVal word As String, Optional ByVal ReplaceWith As String = "", Optional ignoreCase = 0, Optional nospaces = 0, Optional startcheck = 0, Optional wholeWordCheck = 0) As String
    If ignoreCase Then
        str = UCase(str)
        word = UCase(word)
        ReplaceWith = UCase(ReplaceWith)
    End If
    If nospaces = 0 Then
        str = " " & str & " "
        word = " " & word & " "
        ReplaceWith = " " & ReplaceWith & " "
    End If
    If wholeWordCheck = 1 Then
        testa = Split(str, word)
        If testa(0) = str Then
            ReplaceString = str
            Exit Function
        End If
        Dim LeftChar As String
        Dim RightChar As String
        LeftChar = Right(testa(0), 1)
        RightChar = Left(testa(1), 1)
        If IsLetter(LeftChar) Or IsLetter(RightChar) Then
            ReplaceString = str
            Exit Function
        End If
    End If
    If startcheck = 1 Then
        tempstr = Trim(replace(str, " " & word, " " & ReplaceWith))
        If tempstr = str Then tempstr = Trim(replace(str, "," & word, "," & ReplaceWith))
        If tempstr = str Then tempstr = Trim(replace(str, "." & word, "." & ReplaceWith))
        If tempstr = str Then tempstr = Trim(replace(str, word & ",", ReplaceWith & ","))
        If tempstr = str Then tempstr = Trim(replace(str, word & ".", ReplaceWith & "."))
        'If tempstr = str Then tempstr = Trim(replace(str, "." & word & ",", "." & replacewith))
    Else
        tempstr = Trim(replace(str, word, ReplaceWith))
    End If
    ReplaceString = tempstr
End Function

Function ReplaceArrayInString(ByVal str As String, replaceA As Variant, Optional ReplaceWith As String = "", Optional ignoreCase = 0, Optional nospaces = 0) As String
    Dim tempstr As String
    tempstr = str
    If Not IsArray(replaceA) Then
        replaceA = StringToArray(CStr(replaceA))
    End If
    If TypeName(replaceA) = "Range" Then
        replaceA = RangetoArray(replaceA)
    End If
    For i = 0 To UBound(replaceA)
        tempstr = ReplaceString(tempstr, CStr(replaceA(i)), ReplaceWith, ignoreCase, nospaces)
    Next
    ReplaceArrayInString = tempstr
End Function

Function ReplaceArrayInStringWithArray(ByVal str As String, replaceA As Variant, Optional ReplaceWith As Variant, Optional SQlConvert = 0, Optional AddQualifyer = "", Optional ignoreCase = 0, Optional nospaces = 0, Optional startcheck = 0, Optional wholeWordCheck = 0) As String
    Dim tempstr As String
    tempstr = str
    If Not IsArray(replaceA) Then
        replaceA = StringToArray(CStr(replaceA))
    End If
    If TypeName(replaceA) = "Range" Then
        replaceA = RangetoArray(replaceA)
    End If
    If Not IsArray(ReplaceWith) Then
        ReplaceWith = StringToArray(CStr(ReplaceWith))
    End If
    If TypeName(ReplaceWith) = "Range" Then
        ReplaceWith = RangetoArray(ReplaceWith)
    End If
    For i = 0 To UBound(replaceA)
        If i = 88 Then
            Temp = 0
        End If
        If SQlConvert = 1 And ReplaceWith(i) <> "" Then
            tempstr = ReplaceString(tempstr, AddQualifyer & CStr(replaceA(i)), CStr(replaceA(i)) & " AS " & ReplaceWith(i), ignoreCase, nospaces, startcheck, wholeWordCheck)
        ElseIf SQlConvert = 2 And ReplaceWith(i) <> "" Then
            tempstr = ReplaceString(tempstr, AddQualifyer & CStr(replaceA(i)), "AS " & ReplaceWith(i), ignoreCase, nospaces, startcheck, wholeWordCheck)
        Else
            tempstr = ReplaceString(tempstr, AddQualifyer & CStr(replaceA(i)), ReplaceWith(i), ignoreCase, nospaces, startcheck, wholeWordCheck)
        End If
    Next i
    ReplaceArrayInStringWithArray = tempstr
End Function

Function WordCount(str, Optional token = " ")
    On Error Resume Next
    stra = Split(Trim(str), token)
    WordCount = 0
    For i = 0 To UBound(stra)
        If stra(i) <> " " And stra(i) <> "" Then
            WordCount = WordCount + 1
        End If
    Next i
End Function

Function RemoveApostrophes(str As Variant, Optional spaces = 0, Optional zeroreplace = 0)
    On Error Resume Next
    
    illegalChars = Array("'")
    str = str.Text
    
    If zeroreplace Then
        For Each ch In illegalChars
            If spaces = 0 Then
                str = replace(str, ch, 0)
            ElseIf spaces = 1 Then
                str = replace(str, ch, 0)
                str = replace(str, " ", 0)
            ElseIf spaces = 2 Then
                str = replace(str, ch, " ")
            End If
        Next
    Else
        For Each ch In illegalChars
            If spaces = 0 Then
                str = replace(str, ch, "")
            ElseIf spaces = 1 Then
                str = replace(str, ch, "")
                str = replace(str, " ", "")
            ElseIf spaces = 2 Then
                str = replace(str, ch, " ")
            End If
        Next
    End If
    
    str = Trim(str)
    'str = StripExtraSpaces(str)
    
    RemoveApostrophes = str

End Function

Function AreValuesEqual(val1 As Variant, val2 As Variant, MatchType As String, Optional matchcertainty, Optional IgnoreWords, Optional ForceMatchWords, Optional ReplaceWords, Optional ReplaceWith) As Boolean
    AreValuesEqual = False
    If val1 = val2 And Not IsEmpty(val1) And Not val1 = "" Then
        AreValuesEqual = True
        Exit Function
    End If
    If MatchType = "Exact" Then
        Exit Function
    End If
    If MatchType = "Words" Then
        MatchAmount = StringPercentMatchWords(val1, val2, 0, 0, 0)
        If MatchAmount >= matchcertainty Then
            AreValuesEqual = True
        End If
    End If
    If MatchType = "Words-Contains" Then
        MatchAmount = StringPercentMatchWords(val1, val2, 0, 0, 1)
        If MatchAmount >= matchcertainty Then
            AreValuesEqual = True
        End If
    End If
    If MatchType = "Letters" Then
        MatchAmount = StringPercentMatchAllLetters(val1, val2)
        If MatchAmount >= matchcertainty Then
            AreValuesEqual = True
        End If
    End If
End Function

Function RemoveDuplicateWords(str As String, Optional replace As String = "") As String
    tempstra = Split(str, " ")
    NumResults = UBound(tempstra)
    If NumResults = 0 Then
        tempstr = str
    Else
        For i = 0 To UBound(tempstra)
            For j = i + 1 To UBound(tempstra)
                If UCase(Trim(tempstra(i))) = UCase(Trim(tempstra(j))) Then
                    tempstra(j) = ""
                End If
            Next
        Next
        For i = 0 To UBound(tempstra)
            If tempstra(i) <> "" Then
                tempstr = tempstr & " " & Trim(tempstra(i))
            End If
        Next
        tempstr = Trim(tempstr)
    End If
    RemoveDuplicateWords = tempstr
End Function

Function RemoveExtraSpacesFormulas(cc As String)
    cc = replace(cc, vbTab, " ")
    cc = replace(cc, "   ", " ", 1, -1, vbTextCompare)
    cc = replace(cc, "  ", " ", 1, -1, vbTextCompare)
    cc = replace(cc, "  ", " ", 1, -1, vbTextCompare)
    cc = Trim(cc)
    RemoveExtraSpacesFormulas = cc
End Function

Function CollectionHasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    CollectionHasKey = (Err.Number = 0)
    Err.Clear
End Function

Function WorksheetExists(shtName, Optional wb) As Boolean
    Dim sht As Worksheet

    If IsMissing(wb) Then Set wb = ThisWorkbook
    If TypeName(wb) <> "Workbook" Then
        Set wb = Workbooks(wb)
    End If
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function Flatten2DArray(arr, Optional limit = 0, Optional stopatnull = 0)
    Dim tempa()
    vals = 0
    
    If NumberOfArrayDimensions(arr) = 2 Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                If stopatnull And arr(i, j) = "" Then Exit For
                ReDim Preserve tempa(vals)
                tempa(vals) = arr(i, j)
                'test2 = Application.Transpose(Application.index(arr, 0, 2))
                vals = vals + 1
                If limit And i >= limit Then Exit For
            Next j
        Next i
        Flatten2DArray = tempa
    Else
        Flatten2DArray = arr
    End If
    
End Function

Public Function NumberOfArrayDimensions(arr As Variant) As Integer
Dim ErrChk As Long
Dim i%, x%
 
    On Error Resume Next
    For i = 1 To 62
        ErrChk = LBound(arr, i)
        If Err.Number = 9 Then
            x = i - 1
            Exit For
        End If
    Next i
 
    NumberOfArrayDimensions = x

End Function

Public Function GetLastHeaderColumn(Optional sheetname, Optional startrow = 1)
    If IsMissing(sheetname) Then
        Set theSheet = ActiveSheet
    Else
        Set theSheet = Worksheets(sheetname)
    End If
    RangeString = startrow & ":" & startrow
    GetLastHeaderColumn = WorksheetFunction.CountA(theSheet.Range(RangeString))
End Function

Public Function GetLastDataRow(Optional sheetname, Optional StartColumn = "A")
    If IsMissing(sheetname) Then
        Set theSheet = ActiveSheet
    Else
        Set theSheet = Worksheets(sheetname)
    End If
    RangeString = StartColumn & ":" & StartColumn
    GetLastDataRow = WorksheetFunction.CountA(theSheet.Range(RangeString))
End Function

Public Function GetFormulaStartColumn(Optional sheetname, Optional startrow = 2, Optional WorkbookName, Optional theSheet)
    If IsMissing(theSheet) Then
        If IsMissing(WorkbookName) Or WorkbookName = "" Then
            WorkbookName = ThisWorkbook.Name
            Set theWorkbook = ThisWorkbook
        Else
            WorkbookName = SplitString(WorkbookName, "\", -1)
            Set theWorkbook = Workbooks(WorkbookName)
        End If
        If IsMissing(sheetname) Or sheetname = "" Then
            Set theSheet = ActiveSheet
        Else
            Set theSheet = theWorkbook.Worksheets(sheetname)
        End If
    End If
    numcols = WorksheetFunction.CountA(theSheet.Range("1:1"))
    For i = 1 To numcols
        Temp = theSheet.Cells(startrow, i)
        If theSheet.Cells(startrow, i).HasFormula Then
            GetFormulaStartColumn = i
            Exit Function
        End If
    Next i
    GetFormulaStartColumn = 0
End Function

Public Function ConvertArrayToCollection(TheArray) As Collection
    Dim oColl As Collection
    Set oColl = New Collection
    For i = LBound(TheArray) To UBound(TheArray)
        arrval = TheArray(i)
        oColl.Add i, arrval
    Next i
    Set ConvertArrayToCollection = oColl
End Function

Function ColumnNumberToLetter(lngCol)
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    ColumnNumberToLetter = vArr(0)
End Function

Function ReplaceCharacters(str, replace, ReplaceWith, Optional ignoreCase = 0)
    If ignoreCase Then
        Temp = Split(UCase(str), UCase(replace))
    Else
        Temp = Split(str, replace)
    End If
    If UBound(Temp) > 0 Then
        tempstr = Temp(0)
        For i = 1 To UBound(Temp)
            tempstr = tempstr & ReplaceWith & Temp(i)
        Next i
    Else
        tempstr = str
    End If
    ReplaceCharacters = tempstr
End Function

Function Get1DArrayFrom2DArray(arr, i)
    Dim newarr()
    arrstart = LBound(arr, 2)
    arrsize = UBound(arr, 2)
    ReDim newarr(arrsize)
    For j = 0 To arrsize
        newarr(j) = arr(i, j)
    Next j
    Get1DArrayFrom2DArray = newarr
End Function

Function DateToTimestamp(theDate, Optional theFormat)
    'DateToTimestamp = TEXT(Year(theDate), "0000") & "-" & Month(theDate) & "-" & Day(theDate) & " 00:00:00'"
    DateToTimestamp = Format(theDate, "yyyy-mm-dd") & " 00:00:00"
    If theFormat = "MSSQL" Then
        DateToTimestamp = "{ts '" & DateToTimestamp & "'}"
    End If
    Temp = 0
End Function

Function CheckObjectState(obj)
    If obj Is Nothing Then
        CheckObjectState = "Nothing"
    ElseIf IsEmpty(obj) Then
        CheckObjectState = "Empty"
    'ElseIf obj Is Empty Then
    '    CheckObjectState = "Empty"
    ElseIf IsNull(obj) Then
        CheckObjectState = "Null"
    'ElseIf obj Is Null Then
        'CheckObjectState = "Null"
    'ElseIf obj.Len > 0 Then
        'CheckObjectState = "Something"
    'ElseIf TypeName(obj) Then
        
    Else
        CheckObjectState = "Unknown"
    End If
End Function

Function RemoveDoubleQuotes(cc As String)
    cc = replace(cc, Chr(34) & Chr(34), Chr(34), 1, -1, vbTextCompare)
    RemoveDoubleQuotes = cc
End Function

Function EscapeApostrophe(cc As String)
    cc = replace(cc, "'", "''", 1, -1, vbTextCompare)
    EscapeApostrophe = cc
End Function

Function CharacterCount(word, findchar)
    replacechar = ""
    CharacterCount = Len(word) - Len(replace(word, findchar, replacechar))
End Function

Function GenerateMatchString(str1, Optional IgnoreWords, Optional ForceMatchWords, Optional ReplaceWords, Optional ReplaceWith, Optional Artifacts, Optional ReplaceCharacters, Optional bRemoveApostrophes = 1, Optional RemoveCharacters, Optional bReplaceWholeWord = 0)
    
    Dim str1a() As String
    Dim str1t As String
    Dim cw As String
    
    'Debug
    If str1 = "Associate of Applied Science, Information Technology, Cisco" Then
        Temp = 0
    End If
    
    'Set Current String to Uppercase
    str1t = UCase(Trim(str1))
    'Clear extra spaces
    str1t = replace(str1t, "   ", " ")
    str1t = replace(str1t, "  ", " ")
    str1t = Trim(str1t)
    
    'Remove Artifacts
    If Not IsMissing(Artifacts) Then
        If Not IsArray(Artifacts) Then
            Artifacts = StringToArray(CStr(Artifacts))
        End If
        If TypeName(Artifacts) = "Range" Then
            Artifacts = RangetoArray(Artifacts)
        End If
        For i = 0 To UBound(Artifacts)
            cw = UCase(Artifacts(i))
            If InString(str1t, cw) Then
                str1t = ReplaceString(str1t, CStr(cw), "", 1, 1)
                str1t = Trim(str1t)
            End If
        Next
    End If
    
    'Clear extra spaces
    str1t = replace(str1t, "   ", " ")
    str1t = replace(str1t, "  ", " ")
    str1t = Trim(str1t)
    
    'Force Matches
    If Not IsMissing(ForceMatchWords) Then
        If Not IsArray(ForceMatchWords) Then
            ForceMatchWords = StringToArray(CStr(ForceMatchWords))
        End If
        If TypeName(ForceMatchWords) = "Range" Then
            ForceMatchWords = RangetoArray(ForceMatchWords)
        End If
        For Each word In ForceMatchWords
            word = UCase(Trim(word))
            If InString(str1t, word) = 1 Then
                GenerateMatchString = word
                Exit Function
            End If
        Next
    End If
    
    'Replacements
    If Not IsMissing(ReplaceWords) Then
        If Not IsArray(ReplaceWords) Then
            ReplaceWords = StringToArray(CStr(ReplaceWords))
        End If
        If TypeName(ReplaceWords) = "Range" Then
            ReplaceWords = RangetoArray(ReplaceWords)
        End If
        If Not IsArray(ReplaceWith) Then
            ReplaceWith = StringToArray(CStr(ReplaceWith))
        End If
        If TypeName(ReplaceWith) = "Range" Then
            ReplaceWith = RangetoArray(ReplaceWith)
        End If
        For i = 0 To UBound(ReplaceWords)
            cw = CStr(UCase(Trim(ReplaceWords(i))))
            If InString(str1t, cw) Then
                If bReplaceWholeWord Then
                    GenerateMatchString = ReplaceWith(i)
                    Exit Function
                End If
                'Remove problematic characters before replacements
                If Not IsMissing(RemoveCharacters) Then
                    str1t = ReplaceArrayInString(str1t, RemoveCharacters, "", 1, 1)
                    cw = ReplaceArrayInString(cw, RemoveCharacters, "", 1, 1)
                End If
                str1t = ReplaceString(str1t, cw, CStr(ReplaceWith(i)), 1)
                'str1t = replace(str1t, UCase(CStr(ReplaceWords(i))), UCase(ReplaceWith(i)))
                str1t = Trim(str1t)
            End If
        Next
    End If
    
    'Remove any Duplicated Words
    str1t = RemoveDuplicateWords(str1t)
    
    'Clear extra spaces
    str1t = replace(str1t, "   ", " ")
    str1t = replace(str1t, "  ", " ")
    str1t = Trim(str1t)
    
    'Strings to ignore
    If Not IsMissing(IgnoreWords) Then
        If Not IsArray(IgnoreWords) Then
            IgnoreWords = StringToArray(CStr(IgnoreWords))
        End If
        If TypeName(IgnoreWords) = "Range" Then
            IgnoreWords = RangetoArray(IgnoreWords)
        End If
        For Each word In IgnoreWords
            word = UCase(Trim(word))
            If InString(str1t, word) Then
                str1t = ReplaceString(str1t, CStr(word))
            End If
        Next
    End If
    
    'Remove Apostrophes
    If bRemoveApostrophes Then
        str1t = RemoveApostrophes(str1t)
    End If
    
    'Replace problem characters with spaces
    If Not IsMissing(ReplaceCharacters) Then
        If Not IsArray(ReplaceCharacters) Then
            ReplaceCharacters = StringToArray(CStr(ReplaceCharacters))
        End If
        If TypeName(ReplaceCharacters) = "Range" Then
            ReplaceCharacters = RangetoArray(ReplaceCharacters)
        End If
        For i = 0 To UBound(ReplaceCharacters)
            cw = ReplaceCharacters(i)
            If InString(str1t, cw) Then
                str1t = ReplaceString(str1t, CStr(cw), " ", 1, 1)
                str1t = Trim(str1t)
            End If
        Next
    End If
    
    'Remove other problem characters
    If Not IsMissing(RemoveCharacters) Then
        If Not IsArray(RemoveCharacters) Then
            RemoveCharacters = StringToArray(CStr(RemoveCharacters))
        End If
        If TypeName(RemoveCharacters) = "Range" Then
            RemoveCharacters = RangetoArray(RemoveCharacters)
        End If
        For i = 0 To UBound(RemoveCharacters)
            cw = RemoveCharacters(i)
            If InString(str1t, cw) Then
                str1t = ReplaceString(str1t, CStr(cw), "", 1, 1)
                str1t = Trim(str1t)
            End If
        Next
    End If
    
    'Clear extra spaces
    str1t = replace(str1t, "   ", " ")
    str1t = replace(str1t, "  ", " ")
    str1t = Trim(str1t)
    
    'Second Replacement
    If Not IsMissing(ReplaceWords) Then
        For i = 0 To UBound(ReplaceWords)
            cw = Trim(UCase(ReplaceWords(i)))
            If InString(str1t, cw) Then
                If bReplaceWholeWord Then
                    GenerateMatchString = ReplaceWith(i)
                    Exit Function
                End If
                str1t = ReplaceString(str1t, CStr(cw), CStr(ReplaceWith(i)), 1)
                'str1t = replace(str1t, UCase(CStr(ReplaceWords(i))), UCase(ReplaceWith(i)))
                str1t = Trim(str1t)
            End If
        Next
    End If
    
    'Remove any Duplicated Words
    str1t = RemoveDuplicateWords(str1t)
    
    'Clear extra spaces
    str1t = replace(str1t, "   ", " ")
    str1t = replace(str1t, "  ", " ")
    str1t = Trim(str1t)

    'Second Ignore
    If Not IsMissing(IgnoreWords) Then
        For Each word In IgnoreWords
            word = UCase(Trim(word))
            If InString(str1t, word) Then
                str1t = ReplaceString(str1t, CStr(word))
            End If
        Next
    End If
    
    'Clear extra spaces
    str1t = replace(str1t, "   ", " ")
    str1t = replace(str1t, "  ", " ")
    str1t = Trim(str1t)
    
    'Second Force Matches
    If Not IsMissing(ForceMatchWords) Then
        For Each word In ForceMatchWords
            word = UCase(Trim(word))
            If InString(str1t, word) = 1 Then
                GenerateMatchString = word
                Exit Function
            End If
        Next
    End If
    
    'Send null if we were trying to replace keywords. At this point there was no match.
    If bReplaceWholeWord Then
        str1t = ""
    End If
    
    GenerateMatchString = str1t
    
End Function

Function CleanStringWithReplace(ByVal str1, ReplaceWords, ReplaceWith, Optional ReplaceStrength = 0, Optional ReplaceStrengthValues, Optional bRemoveApostrophes = 1, Optional bReplaceWholeWord = 0, Optional bReplaceNull = 0)
    Dim str1a() As String
    Dim str1t As String
    Dim str2t As String
    Dim cw As String
    
    'Debug
    If str1 = "Secondary school diploma or its equivalent" Then
    'If str1 = "A.A.S. in Drafting" Then
        Temp = 0
    End If
    
    'Remove Apostrophes
    'If bRemoveApostrophes Then
    '    str1 = RemoveApostrophes(str1)
    'End If
    
    'Set Current String to Uppercase and Clear extra spaces
    str1t = str1
    str1t = replace(str1t, "    ", " ")
    str1t = replace(str1t, "   ", " ")
    str1t = replace(str1t, "  ", " ")
    str1t = CStr(UCase(Trim(str1t)))
    
    'Remove artifacts and special characters
    
    For j = 0 To UBound(ReplaceWords)
        If ReplaceStrengthValues(j) = 1 Then
            cw = CStr(UCase(Trim(ReplaceWords(j))))
            If InString(str1t, cw) Then
                str1t = ReplaceString(str1t, cw, CStr(ReplaceWith(j)), 1, 1)
                str1t = replace(str1t, "    ", " ")
                str1t = replace(str1t, "   ", " ")
                str1t = replace(str1t, "  ", " ")
                str1t = Trim(str1t)
            End If
        End If
    Next j
    
    str2t = str1t
    'For i = ReplaceStrength To 2 Step -1
    For i = 2 To ReplaceStrength
        For j = 0 To UBound(ReplaceWords)
            If ReplaceStrengthValues(j) <= i Then
                cw = CStr(UCase(Trim(ReplaceWords(j))))
                
                'Debug
                If cw = "SECONDARY SCHOOL" Then
                    Temp = 0
                End If
                
                If InString(str2t, cw, 0) Then
                    If bReplaceWholeWord And ReplaceStrengthValues(j) = ReplaceStrength Then
                        CleanStringWithReplace = ReplaceWith(j)
                        Exit Function
                    End If
                    str2t = ReplaceString(str2t, cw, CStr(ReplaceWith(j)), 1, 0)
                    'Remove any Duplicated Words
                    str2t = RemoveDuplicateWords(str2t)
                    'Clear extra spaces
                    str1t = replace(str1t, "    ", " ")
                    str2t = replace(str2t, "   ", " ")
                    str2t = replace(str2t, "  ", " ")
                    str2t = Trim(str2t)
                End If
            End If
        Next j
    Next i
    
    If bReplaceNull Then
        CleanStringWithReplace = ""
    ElseIf str2t = "" Then
        CleanStringWithReplace = str1t
    Else
        CleanStringWithReplace = str2t
    End If
    
End Function

Function TrimLast(ByVal str)
    TrimLast = Left(str, Len(str) - 1)
End Function

Function GetAvailableFields(ByVal TableName, ByVal DBFullName, Optional ByVal IncludeTable = 0, Optional ByVal Delimiter = ",", Optional ByVal ConnString = "", Optional Enclose = 0)
On Error GoTo ErrHandler

    bNoErrors = False
    
    If DBFullName = "" Then
        DBFullName = ThisWorkbook.FullName
    End If
    
    If ConnString <> "" Then
        DSN = ParseConnectionStringDSN(ConnString)
    End If
    
    If ConnString = "" Or DSN = "Excel Files" Then
        ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source='" & DBFullName & "';" & _
            "Extended Properties=""Excel 12.0 Macro;HDR=Yes;FMT=Delimited;IMEX=1;"";"
    Else
        ConnString = SplitString(ConnString, "ODBC;", -1)
    End If
    
    If DSN = "twist_adhoc" Then
        DBFullName = "twist_adhoc.dbo"
        Query = "SELECT TOP 1 * FROM " & DBFullName & "." & TableName
    ElseIf DSN = "mock_twist_adhoc" Then
        DBFullName = "twist_adhoc.dbo"
        Query = "SELECT TOP 1 * FROM " & TableName
    Else
        Query = "SELECT TOP 1 * FROM [" & TableName & "$]"
    End If
    
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 2
    cn.CommandTimeout = 0
    cn.Open ConnString
    Set rs = cn.Execute(Query)
    Do While Not rs.EOF
      For i = 0 To rs.Fields.count - 1
        'Debug.Print rs.Fields(i).Name, rs.Fields(i).Value
        'strNaam = rs.Fields(0).Value
        If IncludeTable = 1 Then
            If Enclose Then
                Temp = Temp & TableName & "." & "[" & rs.Fields(i).Name & "]" & Delimiter
            Else
                Temp = Temp & TableName & "." & rs.Fields(i).Name & Delimiter
            End If
        ElseIf IncludeTable = 2 Then
            If Enclose Then
                Temp = Temp & "[" & rs.Fields(i).Name & "]" & Delimiter
                Temp = Temp & TableName & "." & "[" & rs.Fields(i).Name & "]" & Delimiter
            Else
                Temp = Temp & rs.Fields(i).Name & Delimiter
                Temp = Temp & TableName & "." & rs.Fields(i).Name & Delimiter
            End If
        Else
            If Enclose Then
                Temp = Temp & "[" & rs.Fields(i).Name & "]" & Delimiter
            Else
                Temp = Temp & rs.Fields(i).Name & Delimiter
            End If
        End If
      Next
      rs.MoveNext
    Loop
    
    GetAvailableFields = Temp
    
    bNoErrors = True
    
ErrHandler:
    
    If Not bNoErrors Then
        MsgBox (Err.Description)
    End If
    If Not IsEmpty(rs) Then
        rs.Close
    End If
    If Not cn Is Nothing And Not cn.State = 0 Then
        cn.Close
    End If
    Set rs = Nothing
    Set cn = Nothing

End Function

Function ExtractSQLFieldFromFormula(ByVal str)
    Temp = str
    Temp = Trim(SplitString(Temp, ",", 0))
    Temp = Trim(SplitString(Temp, "(", -3))
    Temp = Trim(SplitString(Temp, ")", 0))
    'temp = Trim(SplitString(temp, "AS", 0))
    Temp = Trim(SplitString(Temp, "=", 0))
    Temp = Trim(SplitString(Temp, ">", 0))
    Temp = Trim(SplitString(Temp, "<", 0))
    Temp = Trim(SplitString(Temp, "BETWEEN", 0))
    'temp = Trim(SplitString(temp, "IIF(", -1))
    Temp = Trim(SplitString(Temp, ",", 0))
    ExtractSQLFieldFromFormula = Temp
End Function

Function ExtractFormula(ByVal str, Optional includeAs = 1)
    openparens = 0
    closeparens = 0
    doparse = True
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        If char = "[" Then doparse = False
        If char = "]" Then doparse = True
        If doparse Then
            If char = "(" Then
                openparens = openparens + 1
                If openparens = 1 Then
                    firstopenparens = i
                End If
            End If
            If char = ")" Then closeparens = closeparens + 1
            If openparens = closeparens And openparens > 0 And closeparens > 0 Then
                lastchar = i
                If includeAs Then
                    For k = lastchar To Len(str)
                        char3 = Mid(str, k, 1)
                        If char3 = "," Then
                            lastchar = k - 1
                            Exit For
                        ElseIf k = Len(str) Then
                            lastchar = k
                            Exit For
                        End If
                    Next k
                End If
                For j = firstopenparens To 1 Step -1
                    char2 = Mid(str, j, 1)
                    If char2 = " " Or char2 = vbLf Or char2 = "," Then
                        firstchar = j + 1
                        Exit For
                    ElseIf j = 1 Then
                        firstchar = j
                        Exit For
                    End If
                Next j
            End If
            If Not IsEmpty(firstchar) Then
                Exit For
            End If
        End If
    Next i
    If openparens = 0 And closeparens = 0 Then
        ExtractFormula = 0
    Else
        ExtractFormula = Mid(str, firstchar, lastchar - firstchar + 1)
    End If
End Function

Function ExtractSQLFormulaFieldsFromQuery(ByVal str, Optional includePlainFields = 1, Optional includeFormulaNames = 0, Optional includeFormulaFields = 0, Optional includeFormulas = 0)
    Temp = ""
    loopcount = 0
    Dim formulasa() As String
    Do While Temp <> 0
        Temp = ExtractFormula(str)
        If Temp <> 0 Then
            str = replace(str, Temp, "")
            Formulas = Formulas & Temp & ", "
            ReDim Preserve formulasa(loopcount)
            formulasa(loopcount) = Temp
        End If
        loopcount = loopcount + 1
        If loopcount > 1000000 Then
            MsgBox ("Error (ExtractSQLFormulaFieldsFromQuery): Infinite Loop Detected")
            Exit Function
        End If
    Loop
    str = Trim(replace(str, vbLf, " "))
    str = Trim(replace(str, ",,", ","))
    str = Trim(replace(str, ", ,", ","))
    If includePlainFields Then
        cleanquery = str
    End If
    If IsArrayAllocated(formulasa) Then
        numformulas = UBound(formulasa)
    Else
        numformulas = -1
    End If
    If includeFormulaNames Or includeFormulaFields Then
        For i = 0 To numformulas
            If includeFormulaFields Then
                ff = Trim(ExtractSQLFieldFromFormula(formulasa(i)) & ",")
            End If
            If includeFormulaNames Then
                fn = Trim(SplitString(formulasa(i), " AS ", -1)) & ","
            End If
            cleanquery = cleanquery & ff & fn
        Next i
    End If
    If includeFormulas = 1 Then
        cleanquery = cleanquery & Formulas
    End If
    ExtractSQLFormulaFieldsFromQuery = cleanquery
End Function

Function RemoveSQLOperators(ByVal str, Optional replacement = " ")
'On Error GoTo ErrHandler
    Dim Temp As String
    Temp = str
    Temp = replace(Temp, vbLf, replacement)
    Temp = replace(Temp, "&", replacement)
    Temp = replace(Temp, "(", replacement)
    Temp = replace(Temp, ")", replacement)
    Temp = replace(Temp, "[", replacement)
    Temp = replace(Temp, "]", replacement)
    RemoveSQLOperators = Temp
'ErrHandler:
'    MsgBox ("Error (RemoveSQLOperators):" & Err.Description)
End Function

Function RemoveSQLCommandsFromQuery(ByVal str)
    Temp = " " & str & " "
    Temp = replace(Temp, "IIF", " IIF")
    Temp = replace(Temp, vbLf, " ")
    Temp = replace(Temp, "    ", " ")
    Temp = replace(Temp, "   ", " ")
    Temp = replace(Temp, "  ", " ")
    selectTopEnd = InStr(Temp, " TOP ")
    If selectTopEnd > 0 Then
        Temp = Right(Temp, Len(Temp) - (selectTopEnd + 6))
    End If
    Temp = replace(Temp, " SELECT DISTINCT ", "")
    Temp = replace(Temp, " SELECT ", "")
    Temp = replace(Temp, " DISTINCT ", "")
    Temp = replace(Temp, " UNION ALL ", "")
    Temp = replace(Temp, " UNION ", "")
    fromStart = InStr(Temp, " FROM ")
    whereStart = InStr(Temp, " WHERE ")
    If whereStart = 0 Then whereStart = InStr(Temp, " GROUP BY ")
    If whereStart = 0 Then whereStart = InStr(Temp, " ORDER BY ")
    If whereStart = 0 Then whereStart = Len(Temp)
    whereString = Right(Temp, Len(Temp) - whereStart)
    tablesString = Mid(Temp, fromStart, whereStart - fromStart)
    Temp = replace(Temp, tablesString, ",")
    joinStart = InStr(tablesString, " JOIN ")
    joinEnd = InStr(tablesString, " ON ")
    If joinStart > 0 Then
        tablesString = Right(tablesString, Len(tablesString) - (joinEnd + 2))
        joinFields = SplitString(tablesString, " JOIN ", 0)
        joinStart = InStr(tablesString, " JOIN ")
        joinEnd = InStr(tablesString, " ON ")
        If joinStart > 0 Then
            tablesString = Right(tablesString, Len(tablesString) - (joinEnd + 2))
            joinFields = joinFields & SplitString(tablesString, " JOIN ", 0)
        End If
        joinFields = replace(joinFields, " AND ", ",")
        joinFields = replace(joinFields, "=", ",")
    End If
    whereFieldsa = Split(whereString, " AND ")
    For Each wf In whereFieldsa
        wf = SplitString(wf, " IN (", 0)
        wf = SplitString(wf, " IN(", 0)
        wf = SplitString(wf, " BETWEEN ", 0)
        wf = SplitString(wf, " LIKE ", 0)
        wf = SplitString(wf, ">", 0)
        wf = SplitString(wf, "<", 0)
        wf = SplitString(wf, "=", 0)
        whereFields = whereFields & wf & ","
    Next wf
    Temp = replace(Temp, whereString, ",")
    Temp = Temp & ", " & joinFields
    Temp = Temp & ", " & whereFields
    Temp = replace(Temp, " LEFT ", ",")
    Temp = replace(Temp, " RIGHT ", ",")
    Temp = replace(Temp, " OUTER ", "")
    Temp = replace(Temp, " INNER ", "")
    Temp = replace(Temp, " JOIN ", "")
    Temp = replace(Temp, " WHERE ", "")
    Temp = replace(Temp, " GROUP BY ", ",")
    Temp = replace(Temp, " ORDER BY ", ",")
    Temp = replace(Temp, " ASC ", ",")
    Temp = replace(Temp, " DESC ", ",")
    Temp = replace(Temp, "(", "")
    Temp = replace(Temp, ")", "")
    Temp = replace(Temp, "    ", " ")
    Temp = replace(Temp, "   ", " ")
    Temp = replace(Temp, "  ", " ")
    Temp = replace(Temp, " , ", ",")
    Temp = replace(Temp, ", ", ",")
    Temp = replace(Temp, " ,", ",")
    Temp = replace(Temp, ",,,,", ",")
    Temp = replace(Temp, ",,,", ",")
    Temp = replace(Temp, ",,", ",")
    Temp = Trim(Temp)
    'This shouldn't happen, if it does the query uses special functions with improper syntax
        'temp = Replace(temp, "*", "")
        'If Left(temp, 1) = "," Then temp = Right(temp, Len(temp) - 1)
    RemoveSQLCommandsFromQuery = Temp
End Function

Function VerifyQuery(ByVal Query, Optional ByVal Fields, Optional ErrorMessage)
    'TODO: Handle other formulas besides IFF
    Temp = Query
    Temp = RemoveSQLCommandsFromQuery(Temp)
    Temp = ExtractSQLFormulaFieldsFromQuery(Temp, 1, 0, 1)
    QueryFieldsa = Split(Temp, ",")
    Fieldsa = Split(Fields, ",")
    For Each qf In QueryFieldsa
        qf = ExtractSQLFieldFromFormula(qf)
        qf = RemoveSQLOperators(qf)
        bMatchFound = False
        qft = SplitString(qf, ".", 0)
        qfval = SplitString(qf, ".", -1)
        For Each tf In Fieldsa
            tft = SplitString(tf, ".", 0)
            If (qf <> "" And tf <> "") Then
                If InString(qf, tf, 0) Or (qf = "*") Or (qfval = "*" And qft = tft) Then
                    bMatchFound = True
                    Exit For
                End If
            End If
        Next tf
        If bMatchFound = False Then
            ErrorMessage = "Error (Query Field Not Found): " & qf
            VerifyQuery = False
            'VerifyQuery = False & ErrorMessage
            Exit Function
        End If
    Next qf
    VerifyQuery = True
End Function

Function MakeOrStringFromCategories(field, category, cats, vals, Optional skip1 = 1, Optional thevartype = 2)

    'categoryvalue = category.Value
    If Not IsArray(cats) Or TypeName(cats) = "Range" Then
        catsa = RangetoArray(cats, thevartype, skip1, 1)
    Else
        catsa = cats
    End If
    If Not IsArray(vals) Or TypeName(vals) = "Range" Then
        valsa = RangetoArray(vals, thevartype, skip1, 1)
    Else
        valsa = cats
    End If
    tempstr = ""
    For i = 0 To UBound(valsa)
        If catsa(i) = category Then
            If tempstr <> "" Then
                tempstr = tempstr & " OR "
            End If
            If thevartype = 1 Then
                tempstr = tempstr & field & " = " & valsa(i)
            ElseIf thevartype = 2 Then
                tempstr = tempstr & field & " = '" & valsa(i) & "'"
            End If
        End If
    Next i
    
    MakeOrStringFromCategories = tempstr
    
End Function

Function IsArrayAllocated(arr As Variant) As Boolean
On Error Resume Next
    IsArrayAllocated = IsArray(arr) And _
    Not IsError(LBound(arr, 1)) And _
    LBound(arr, 1) <= UBound(arr, 1)
End Function

Public Function ConnectionTest(ConnectionString As String, Optional Query As String)

    ' Late-binding: requires less effort, but he correct aproach is
    ' to create a reference to 'Microsoft ActiveX Data Objects' -
    
    'Dim conADO As ADODB.Connection
    'Set conADO = New ADODB.Connection
    Dim conADO As Object
    Set conADO = CreateObject("ADODB.Connection")
    
    ConnectionString = SplitString(ConnectionString, "ODBC;", -1)
    
    Dim i As Integer
    
    conADO.ConnectionTimeout = 30
    conADO.ConnectionString = ConnectionString
    
    On Error Resume Next
    
    conADO.Open
    
    connresult = ""
    If conADO.State = 1 Then
        connresult = "Connection string is valid"
        If Query <> "" Then
            Set rs = conADO.Execute(Query)
            For i = 0 To conADO.Errors.count
                With conADO.Errors(i)
                    connresult = connresult & "Query error: " & .Number & " (native error '" & .NativeError & "') from '" & .Source & "': " & .Description & vbLf
                End With
            Next i
        End If
    Else
        connresult = "Connection failed:"
        For i = 0 To conADO.Errors.count
            With conADO.Errors(i)
                connresult = connresult & "ADODB connection returned error: " & .Number & " (native error '" & .NativeError & "') from '" & .Source & "': " & .Description & vbLf
            End With
        Next i
    End If
    
    'conADO.Close
    Set conADO = Nothing
    
    ConnectionTest = connresult

End Function

Function MyIsFormula(cell)
    MyIsFormula = cell.Formula <> cell And cell.Formula <> cell.Text
End Function

Function UnduplicateSQLFields(ByVal AvailableFieldsa As Variant, Optional Table)
    
    'Make sure data is in the right format
    If IsArray(AvailableFieldsa) Then
    Else
        If TypeName(AvailableFieldsa) = "Range" Then
            If AvailableFieldsa.count > 1 Then
                AvailableFieldsa = RangetoArray(AvailableFieldsa, vbString, 1, 1)
            End If
        End If
        If Not IsArray(AvailableFieldsa) Then
            Dim AvailableFields As String
            AvailableFields = replace(replace(Trim(AvailableFieldsa), vbLf, " "), "  ", " ")
            If Right(AvailableFields, 1) = "," Then AvailableFields = Left(AvailableFields, Len(AvailableFields) - 1)
            AvailableFieldsa = StringToArray(AvailableFields, ",", 1)
        End If
    End If
    
    'No default table selected, we should have a list of fields with the table names.
    If IsMissing(Table) Then
        'We will remove any plain field names that might be in the array.
        For i = 0 To UBound(AvailableFieldsa)
            HasTableName = InStr(AvailableFieldsa(i), ".")
            If HasTableName = 0 Then
                AvailableFieldsa(i) = ""
                Temp = 0
            End If
        Next i
    End If
    'Call RemoveEmptyArrayElements(AvailableFieldsa)
    'Remove empty fields and duplicates, keep first value. Function will evaluate duplicates without the table name.
    AvailableFieldsa = RemoveArrayDuplicates(AvailableFieldsa, 1, ".", -1)
    'Call QuickSort(AvailableFieldsa, 0, UBound(AvailableFieldsa))
    For j = 0 To UBound(AvailableFieldsa)
            If Not IsMissing(Table) Then
                HasTableName = InStr(AvailableFieldsa(j), ".")
                If HasTableName = 0 Then
                    NewFieldsString = NewFieldsString & Table & "." & AvailableFieldsa(j) & ","
                Else
                    NewFieldsString = NewFieldsString & AvailableFieldsa(j) & ","
                End If
            Else
                NewFieldsString = NewFieldsString & AvailableFieldsa(j) & ","
            End If
    Next j
    If Right(NewFieldsString, 1) = "," Then NewFieldsString = Left(NewFieldsString, Len(NewFieldsString) - 1)
    UnduplicateSQLFields = NewFieldsString
End Function

Function ReplaceF(str, find, rep, Optional start = 1, Optional count = -1, Optional compare = vbBinaryCompare)
    ReplaceF = replace(str, find, rep, start, count, compare)
End Function

Function QtoPY(myQ)
    tYear = Left(myQ, 4)
    quarter = Right(myQ, 1)
    If quarter = 4 Then
        tYear = tYear + 1
    End If
    QtoPY = tYear
End Function

Function HasMaxValue(id, val, idr, valr)
    On Error Resume Next
    'Temp = Split(cell, ";")
    'searchsheet = searchr.Parent.Name
    'ArraySearchGet = ""
    
    ismax = True
    
    For Each c In valr
        If (IsEmpty(c)) Then
            Exit For
        End If
        If c > val And Cells(c.row, idr.Column) = id Then
            ismax = False
            Exit For
        End If
    Next
    
    HasMaxValue = ismax

End Function

Public Function RemoveEmptyArrayElements(arr As Variant)
    Dim temparr() As Variant
    validCount = 0
    For i = 0 To UBound(arr)
        If arr(i) <> "" Then
            ReDim Preserve temparr(validCount)
            temparr(validCount) = arr(i)
            validCount = validCount + 1
        End If
    Next i
    RemoveEmptyArrayElements = temparr
End Function

Function ArrayToString(arr As Variant, Optional delim = ",", Optional removeBlanks = 0, Optional removeLastDelim = 1)
    ArrayToString = ""
    For i = LBound(arr) To UBound(arr)
        If removeBlanks And arr(i) = "" Then
        Else
            ArrayToString = ArrayToString & arr(i) & delim
        End If
    Next i
    If Right(ArrayToString, 1) = delim And removeLastDelim Then
        ArrayToString = Left(ArrayToString, Len(ArrayToString) - 1)
    End If
End Function

Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function

Function ReplaceNewLine(str)
    Temp = replace(str, vbCr, " ")
    Temp = replace(Temp, vbLf, " ")
    Temp = replace(Temp, vbCrLf, " ")
    ReplaceNewLine = Temp
End Function

Function SortArrayAtoZ(myArray As Variant)

    Dim i As Long
    Dim j As Long
    Dim Temp
    
    'Sort the Array A-Z
    For i = LBound(myArray) To UBound(myArray) - 1
        For j = i + 1 To UBound(myArray)
            If UCase(myArray(i)) > UCase(myArray(j)) Then
                Temp = myArray(j)
                myArray(j) = myArray(i)
                myArray(i) = Temp
            End If
        Next j
    Next i
    
    SortArrayAtoZ = myArray

End Function

'1) Get dictionary
'2) Parse/format dictionary
'3) Apply algorithm
'   a) Check length
'   b) Check first x letters for match
'   c) Watch number of differences
'   d) Determine similarity
'       -Check neighboring x characters
'       -Check consecutive letter matches
'       -Check same number of letters
'       -Check first, second, last, and second to last (if size > 3) letters match
Function MatchWithDictionary(ByVal str As String, Optional Dictionary, Optional DictionaryColName, Optional OverallMatchThresh, Optional MatchType, Optional ReplaceWith, Optional UseNullReplace, Optional CheckFlags, Optional FlagAction, Optional LettersOnly, Optional UseBest, Optional ReturnType, Optional DictionaryHasHeaders)
    'Backup str
    strbak = str
    
    'Exit if string to check is blank
    If str = "" Or str = " " Then
        MatchWithDictionary = 0
        Exit Function
    End If
    
    'set default values
    If IsMissing(OverallMatchThresh) Or IsEmpty(OverallMatchThresh) Then OverallMatchThresh = 0.86
    'MatchType 1 = Letters; 2 = Words; 3 = Match all words in dictionary and replace; 4 = Match all words in dictionary and remove
    If IsMissing(MatchType) Or IsEmpty(MatchType) Then MatchType = 1
    'ReturnType 1 = Processed result; 2 = Index;
    If IsMissing(ReturnType) Or IsEmpty(ReturnType) Then ReturnType = 1
    If IsMissing(MaxNumResults) Or IsEmpty(MaxNumResults) Then MaxNumResults = 1
    If IsMissing(UseBest) Or IsEmpty(UseBest) Then UseBest = False
    If IsMissing(LettersOnly) Or IsEmpty(LettersOnly) Then LettersOnly = True
    If IsMissing(LengthDiff) Or IsEmpty(LengthDiff) Then LengthDiff = 2
    If IsMissing(InitalCharsDiff) Or IsEmpty(InitalCharsDiff) Then InitalCharsDiff = 2
    If IsMissing(LettersPosDiff) Or IsEmpty(LettersPosDiff) Then LettersPosDiff = 1
    If IsMissing(MaxDiffs) Or IsEmpty(MaxDiffs) Then MaxDiffs = 3
    If MatchType = 1 Then
        If IsMissing(MinWordSize) Or IsEmpty(MinWordSize) Then MinWordSize = 3
    Else
        If IsMissing(MinWordSize) Or IsEmpty(MinWordSize) Then MinWordSize = 1
    End If
    If IsMissing(UseNullReplace) Or IsEmpty(UseNullReplace) Then UseNullReplace = True
    'What to do when word is flagged 0 = include only non-flagged/exclude all flagged 1 = include only flagged/exclude non-flagged
    If IsMissing(FlagAction) Or IsEmpty(FlagAction) Then FlagAction = 0
    If IsMissing(DictionaryHasHeaders) Or IsEmpty(DictionaryHasHeaders) Then DictionaryHasHeaders = 1
    Call GetDictionary(Dictionary, DictionaryColName, ReplaceWith, CheckFlags)
    'If IsMissing(Dictionary) Then ReplaceWith = "Z:\Work\VBA\Special\SpellCheck\dictionary.txt"
    'If IsMissing(ReplaceWith) Then ReplaceWith = "Z:\Work\VBA\Special\SpellCheck\replace_words.txt"
    'If IsMissing(CheckFlags) Then CheckFlags = "Z:\Work\VBA\Special\SpellCheck\flags_abbreviation.txt"
    
    str = Trim(UCase(str))
    If LettersOnly Then
        For k = 1 To Len(str)
            cs = Mid(str, k, 1)
            If Not (cs Like "[A-Z]" Or (MatchType > 1 And cs = " ")) Then
                str = replace(str, cs, "")
                k = k - 1
            End If
            If k >= Len(str) Then Exit For
        Next k
    End If
    
    'first pass, check for exact match
    If MatchType < 3 Then
        For i = 1 + DictionaryHasHeaders To UBound(Dictionary)
            word = Dictionary(i)
            word = replace(Trim(UCase(word)), vbCr, "")
            If word = "HVAC" Or word = "CONSULTING SOLUTIONS" Then
                Temp = 0
            End If
            If str = word Then GoTo SkipMatch
        Next i
    End If
    
    Dim stra() As Variant
    If MatchType = 1 Then
        ReDim stra(1 To Len(str))
        For l = 1 To Len(str)
            stra(l) = Mid(str, l, 1)
        Next l
    ElseIf MatchType > 1 Then
        tempstra = Split(str, " ")
        ReDim stra(1 To UBound(tempstra) + 1)
        For w = 0 To UBound(tempstra)
            stra(w + 1) = tempstra(w)
        Next w
    End If
    
    If UBound(stra) < MinWordSize Or str = " " Or str = "" Then
        If ReturnType = 1 Then
            MatchWithDictionary = str
        ElseIf ReturnType = 2 Then
            MatchWithDictionary = 0
        End If
        Exit Function
    End If
    
    bestMatchesRatio = 0
    bestconsecMatchesRatio = 0
    bestLetterMatchesRatio = 0
    bestOverallMatchesRatio = 0
    NumResults = 0
    den = UBound(stra)
    
    For i = 1 + DictionaryHasHeaders To UBound(Dictionary)
    'For Each word In Dictionary
        word = Dictionary(i)
        word = replace(Trim(UCase(word)), vbCr, "")
        
        If word = "PHARMACY" Then
            Temp = 0
        End If
        
        'Before we do anything, if MatchType = 3 or 4 and no flag, check if we should ignore this phrase
        If MatchType > 2 And Not IsMissing(CheckFlags) Then
            If (CheckFlags(i) <> 1 And CheckFlags(i) <> "1" And FlagAction = 1) Or ((CheckFlags(i) = 1 Or CheckFlags(i) = "1") And FlagAction = 0) Then
                GoTo SkipWord
            End If
        End If
        'If MatchType = 3 or 4 and no flag, we can just look for the phrase directly
        If MatchType = 3 Or MatchType = 4 Then
            tempstr = " " & str & " "
            tempword = " " & word & " "
            If InStr(1, tempstr, tempword, vbTextCompare) Then
                If MatchType = 3 Then
                    GoTo SkipMatch
                ElseIf MatchType = 4 Then
                    str = replace(str, word, " ")
                    str = RemoveExtraSpaces(str)
                End If
            End If
            GoTo SkipWord
        End If
        
        If word = str Then GoTo SkipMatch
        If word = "" Or word = " " Then GoTo SkipWord
        
        Dim worda() As Variant
        If MatchType = 1 Then
            ReDim worda(1 To Len(word))
            For l = 1 To Len(word)
                worda(l) = Mid(word, l, 1)
            Next l
        ElseIf MatchType > 1 Then
            tempworda = Split(word, " ")
            ReDim worda(1 To UBound(tempworda) + 1)
            For w = 0 To UBound(tempworda)
                worda(w + 1) = tempworda(w)
            Next w
        End If
        
        If UBound(worda) < MinWordSize Then
            GoTo SkipWord
        End If
        
        'Check if words appear to be vastly different. First, check if the first few or last items match. Next, check if lengths are substantially different. Before giving up, check that the words aren't really similar.
        If UBound(stra) > 2 And UBound(worda) > 2 Then
            If Not (worda(1) = stra(1) Or worda(1) = stra(2) Or worda(2) = stra(1) Or worda(2) = stra(2) Or worda(UBound(worda)) = stra(UBound(stra))) Then GoTo SkipWord
        End If
        If Abs(UBound(worda) - UBound(stra)) > LengthDiff Or Abs(UBound(worda) - UBound(stra)) >= UBound(worda) Or Abs(UBound(worda) - UBound(stra)) >= UBound(stra) Or word = "" Or word = " " Then
            If UBound(stra) > 3 And UBound(worda) > 3 Then
                If Not worda(1) = stra(1) And worda(2) = stra(2) And worda(3) = stra(3) And worda(4) = stra(4) Then GoTo SkipWord
            Else
                GoTo SkipWord
            End If
        End If
        
        matches = 0
        consecMatches = 0
        LetterMatches = 0
        PositionMatches = 0
        For j = 1 To UBound(worda)
            MatchFound = False
            c = worda(j)
            If InStr(1, str, c) Then LetterMatches = LetterMatches + 2
            If j <= UBound(stra) Then
                k = j
                cs = stra(k)
            Else
                cs = ""
            End If
            If c = cs Then
                MatchFound = True
            Else
                'If MatchType > 2 Then
                '    MinPos = LBound(stra)
                '    MaxPos = UBound(stra)
                'Else
                    If j - LettersPosDiff < 1 Then MinPos = 1 Else MinPos = j - LettersPosDiff
                    If j + LettersPosDiff > UBound(stra) Then MaxPos = UBound(stra) Else MaxPos = j + LettersPosDiff
                'End If
                For k = MaxPos To MinPos Step -1
                    If k <> j Then
                        cs = stra(k)
                        If c = cs Then
                            MatchFound = True
                            Exit For
                        End If
                    End If
                Next k
            End If
            If MatchFound Then
                matches = matches + 1
                If (j = 1 And k = 1) Or (UBound(worda) = j And UBound(stra) = k) Or (j - 1 = lastj And k - 1 = lastk) Then consecMatches = consecMatches + 1
                If (j = 1 And k = 1) Or (j = 2 And k = 2) Or (j = UBound(worda) And k = UBound(stra)) Then PositionMatches = PositionMatches + 1
                If UBound(worda) > 3 And UBound(stra) > 3 And j = UBound(worda) - 1 And k = UBound(stra) - 1 Then PositionMatches = PositionMatches + 1
                lastj = j
                lastk = k
                'If MatchType = 3 And LetterMatches / 2 / UBound(worda) >= 1 Then GoTo SkipMatch
                'If MatchType = 4 Then tempstr = Mid(tempstr, 1, k - 1) & " " & Mid(tempstr, k + 1, UBound(stra))
                GoTo SkipC
            End If
            If (j > InitalCharsDiff And matches = 0) Or (j - matches > MaxDiffs) Then GoTo SkipWord
SkipC:
        Next j
        MatchesRatio = matches / den
        consecMatchesRatio = consecMatches / UBound(worda)
        LetterMatchesRatio = Application.WorksheetFunction.Min(1, LetterMatches / (UBound(stra) + UBound(worda)))
        PositionMatchesRatio = PositionMatches / Application.WorksheetFunction.Min(4, UBound(worda))
        OverallMatchesRatio = (MatchesRatio + consecMatchesRatio + LetterMatchesRatio + PositionMatchesRatio) / 4
        'If (MatchesRatio >= bestMatchesRatio And consecMatchesRatio >= ConsecMatchThresh And LetterMatchesRatio >= bestLetterMatchesRatio) Or OverallMatchesRatio > bestOverallMatchesRatio Then
        If OverallMatchesRatio > bestOverallMatchesRatio Or OverallMatchesRatio = 1 Then
            bestMatch = word
            bestMatchIndex = i
            bestMatchesRatio = MatchesRatio
            bestconsecMatchesRatio = MatchesRatio
            bestLetterMatchesRatio = LetterMatchesRatio
            bestPositionMatchesRatio = PositionMatchesRatio
            bestOverallMatchesRatio = OverallMatchesRatio
            'If (MatchesRatio >= MatchThresh And consecMatchesRatio >= ConsecMatchThresh And LetterMatchesRatio >= LetterMatchThresh) Or OverallMatchesRatio > OverallMatchThresh Then
            If OverallMatchesRatio > OverallMatchThresh Or OverallMatchesRatio = 1 Then
                NumResults = NumResults + 1
                If NumResults >= MaxNumResults Then
SkipMatch:
                    If IsMissing(CheckFlags) Then
                        FlagOkay = True
                    ElseIf (CheckFlags(i) <> 1 And FlagAction = 0) Or (CheckFlags(i) = 1 And FlagAction = 1) Then
                        FlagOkay = True
                    Else
                        FlagOkay = False
                    End If
                    If IsMissing(ReplaceWith) Then
                        ReplaceOkay = False
                    ElseIf UseNullReplace Or ReplaceWith(i) <> "" Then
                        ReplaceOkay = True
                    Else
                        ReplaceOkay = False
                    End If
                    If FlagOkay Then
                        If ReplaceOkay And ReturnType = 1 Then
                            MatchWithDictionary = ReplaceWith(i)
                        ElseIf ReturnType = 2 Then
                            MatchWithDictionary = i
                        Else
                            MatchWithDictionary = word
                        End If
                    Else
                        If UseNullReplace And ReturnType = 1 Then
                            MatchWithDictionary = ""
                        ElseIf ReturnType = 2 Then
                            MatchWithDictionary = 0
                        Else
                            MatchWithDictionary = str
                        End If
                    End If

                    Exit Function
                End If
            End If
        End If
SkipWord:
    Next i
    
    If UseBest And bestOverallMatchesRatio >= OverallMatchThresh / 2 Then
        If ReturnType = 1 Then
            MatchWithDictionary = bestMatch
        ElseIf ReturnType = 2 Then
            MatchWithDictionary = bestMatchIndex
        End If
    ElseIf IsEmpty(MatchWithDictionary) Then
        If ReturnType = 1 Then
            If str = "" Then
                MatchWithDictionary = strbak
            Else
                MatchWithDictionary = str
            End If
        ElseIf ReturnType = 2 Then
            MatchWithDictionary = 0
        End If
    End If
    
End Function

Function CellMatchWithDictionary(ByVal cc, Optional Dictionary, Optional DictionaryColName, Optional OverallMatchThresh, Optional MatchType, Optional ReplaceWith, Optional UseNullReplace, Optional CheckFlags, Optional FlagAction, Optional LettersOnly, Optional UseBest, Optional ReturnType, Optional DictionaryHasHeaders)

    'If IsMissing(Dictionary) Then Dictionary = "Z:\Work\VBA\Special\SpellCheck\dictionary.txt"
    'MatchType 1 = Letters; 2 = Words;
    If IsMissing(MatchType) Then MatchType = 1
    If IsMissing(LettersOnly) Then LettersOnly = True
    
    Call GetDictionary(Dictionary, DictionaryColName, ReplaceWith, CheckFlags)
    
    cctype = TypeName(cc)
    
    If cctype = "Range" Then
        ccstr = Trim(UCase(cc.Value2))
    Else
        ccstr = Trim(UCase(cc))
    End If
    If LettersOnly Then
        For k = 1 To Len(ccstr)
            cs = Mid(ccstr, k, 1)
            If Not (cs Like "[A-Z]") And cs <> " " Then
                ccstr = replace(ccstr, cs, " ")
                k = k - 1
            End If
            If k >= Len(ccstr) Then Exit For
        Next k
    End If
    ccstr = replace(replace(ccstr, vbCr, " "), vbLf, " ")
    ccstr = replace(ccstr, "    ", " ")
    ccstr = replace(ccstr, "   ", " ")
    ccstr = replace(ccstr, "  ", " ")
    ccstr = Trim(ccstr)
    
    Dim tempstr As String
    If MatchType = 1 Then
        ccA = Split(ccstr, " ")
        tempstr = ""
        For Each ccword In ccA
            Trim (ccword)
            If ccword <> "" And ccword <> " " Then
                tempstr = tempstr & MatchWithDictionary(ccword, Dictionary, DictionaryColName, OverallMatchThresh, MatchType, ReplaceWith, UseNullReplace, CheckFlags, FlagAction, LettersOnly, UseBest, ReturnType, DictionaryHasHeaders) & " "
            End If
        Next ccword
    ElseIf MatchType > 1 Then
        tempstr = tempstr & MatchWithDictionary(ccstr, Dictionary, DictionaryColName, OverallMatchThresh, MatchType, ReplaceWith, UseNullReplace, CheckFlags, FlagAction, LettersOnly, UseBest, ReturnType, DictionaryHasHeaders)
    End If
    CellMatchWithDictionary = RemoveExtraSpacesFormulas(tempstr)
    
End Function

Function RangeMatchWithDictionary(ByVal cc, Optional Dictionary, Optional DictionaryColName, Optional OverallMatchThresh, Optional MatchType, Optional ReplaceWith, Optional UseNullReplace, Optional CheckFlags, Optional FlagAction, Optional LettersOnly, Optional UseBest, Optional ReturnType, Optional DictionaryHasHeaders, Optional SkipRows = 1)
On Error GoTo ErrHandle

    'MatchType 1 = Letters; 2 = Words;
    If IsMissing(MatchType) Then MatchType = 1
    If IsMissing(LettersOnly) Then LettersOnly = True
    'If IsMissing(Dictionary) Then Dictionary = "Z:\Work\VBA\Special\SpellCheck\dictionary.txt"
    Call GetDictionary(Dictionary, DictionaryColName, ReplaceWith, CheckFlags)
    
    'Try different approaches to get range size, check row size and count a
    'limit to 10000 rows for now
    maxrangesize = 10000
    If TypeName(cc) = "Range" Then
        rangesizea = Application.WorksheetFunction.CountA(cc)
        If rangesizea > 1 Then
            rangesizeb = cc.Cells(Rows.count, 1).End(xlUp).row
            rangesizec = cc.Cells.SpecialCells(xlCellTypeLastCell).row
        End If
        'Use largest size
        rangesize = Application.WorksheetFunction.Max(rangesizea, rangesizeb, rangesizec)
        rangesize = Application.WorksheetFunction.Min(rangesize, maxrangesize)
        cc = RangetoArray(cc, 2, 0, 1, 1, 1, rangesize)
    Else
        If NumberOfArrayDimensions(cc) > 1 Then
            cc = Flatten2DArray(cc, maxrangesize, 1)
        End If
        rangesize = UBound(cc)
        rangesize = Application.WorksheetFunction.Min(rangesize, maxrangesize)
    End If
    
    If rangesize = 1 Then SkipRows = 0
    Dim MatchedValuesa()
    ReDim MatchedValuesa(rangesize - SkipRows)
    For i = SkipRows To rangesize - 1
        'Temp = cc.Cells(i, 1)
        cv = cc(i)
        MatchedValuesa(i - SkipRows) = CellMatchWithDictionary(cv, Dictionary, DictionaryColName, OverallMatchThresh, MatchType, ReplaceWith, UseNullReplace, CheckFlags, FlagAction, LettersOnly, UseBest, ReturnType, DictionaryHasHeaders)
    Next i
    
    RangeMatchWithDictionary = MatchedValuesa
    Exit Function
    
ErrHandle:
    Debug.Print "Error (RangeMatchWithDictionary): " & Err.Description
    Temp = 0
    
End Function

Function SplitQ(str, delim, qualstart, qualend, Optional tempdelim = ";")
    dontsplit = False
    tempstr = ""
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If c = qualstart Then dontsplit = True
        If c = qualend Then dontsplit = False
        If dontsplit And c = delim Then
            tempstr = tempstr & tempdelim
        Else
            tempstr = tempstr & c
        End If
    Next i
    stra = Split(tempstr, delim)
    For j = 0 To UBound(stra)
        stra(j) = replace(stra(j), tempdelim, delim)
    Next j
    SplitQ = stra
End Function

Sub GetDictionary(Optional Dictionary, Optional DictionaryColName, Optional ReplaceWith, Optional CheckFlags)

    If Not IsArray(Dictionary) And Not TypeName(Dictionary) = "Range" Then
        If IsMissing(Dictionary) Then
            DictionaryRow = Application.match("Dictionary", ThisWorkbook.Worksheets("Settings").Range("A:A"), 0)
            DictionarySource = ThisWorkbook.Sheets("Settings").Cells(DictionaryRow, 2)
        ElseIf Not FileExists(Dictionary) Then
            DictionaryRow = Application.match(Dictionary, ThisWorkbook.Worksheets("Settings").Range("A:A"), 0)
            DictionarySource = ThisWorkbook.Sheets("Settings").Cells(DictionaryRow, 2)
        End If
    End If
    
    'Get Dictionary
    If Not IsEmpty(DictionarySource) Then
        DictionaryData = TextFileToArray(DictionarySource, vbLf, ",")
        If IsMissing(DictionaryColName) Then
            ddcolname = "raw_word"
        Else
            ddcolname = DictionaryColName
        End If
        For i = 0 To UBound(DictionaryData, 2)
            If DictionaryData(0, i) = ddcolname Then
                Dictionary = Application.Transpose(Application.index(DictionaryData, 0, i + 1))
                Exit For
            End If
        Next i
    ElseIf Not IsArray(Dictionary) And Not TypeName(Dictionary) = "Range" Then
        Dictionary = TextFileToArray(Dictionary)
    ElseIf TypeName(Dictionary) = "Range" Then
        Dictionary = RangetoArray(Dictionary, 2, 0, 1, 1, 1, , 1)
    End If
    
    'Get Replace with
    If Not IsMissing(ReplaceWith) And Not IsArray(ReplaceWith) Then
        If Not IsMissing(DictionaryData) Then
            ddcolname = ReplaceWith
            For i = 0 To UBound(DictionaryData, 2)
                If DictionaryData(0, i) = ddcolname Then
                    ReplaceWith = Application.Transpose(Application.index(DictionaryData, 0, i + 1))
                    Exit For
                End If
            Next i
        ElseIf Not IsArray(ReplaceWith) And Not TypeName(ReplaceWith) = "Range" Then
            ReplaceWith = TextFileToArray(ReplaceWith)
        ElseIf TypeName(ReplaceWith) = "Range" Then
            ReplaceWith = RangetoArray(ReplaceWith, 2, 0, 1, 1, 1, , 1)
        End If
    End If
    
    'Get Flags
    If Not IsMissing(CheckFlags) And Not IsArray(CheckFlags) Then
        If Not IsMissing(DictionaryData) Then
            ddcolname = CheckFlags
            For i = 0 To UBound(DictionaryData, 2)
                If DictionaryData(0, i) = ddcolname Then
                    CheckFlags = Application.Transpose(Application.index(DictionaryData, 0, i + 1))
                    Exit For
                End If
            Next i
        ElseIf Not IsArray(CheckFlags) And Not TypeName(CheckFlags) = "Range" Then
            CheckFlags = TextFileToArray(CheckFlags)
        ElseIf TypeName(CheckFlags) = "Range" Then
            CheckFlags = RangetoArray(CheckFlags, 2, 0, 1, 1, 1, , 1)
        End If
    End If
    
End Sub

' return type: 1 = workbook; 2 = workbookname
Function FindSheetWorkbook(sheetname, Optional ReturnType = 1)
    For Each wb In Workbooks
        If WorksheetExists(sheetname, wb) Then
            wbn = wb.Name
            Exit For
        End If
    Next wb
    If ReturnType = 1 Then
        Set FindSheetWorkbook = wb
    ElseIf ReturnType = 2 Then
        FindSheetWorkbook = wbn
    End If
End Function

