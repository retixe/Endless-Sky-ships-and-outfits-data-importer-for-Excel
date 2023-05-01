Attribute VB_Name = "sf"

'========== Shared Functions =========='

Option Explicit

Function cleanStrArray(ByVal strArray As Variant)
    
    If UBound(strArray) <= 0 Then
        MsgBox ("cleanStrArray: Empty array")
        Exit Function
    End If
    
    Dim buffer() As String
    ReDim buffer(UBound(strArray))
    
    Dim bufferIndex As Integer
    bufferIndex = 0
    
    Dim i As Integer
    For i = LBound(strArray) To UBound(strArray)
        
        strArray(i) = Replace(strArray(i), vbTab, "")
        strArray(i) = Trim(strArray(i))
        
        If StrComp(strArray(i), "") <> 0 Then
            buffer(bufferIndex) = strArray(i)
            bufferIndex = bufferIndex + 1
        End If
        
    Next
    
    ReDim Preserve buffer(bufferIndex - 1)
    
    cleanStrArray = buffer
    
End Function

Function inStringArray(ByRef inputStrArray As Variant, ByRef targetValue As Variant)
    
    Dim i As Integer
    For i = LBound(inputStrArray) To UBound(inputStrArray)
        
        If StrComp(inputStrArray(i), targetValue, vbTextCompare) = 0 Then
            
            inStringArray = i
            Exit Function
            
        End If
    Next
    
    inStringArray = -1
    
End Function

Function getPathArray(rangeStart As String)
    
    Dim startPathCell As range
    Set startPathCell = ThisWorkbook.Worksheets("filepath").range(rangeStart)
    
    
    Dim offsetY As Integer
    offsetY = 0
    While Trim(startPathCell.Offset(offsetY, 0).Value) <> ""
        offsetY = offsetY + 1
    Wend
    
    If offsetY = 0 Then
        MsgBox ("no stats To search")
        Exit Function
    End If
    
    Dim rangeEnd As String
    rangeEnd = startPathCell.Offset(offsetY - 1, 0).Address
    
    getPathArray = Application.Transpose(Application.Index(ThisWorkbook.Worksheets("filepath").range(rangeStart & ":" & rangeEnd).Value, 0, 1))
    
End Function

Function getHeadingArray(ByRef worksheetName As String, ByRef rangeStart As String)
    
    Dim startHeadingCell As range
    Set startHeadingCell = ThisWorkbook.Worksheets(worksheetName).range(rangeStart)
    
    
    Dim offsetX As Integer
    offsetX = 0
    While Trim(startHeadingCell.Offset(0, offsetX).Value) <> ""
        offsetX = offsetX + 1
    Wend
    
    If offsetX = 0 Then
        MsgBox ("no stats To search")
        Exit Function
    End If
    
    Dim rangeEnd As String
    rangeEnd = startHeadingCell.Offset(0, offsetX - 1).Address
    
    getHeadingArray = Application.Index(ThisWorkbook.Worksheets(worksheetName).range(rangeStart & ":" & rangeEnd).Value, 1, 0)
    
End Function

Function getSubmunitionData(ByRef txtLine As Variant, ByRef submunitionName As String, ByRef submunitionCount As Integer, _
                            ByRef totalShieldDmg As Single, ByRef totalHullDmg As Single, ByRef lifetime As Integer, ByRef range As Integer)
    
    Dim lineNum As Integer
    Dim outfitNameExtract() As String
    Dim tmpArray() As String
    
    Dim velocity As Integer
    Dim lifetimeTemp As Integer
    
    
    For lineNum = LBound(txtLine) To UBound(txtLine)
        
        If InStr(1, txtLine(lineNum), "outfit") = 1 Then
            
            If InStr(1, txtLine(lineNum), Chr(96)) <> 0 Then
                outfitNameExtract = sf.cleanStrArray(Split(txtLine(lineNum), Chr(96)))
            ElseIf InStr(1, txtLine(lineNum), Chr(34)) <> 0 Then
                outfitNameExtract = sf.cleanStrArray(Split(txtLine(lineNum), Chr(34)))
            Else
                outfitNameExtract = sf.cleanStrArray(Split(txtLine(lineNum), " "))
            End If
            
            If outfitNameExtract(UBound(outfitNameExtract)) = submunitionName Then
                
                velocity = 0
                lifetimeTemp = 0
                
                lineNum = lineNum + 1
                While InStr(1, txtLine(lineNum), vbTab) = 1
                    
                    If (InStr(1, txtLine(lineNum), " ") <> 0 Or InStr(1, txtLine(lineNum), Chr(34)) <> 0) Then
                        
                        If InStr(1, txtLine(lineNum), Chr(34)) <> 0 Then
                            tmpArray = sf.cleanStrArray(Split(txtLine(lineNum), Chr(34)))
                        Else
                            tmpArray = sf.cleanStrArray(Split(txtLine(lineNum), " "))
                        End If
                        
                        
                        Select Case tmpArray(LBound(tmpArray))
                            Case "velocity"
                                velocity = CInt(tmpArray(UBound(tmpArray)))
                                
                            Case "lifetime"
                                lifetimeTemp = CInt(tmpArray(UBound(tmpArray)))
                                
                            Case "shield damage"
                                totalShieldDmg = totalShieldDmg + (CSng(tmpArray(UBound(tmpArray))) * submunitionCount)
                                
                            Case "hull damage"
                                totalHullDmg = totalHullDmg + (CSng(tmpArray(UBound(tmpArray))) * submunitionCount)
                                
                        End Select
                        
                    End If
                    lineNum = lineNum + 1
                    
                Wend
                
                If (velocity = 0 Or lifetimeTemp = 0) Then
                    lifetime = lifetime + lifetimeTemp
                Else
                    range = range + (lifetimeTemp * velocity)
                End If
                
                Exit Function
                
            End If
        End If
    Next lineNum
    
    MsgBox (submunitionName & "not found")
End Function


