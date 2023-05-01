Attribute VB_Name = "Ship"
Option Explicit

Sub Data()
    
    
    Dim path As Variant
    path = sf.getPathArray("B3")
    
    Dim txtBuffer As String
    Dim txtLine() As String
    
    
    Dim startDataCell As Object
    Set startDataCell = ThisWorkbook.Worksheets("Ships").range("B3")
    
    Dim heading As Variant
    heading = sf.getHeadingArray("Ships", "C2")
    
    Dim offsetX As Integer
    Dim offsetY As Integer
    offsetY = 0
    
    
    Dim tmpString As String
    Dim tmpArray() As String
    Dim lineLength As Long
    Dim shipNameExtract() As String
    Dim stringPosition As Integer
    
    
    Dim gunCount As Integer
    Dim turretCount As Integer
    Dim fighterBay As Integer
    Dim droneBay As Integer
    
    
    Dim haveAttributes As Integer
    Dim foundAttributes As Boolean
    
    
    Dim lineNum As Integer
    Dim pathNum As Integer
    
    
    For pathNum = LBound(path) To UBound(path)
        
        Open path(pathNum) For Input As #1
        txtBuffer = Input(LOF(1), #1)
        txtLine = Split(txtBuffer, vbLf)
        
        haveAttributes = 0
        foundAttributes = False
        
        
        For lineNum = LBound(txtLine) To UBound(txtLine)
            
            
            
            If InStr(1, txtLine(lineNum), "ship") = 1 Then
                
                
                ThisWorkbook.Worksheets("Ships").range(startDataCell.Offset(offsetY, 0).Address & ":" & startDataCell.Offset(offsetY, UBound(heading)).Address).ClearContents
                
                shipNameExtract = sf.cleanStrArray(Split(txtLine(lineNum), Chr(34)))
                startDataCell.Offset(offsetY, 0).Value = shipNameExtract(UBound(shipNameExtract))
                
                
                gunCount = 0
                turretCount = 0
                fighterBay = 0
                droneBay = 0
                
                
                lineNum = lineNum + 1
                While InStr(1, txtLine(lineNum), "ship") <> 1 And lineNum < UBound(txtLine)
                    
                    If haveAttributes = 1 Then
                        
                        
                        If InStr(1, txtLine(lineNum), "licenses") <> 0 Then
                            
                            lineNum = lineNum + 1
                            
                            offsetX = sf.inStringArray(heading, "licenses")
                            
                            If offsetX > -1 Then
                                
                                tmpString = txtLine(lineNum)
                                
                                tmpString = Replace(tmpString, vbTab, "")
                                tmpString = Replace(tmpString, Chr(34), "")
                                tmpString = Trim(tmpString)
                                
                                startDataCell.Offset(offsetY, offsetX).Value = tmpString
                            End If
                            
                        ElseIf InStr(2, txtLine(lineNum), vbTab) = 2 And (InStr(1, txtLine(lineNum), " ") <> 0 Or InStr(1, txtLine(lineNum), Chr(34)) <> 0) Then
                            
                            If InStr(1, txtLine(lineNum), Chr(34)) <> 0 Then
                                tmpArray = sf.cleanStrArray(Split(txtLine(lineNum), Chr(34)))
                            Else
                                tmpArray = sf.cleanStrArray(Split(txtLine(lineNum), " "))
                            End If
                            
                            
                            offsetX = sf.inStringArray(heading, tmpArray(LBound(tmpArray)))
                            If offsetX > -1 Then
                                startDataCell.Offset(offsetY, offsetX).Value = tmpArray(UBound(tmpArray))
                            End If
                            
                        Else
                            haveAttributes = 0
                        End If
                        
                        
                    ElseIf haveAttributes = 2 Then
                        
                        
                        If InStr(2, txtLine(lineNum), vbTab) = 2 And (InStr(1, txtLine(lineNum), " ") <> 0 Or InStr(1, txtLine(lineNum), Chr(34)) <> 0) Then
                            
                            If InStr(1, txtLine(lineNum), Chr(34)) <> 0 Then
                                tmpArray = sf.cleanStrArray(Split(txtLine(lineNum), Chr(34)))
                            Else
                                tmpString = txtLine(lineNum)
                                tmpString = Replace(tmpString, Chr(34), " ")
                                tmpArray = sf.cleanStrArray(Split(tmpString, " "))
                            End If
                            
                            offsetX = sf.inStringArray(heading, tmpArray(LBound(tmpArray)))
                            
                            If offsetX > -1 Then
                                startDataCell.Offset(offsetY, offsetX).Value = CLng(startDataCell.Offset(offsetY, offsetX).Value) + CLng(tmpArray(UBound(tmpArray)))
                            End If
                            
                        Else
                            haveAttributes = 0
                        End If
                        
                        
                    Else
                        
                        
                        stringPosition = InStr(1, txtLine(lineNum), "attributes")
                        
                        If stringPosition = 2 Then              'attributes
                            haveAttributes = 1
                            foundAttributes = True
                        
                        ElseIf stringPosition = 6 Then          'add attributes
                            haveAttributes = 2
                            foundAttributes = True
                    
                            Dim stockShipCell As Object
                            Set stockShipCell = ThisWorkbook.Worksheets("Ships").range(searchContentUpward(startDataCell.Offset(offsetY, 0).Address, shipNameExtract(1), 5))         'find the cell address of stock(original) ship
                    
                            ThisWorkbook.Worksheets("Ships").range(stockShipCell.Offset(0, 1).Address & ":" & stockShipCell.Offset(0, UBound(heading)).Address).Copy _
                                (ThisWorkbook.Worksheets("Ships").range(startDataCell.Offset(offsetY, 1).Address & ":" & startDataCell.Offset(offsetY, UBound(heading)).Address))    'copy stats written for the stock ship
                        End If
                
                
                    End If
            
            
                    If InStr(1, txtLine(lineNum), "gun") = 2 Then
                        gunCount = gunCount + 1
                    ElseIf InStr(1, txtLine(lineNum), "turret") = 2 Then
                        turretCount = turretCount + 1
                    ElseIf InStr(1, txtLine(lineNum), "bay ""Fighter") = 2 Then
                        fighterBay = fighterBay + 1
                    ElseIf InStr(1, txtLine(lineNum), "bay ""Drone") = 2 Then
                        droneBay = droneBay + 1
                    End If
            
            
                    lineNum = lineNum + 1
                Wend
                lineNum = lineNum - 1
        
        
                offsetX = sf.inStringArray(heading, "gun")
                If offsetX > -1 Then
                    startDataCell.Offset(offsetY, offsetX).Value = gunCount
                End If
        
                offsetX = sf.inStringArray(heading, "turret")
                If offsetX > -1 Then
                    startDataCell.Offset(offsetY, offsetX).Value = turretCount
                End If
        
                offsetX = sf.inStringArray(heading, "fighter bay")
                If offsetX > -1 Then
                    startDataCell.Offset(offsetY, offsetX).Value = fighterBay
                End If
        
                offsetX = sf.inStringArray(heading, "drone bay")
                If offsetX > -1 Then
                    startDataCell.Offset(offsetY, offsetX).Value = droneBay
                End If
        
        
                If foundAttributes = True Then
                    offsetY = offsetY + 1
                    foundAttributes = False
                End If
        
        
            End If
    
    
        Next lineNum


        Close #1
    Next pathNum



    MsgBox ("Ship data imported")
End Sub



Function searchContentUpward(startCellCoor As String, targetContent As String, distance As Integer)
    
    Dim startCell As Object
    Set startCell = ThisWorkbook.Worksheets("Ships").range(startCellCoor)
    
    Dim i As Integer
    For i = 1 To distance
        If StrComp(targetContent, startCell.Offset(-i, 0).Value) = 0 Then
            searchContentUpward = startCell.Offset(-i, 0).Address
            Exit Function
        End If
    Next
    
    MsgBox ("stock ship not found")
    
End Function


