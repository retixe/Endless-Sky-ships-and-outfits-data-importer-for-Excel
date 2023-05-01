Attribute VB_Name = "HandToHandOutfit"
Option Explicit

Sub Data()
    
    
    Dim path As Variant
    path = sf.getPathArray("D3")
    
    Dim txtBuffer As String
    Dim txtLine() As String
    
    
    Dim startDataCell As range
    Set startDataCell = ThisWorkbook.Worksheets("Hand to Hand").range("B3")
    
    Dim heading As Variant
    heading = sf.getHeadingArray("Hand to Hand", "C2")
    
    Dim offsetX As Integer
    Dim offsetY As Integer
    offsetY = 0
    
    
    Dim tmpString As String
    Dim tmpArray() As String
    Dim lineLength As Long
    Dim outfitNameExtract() As String
    Dim stringPosition As Integer
    
    
    Dim haveCategory As Integer
    Dim foundCategory As Boolean
    
    
    Dim lineNum As Integer
    Dim pathNum As Integer
    
    
    For pathNum = LBound(path) To UBound(path)
        
        Open path(pathNum) For Input As #1
        txtBuffer = Input(LOF(1), #1)
        txtLine = Split(txtBuffer, vbLf)
        
        haveCategory = 0
        foundCategory = False
        
        
        For lineNum = LBound(txtLine) To UBound(txtLine)
            
            
            
            If InStr(1, txtLine(lineNum), "outfit") = 1 Then
                
                
                ThisWorkbook.Worksheets("Hand to Hand").range(startDataCell.Offset(offsetY, 0).Address & ":" & startDataCell.Offset(offsetY, UBound(heading)).Address).ClearContents
                
                
                If InStr(1, txtLine(lineNum), Chr(96)) <> 0 Then
                    outfitNameExtract = sf.cleanStrArray(Split(txtLine(lineNum), Chr(96)))
                ElseIf InStr(1, txtLine(lineNum), Chr(34)) <> 0 Then
                    outfitNameExtract = sf.cleanStrArray(Split(txtLine(lineNum), Chr(34)))
                Else
                    outfitNameExtract = sf.cleanStrArray(Split(txtLine(lineNum), " "))
                End If
                
                startDataCell.Offset(offsetY, 0).Value = outfitNameExtract(UBound(outfitNameExtract))
                
                
                haveCategory = 0
                
                
                lineNum = lineNum + 1
                While InStr(1, txtLine(lineNum), Chr(9)) = 1
                    
                    If haveCategory = 1 Then
                        
                        
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
                            
                            
                        ElseIf (InStr(1, txtLine(lineNum), " ") <> 0 Or InStr(1, txtLine(lineNum), Chr(34)) <> 0) Then
                            
                            If InStr(1, txtLine(lineNum), Chr(34)) <> 0 Then
                                tmpArray = sf.cleanStrArray(Split(txtLine(lineNum), Chr(34)))
                            Else
                                tmpArray = sf.cleanStrArray(Split(txtLine(lineNum), " "))
                            End If
                            
                            
                            offsetX = sf.inStringArray(heading, tmpArray(LBound(tmpArray)))
                            
                            If offsetX > -1 Then
                                startDataCell.Offset(offsetY, offsetX).Value = tmpArray(UBound(tmpArray))
                            End If
                            
                        End If
                        
                        
                    ElseIf InStr(1, txtLine(lineNum), "category ""Hand to Hand""") = 2 Then
                        
                        haveCategory = 1
                        foundCategory = True
                        
                    End If
                    
                    
                    lineNum = lineNum + 1
                Wend
                
                
                If foundCategory = True Then
                    offsetY = offsetY + 1
                    foundCategory = False
                End If
                
                
            End If
            
            
        Next lineNum
        
        
        Close #1
    Next pathNum
    
    ThisWorkbook.Worksheets("Hand to Hand").range(startDataCell.Offset(offsetY, 0).Address & ":" & startDataCell.Offset(offsetY, UBound(heading)).Address).ClearContents
    
    MsgBox ("Hand to hand outfit data imported")
End Sub



