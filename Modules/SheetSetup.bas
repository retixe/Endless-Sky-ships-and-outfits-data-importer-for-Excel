Attribute VB_Name = "SheetSetup"
Option Explicit

Sub Setup()
    
    
    Dim gameFilePath As String
    '---------------------------------------------------------------------------
    gameFilePath = "C:\Program Files (x86)\Steam\steamapps\common\Endless Sky" ' <<<< Edit this
    '---------------------------------------------------------------------------
    
    
    Dim sheetNameList() As Variant
    sheetNameList = Array("filepath", "Ships", "Guns", "Secondary Weapons", "turrets", "Systems", "Power", "Engines", "Hand to Hand", "Unique")
    
    Dim sheetName As Variant
    For Each sheetName In sheetNameList
        
        If sheetExist(sheetName) = False Then
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = sheetName
        End If
        
    Next sheetName
    
    
    
    
    ThisWorkbook.Worksheets("filepath").range("b2").Value = "Ship file path"
    ThisWorkbook.Worksheets("filepath").range("d2").Value = "Outfit file path"
    
    gameFilePath = gameFilePath & "\data\"
    
    Dim shipPathList() As Variant
    shipPathList = Array("coalition\coalition ships.txt", "drak\drak ships.txt", "drak\indigenous.txt", "human\ships.txt", "human\marauders.txt", "hai\hai ships.txt", "kahet\kahet ships.txt", "korath\korath ships.txt", "pug\pug.txt", "quarg\quarg ships.txt", "remnant\remnant ships.txt", "sheragi\sheragi ships.txt", "wanderer\wanderer ships.txt")
    
    Dim outfitPathList As Variant
    outfitPathList = Array("coalition\coalition outfits.txt", "coalition\coalition weapons.txt", "drak\drak outfits.txt", "drak\indigenous.txt", "hai\hai outfits.txt", "human\weapons.txt", "human\outfits.txt", "human\engines.txt", "human\power.txt", "kahet\kahet outfits.txt", "korath\korath weapons.txt", "korath\korath outfits.txt", "pug\pug.txt", "quarg\quarg outfits.txt", "remnant\remnant outfits.txt", "sheragi\sheragi outfits.txt", "wanderer\wanderer outfits.txt")
    
    Dim offsetY As Integer
    For offsetY = LBound(shipPathList) To UBound(shipPathList)
        ThisWorkbook.Worksheets("filepath").range("b3").Offset(offsetY, 0).Value = gameFilePath & shipPathList(offsetY)
    Next offsetY
    For offsetY = LBound(outfitPathList) To UBound(outfitPathList)
        ThisWorkbook.Worksheets("filepath").range("d3").Offset(offsetY, 0).Value = gameFilePath & outfitPathList(offsetY)
    Next offsetY
    
    
    
    
    
    ThisWorkbook.Worksheets("Ships").range("c2:u2").Value = Array("licenses", "cost", "category", "shields", "hull", "gun", "turret", "required crew", "bunks", "fuel capacity", "cargo space", "outfit space", "weapon capacity", "engine capacity", "mass", "drag", "cloak", "fighter bay", "drone bay")
    ThisWorkbook.Worksheets("Guns").range("c2:l2").Value = Array("licenses", "cost", "outfit space", "shield damage", "hull damage", "reload", "firing energy", "firing heat", "range", "blast radius")
    ThisWorkbook.Worksheets("Secondary Weapons").range("c2:l2").Value = Array("licenses", "cost", "outfit space", "shield damage", "hull damage", "reload", "firing energy", "firing heat", "range", "blast radius")
    ThisWorkbook.Worksheets("turrets").range("c2:m2").Value = Array("licenses", "cost", "outfit space", "shield damage", "hull damage", "reload", "firing energy", "firing heat", "range", "blast radius", "anti-missile")
    ThisWorkbook.Worksheets("Systems").range("c2:n2").Value = Array("licenses", "cost", "mass", "outfit space", "energy consumption", "heat generation", "shield generation", "shield energy", "shield heat", "hull repair rate", "hull energy", "hull heat")
    ThisWorkbook.Worksheets("Power").range("c2:i2").Value = Array("licenses", "cost", "mass", "outfit space", "energy generation", "heat generation", "energy capacity")
    ThisWorkbook.Worksheets("Engines").range("c2:w2").Value = Array("licenses", "cost", "mass", "outfit space", "engine capacity", "thrust", "thrusting energy", "thrusting heat", "cooling", "energy consumption", "heat generation", "turn", "turning energy", "turning heat", "reverse thrust", "reverse thrusting energy", "reverse thrusting heat", "afterburner thrust", "afterburner energy", "afterburner heat", "afterburner fuel")
    ThisWorkbook.Worksheets("Hand to Hand").range("c2:g2").Value = Array("licenses", "cost", "capture attack", "capture defense", "outfit space")
    ThisWorkbook.Worksheets("Unique").range("c2:d2").Value = Array("licenses", "cost")
    
    
    
    
    Call Ship.Data
    Call GunOutfit.Data
    Call SecondaryWeaponOutfit.Data
    Call TurretOutfit.Data
    Call SystemOutfit.Data
    Call PowerOutfit.Data
    Call EngineOutfit.Data
    Call HandToHandOutfit.Data
    Call UniqueOutfit.Data
    
    
End Sub


Function sheetExist(ByRef targetName)
    
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        
        If sheet.Name = targetName Then
            sheetExist = True
            Exit Function
        End If
        
    Next sheet
    
    
    sheetExist = False
End Function
