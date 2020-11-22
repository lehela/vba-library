Attribute VB_Name = "FormatConditionsTool"
Option Explicit

' ------------------------------------------------------------------------------------------------------------
'
' Public Sub
'
' ------------------------------------------------------------------------------------------------------------

Public Sub Save()
    
    Dim dict As Dictionary
    Set dict = serializeFormatConditions(ActiveSheet.Cells.FormatConditions)
    
    Dim json As String
    json = JsonConverter.ConvertToJson(dict)
    
    KeyValueStore.SetValue "FormatConditions", json
    
    Clipboard.SetClipboard json
    
End Sub


Public Sub Restore()

    Dim json As String
    json = KeyValueStore.GetValue("FormatConditions")

    ' Place a copy in the clipboard
    Clipboard.SetClipboard json

    Dim dict As Dictionary
    Set dict = JsonConverter.ParseJson(json)

    deserializeFormatConditions dict

End Sub

' ------------------------------------------------------------------------------------------------------------
'
' Format Conditions
'
' ------------------------------------------------------------------------------------------------------------


Public Function serializeFormatConditions(ByRef obj As FormatConditions) As Dictionary

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    Dim fc As Object
    Dim idx As Integer: idx = 0
    Dim fc_name As String
    
    
    For Each fc In obj
        idx = idx + 1
        
        fc.AppliesTo.Select
        
        fc_name = Format(idx, "00_") & TypeName(fc) & "_" & fc.AppliesTo.Address
        
        Select Case TypeName(fc)
        
        Case Is = "FormatCondition": dict.Add fc_name, serializeFormatCondition(fc)
        Case Is = "ColorScale": dict.Add fc_name, serializeColorScale(fc)
        Case Is = "IconSetCondition": dict.Add fc_name, serializeIconSetCondition(fc)
        Case Is = "Databar": dict.Add fc_name, serializeDatabar(fc)
        Case Is = "Top10": dict.Add fc_name, serializeTop10(fc)
        Case Is = "AboveAverage": dict.Add fc_name, serializeAboveAverage(fc)
        Case Is = "UniqueValues": dict.Add fc_name, serializeUniqueValues(fc)
        End Select
        
    Next

    Set serializeFormatConditions = dict
    
End Function

Public Function deserializeFormatConditions(ByRef fcs As Dictionary, Optional Worksheet As Worksheet)
    
    ' Delete existing Format Conditions
    ws(Worksheet).Cells.FormatConditions.Delete
    
    ' Restore Format Conditions from Dictionary
    Dim key
    Dim obj As Object
    
    On Error GoTo continue_
    
    For Each key In fcs.Keys
        Set obj = fcs(key)
        
        ws(Worksheet).Range(obj("AppliesTo")).Select
        
        Select Case TypeName(obj)
        Case Is = "Dictionary"
            Select Case obj("Class")
                Case Is = "FormatCondition": deserializeFormatCondition obj, Worksheet
                Case Is = "ColorScale": deserializeColorScale obj, Worksheet
                Case Is = "IconSetCondition": deserializeIconSetCondition obj, Worksheet
                Case Is = "Databar": deserializeDatabar obj, Worksheet
                Case Is = "Top10": deserializeTop10 obj, Worksheet
                Case Is = "AboveAverage": deserializeAboveAverage obj, Worksheet
                Case Is = "UniqueValues": deserializeUniqueValues obj, Worksheet
            End Select
        End Select
continue_:
    Next
    
End Function

Public Function serializeFormatCondition(ByRef obj As FormatCondition)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        
        dict.Add "AppliesTo", .AppliesTo.Address
        dict.Add "AppliesToLocal", .AppliesTo.AddressLocal
        
        dict.Add "Type", .Type
        dict.Add "Operator", .Operator
        dict.Add "DateOperator", .DateOperator
        dict.Add "Formula1", .Formula1
        dict.Add "Formula2", .Formula2
        
        dict.Add "NumberFormat", .NumberFormat
        dict.Add "StopIfTrue", .StopIfTrue
        dict.Add "Priority", .Priority
        dict.Add "PTCondition", .PTCondition
        dict.Add "ScopeType", .ScopeType
        dict.Add "Text", .Text
        dict.Add "TextOperator", .TextOperator
        
        dict.Add "Font", serializeFont(.Font)
        dict.Add "Interior", serializeInterior(.Interior)
        dict.Add "Borders", serializeBorders(.borders)
        
    End With
    
    Set serializeFormatCondition = dict
    
End Function

Public Function deserializeFormatCondition(ByRef dict As Dictionary, Optional Worksheet As Worksheet)

    Dim rgAppliesTo As Range
    Set rgAppliesTo = ws(Worksheet).Range(dict("AppliesTo"))

    rgAppliesTo.Select

    Dim fc As FormatCondition
    
    Select Case dict("Type")
    
    ' Expression
    Case XlFormatConditionType.xlExpression
        Set fc = rgAppliesTo.FormatConditions.Add( _
            Type:=dict("Type"), _
            Formula1:=dict("Formula1") _
            )
            
    ' Others
    Case Else
        Set fc = rgAppliesTo.FormatConditions.Add( _
            Type:=dict("Type"), _
            Operator:=dict("Operator"), _
            Formula1:=dict("Formula1"), _
            Formula2:=dict("Formula2") _
            )
            
    End Select
    
    On Error Resume Next
    With fc
                
        .ScopeType = dict("ScopeType")
        .DateOperator = dict("DateOperator")
        .StopIfTrue = dict("StopIfTrue")
        .Priority = dict("Priority")
        .NumberFormat = IfEmpty(dict("NumberFormat"), "General")
        
        If Not IsEmpty(dict("Text")) Then .Text = dict("Text")
        If Not IsEmpty(dict("TextOperator")) Then .TextOperator = dict("TextOperator")
        
        deserializeInterior dict("Interior"), .Interior
        deserializeFont dict("Font"), .Font
        deserializeBorders dict("Borders"), .borders
        
    End With

End Function

Public Function serializeTop10(ByRef obj As Top10) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "AppliesTo", .AppliesTo.Address
        dict.Add "AppliesToLocal", .AppliesTo.AddressLocal
        dict.Add "Type", .Type
        
        dict.Add "CalcFor", .CalcFor
        dict.Add "NumberFormat", .NumberFormat
        dict.Add "Percent", .Percent
        dict.Add "Priority", .Priority
        dict.Add "Rank", .Rank
        dict.Add "ScopeType", .ScopeType
        dict.Add "StopIfTrue", .StopIfTrue
        dict.Add "TopBottom", .TopBottom
        
        dict.Add "Borders", serializeBorders(.borders)
        dict.Add "Font", serializeFont(.Font)
        dict.Add "Interior", serializeInterior(.Interior)
        
    End With

    Set serializeTop10 = dict
    
End Function

Public Function deserializeTop10(ByRef dict As Dictionary, Optional Worksheet As Worksheet)
    
    Dim rgAppliesTo As Range
    Set rgAppliesTo = ws(Worksheet).Range(dict("AppliesTo"))

    rgAppliesTo.Select
    
    Dim fc As Top10
    Set fc = rgAppliesTo.FormatConditions.AddTop10
        
    On Error Resume Next
    
    With fc
        
        .CalcFor = dict("CalcFor")
        .NumberFormat = IfEmpty(dict("NumberFormat"), "General")
        .Percent = dict("Percent")
        .Priority = dict("Priority")
        .Rank = dict("Rank")
        .ScopeType = dict("ScopeType")
        .StopIfTrue = dict("StopIfTrue")
        .TopBottom = dict("TopBottom")
        
        deserializeBorders dict("Borders"), .borders
        deserializeFont dict("Font"), .Font
        deserializeInterior dict("Interior"), .Interior
    
    End With
    
End Function


Public Function serializeAboveAverage(ByRef obj As AboveAverage) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "AppliesTo", .AppliesTo.Address
        dict.Add "AppliesToLocal", .AppliesTo.AddressLocal
        dict.Add "Type", .Type
        
        dict.Add "AboveBelow", .AboveBelow
        dict.Add "CalcFor", .CalcFor
        dict.Add "NumberFormat", .NumberFormat
        dict.Add "NumStdDev", .NumStdDev
        dict.Add "Priority", .Priority
        dict.Add "ScopeType", .ScopeType
        dict.Add "StopIfTrue", .StopIfTrue
        
        dict.Add "Borders", serializeBorders(.borders)
        dict.Add "Font", serializeFont(.Font)
        dict.Add "Interior", serializeInterior(.Interior)
        
    End With

    Set serializeAboveAverage = dict
    
End Function

Public Function deserializeAboveAverage(ByRef dict As Dictionary, Worksheet As Worksheet)
    
    Dim rgAppliesTo As Range
    Set rgAppliesTo = ws(Worksheet).Range(dict("AppliesTo"))

    rgAppliesTo.Select
    
    Dim fc As AboveAverage
    Set fc = rgAppliesTo.FormatConditions.AddAboveAverage
        
    On Error Resume Next
    
    With fc
    
        .AboveBelow = dict("AboveBelow")
        .CalcFor = dict("CalcFor")
        .NumberFormat = IfEmpty(dict("NumberFormat"), "General")
        .NumStdDev = dict("NumStdDev")
        .Priority = dict("Priority")
        .ScopeType = dict("ScopeType")
        .Rank = dict("Rank")
        .StopIfTrue = dict("StopIfTrue")
        
        deserializeBorders dict("Borders"), .borders
        deserializeFont dict("Font"), .Font
        deserializeInterior dict("Interior"), .Interior
        
    End With
    
End Function

Public Function serializeUniqueValues(ByRef obj As UniqueValues) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "AppliesTo", .AppliesTo.Address
        dict.Add "AppliesToLocal", .AppliesTo.AddressLocal
        dict.Add "Type", .Type
    
        dict.Add "DupeUnique", .DupeUnique
        dict.Add "NumberFormat", .NumberFormat
        dict.Add "Priority", .Priority
        dict.Add "ScopeType", .ScopeType
        dict.Add "StopIfTrue", .StopIfTrue
        
        dict.Add "Borders", serializeBorders(.borders)
        dict.Add "Font", serializeFont(.Font)
        dict.Add "Interior", serializeInterior(.Interior)
        
    End With

    Set serializeUniqueValues = dict
    
End Function

Public Function deserializeUniqueValues(ByRef dict As Dictionary, Worksheet As Worksheet)
    
    Dim rgAppliesTo As Range
    Set rgAppliesTo = ws(Worksheet).Range(dict("AppliesTo"))

    rgAppliesTo.Select
    
    Dim fc As UniqueValues
    Set fc = rgAppliesTo.FormatConditions.AddUniqueValues
        
    On Error Resume Next
    
    With fc
        .DupeUnique = dict("DupeUnique")
        .NumberFormat = IfEmpty(dict("NumberFormat"), "General")
        .Priority = dict("Priority")
        .ScopeType = dict("ScopeType")
        .StopIfTrue = dict("StopIfTrue")
        
        deserializeBorders dict("Borders"), .borders
        deserializeFont dict("Font"), .Font
        deserializeInterior dict("Interior"), .Interior
        
    End With
    
End Function



' ------------------------------------------------------------------------------------------------------------
'
' Font
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeFont(ByRef obj As Font)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Name", .Name
        dict.Add "Background", .Background
        dict.Add "Bold", .Bold
        dict.Add "Color", .Color
        dict.Add "ColorIndex", .ColorIndex
        dict.Add "FontStyle", .FontStyle
        dict.Add "Size", .Size
        dict.Add "Strikethrough", .Strikethrough
        dict.Add "Subscript", .Subscript
        dict.Add "Superscript", .Superscript
        dict.Add "ThemeColor", .ThemeColor
        dict.Add "ThemeFont", .ThemeFont
        dict.Add "TintAndShade", .TintAndShade
        dict.Add "Underline", .Underline
    End With
    
    Set serializeFont = dict
    
End Function

Public Function deserializeFont(ByRef dict As Dictionary, obj As Font)

    On Error Resume Next

    With obj
        
        .Name = dict("Name")
        
        If dict("ThemeColor") <> 0 Then
            .ThemeFont = dict("ThemeFont")
            .ThemeColor = dict("ThemeColor")
        Else
            .ColorIndex = IfEmpty(dict("ColorIndex"), XlColorIndex.xlColorIndexAutomatic)
            .Color = dict("Color")
        End If
        
        .TintAndShade = CDbl(dict("TintAndShade"))
        .FontStyle = dict("FontStyle")
        .Italic = dict("Italic")
        .Size = dict("Size")
        .Strikethrough = dict("Strikethrough")
        .Subscript = dict("Subscript")
        .Superscript = dict("Superscript")
        .Underline = dict("Underline")
    
    End With

End Function

' ------------------------------------------------------------------------------------------------------------
'
' Interior
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeInterior(ByRef obj As Interior) As Dictionary

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        Select Case True
        
        ' Default Case, no pattern used
        Case IsNull(.Pattern)
            dict.Add "Color", .Color
            dict.Add "ColorIndex", .ColorIndex
            dict.Add "InvertIfNegative", .InvertIfNegative
            dict.Add "ThemeColor", .ThemeColor
            dict.Add "TintAndShade", .TintAndShade
        
        ' Gradient Patterns
        Case .Pattern = XlPattern.xlPatternRectangularGradient
            dict.Add "Pattern", .Pattern
            dict.Add "Gradient", serializeRectangularGradient(.Gradient)
                        
        Case .Pattern = XlPattern.xlPatternLinearGradient
            dict.Add "Pattern", .Pattern
            dict.Add "Gradient", serializeLinearGradient(.Gradient)
        
        ' Solid Color
        Case .Pattern = XlPattern.xlPatternSolid
            dict.Add "Pattern", .Pattern
            dict.Add "Color", .Color
            dict.Add "ColorIndex", .ColorIndex
            dict.Add "InvertIfNegative", .InvertIfNegative
            dict.Add "ThemeColor", .ThemeColor
            dict.Add "TintAndShade", .TintAndShade
        
        ' Any other Pattern
        Case Else
            dict.Add "Pattern", .Pattern
            dict.Add "PatternColor", .PatternColor
            dict.Add "PatternColorIndex", .PatternColorIndex
            dict.Add "PatternThemeColor", .PatternThemeColor
            dict.Add "PatternTintAndShade", .PatternTintAndShade
        End Select
        
    End With

    Set serializeInterior = dict
    
End Function

Public Function deserializeInterior(ByRef dict As Dictionary, ByRef intr As Interior)

    On Error Resume Next
    
    With intr
        
        .InvertIfNegative = dict("InvertIfNegative")
        
        ' Patterns
        If dict("Pattern") <> 0 Then
            .Pattern = dict("Pattern")
            
            Select Case True
            
            Case .Pattern = XlPattern.xlPatternRectangularGradient
                deserializeRectangularGradient dict("Gradient"), .Gradient
                
            Case .Pattern = XlPattern.xlPatternLinearGradient
                deserializeLinearGradient dict("Gradient"), .Gradient
            
            Case .Pattern = XlPattern.xlPatternSolid
                .ColorIndex = dict("ColorIndex")
                .Color = dict("Color")
                .TintAndShade = CDbl(dict("TintAndShade"))
                
            Case Else
                .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                If dict("PatternThemeColor") <> 0 Then
                    .PatternThemeColor = dict("PatternThemeColor")
                Else
                    .PatternColor = dict("PatternColor")
                End If
                .PatternTintAndShade = CDbl(dict("PatternTintAndShade"))
            
            End Select
            
            Exit Function
        End If
        
        ' No Pattern, with Themes
        If dict("ThemeColor") <> 0 Then
            .PatternColorIndex = XlPattern.xlPatternAutomatic
            .ThemeColor = dict("ThemeColor")
            .TintAndShade = CDbl(dict("TintAndShade"))
            Exit Function
        End If
        
        ' No Pattern, no Themes
        If dict("Color") <> 0 Then
            .PatternColorIndex = XlPattern.xlPatternAutomatic
            .ColorIndex = IfEmpty(dict("ColorIndex"), XlColorIndex.xlColorIndexAutomatic)
            .Color = dict("Color")
            .TintAndShade = CDbl(dict("TintAndShade"))
            Exit Function
        End If

    
    End With

End Function


Public Function serializeRectangularGradient(ByRef obj As RectangularGradient) As Dictionary

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "RectangleTop", .RectangleTop
        dict.Add "RectangleBottom", .RectangleBottom
        dict.Add "RectangleLeft", .RectangleLeft
        dict.Add "RectangleRight", .RectangleRight
        dict.Add "ColorStops", serializeColorStops(.ColorStops)
    End With

    Set serializeRectangularGradient = dict

End Function

Public Function deserializeRectangularGradient(ByRef dict As Dictionary, obj As RectangularGradient)
    
    On Error Resume Next
    
    With obj
        .RectangleTop = dict("RectangleTop")
        .RectangleBottom = dict("RectangleBottom")
        .RectangleLeft = dict("RectangleLeft")
        .RectangleRight = dict("RectangleRight")
        
        deserializeColorStops dict("ColorStops"), .ColorStops
        
    End With
    
End Function



Public Function serializeColorStops(ByRef obj As ColorStops) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    Dim cs As ColorStop
    Dim idx: idx = 1
    
    On Error Resume Next
    
    With obj
        For Each cs In obj
            dict.Add Format(idx, "000_"), serializeColorStop(cs)
            idx = idx + 1
        Next
    End With

    Set serializeColorStops = dict
End Function

Public Function deserializeColorStops(ByRef dict As Dictionary, obj As ColorStops)
    
    Dim key
    
    On Error Resume Next
    
    obj.Clear
    
    For Each key In dict.Keys
        'If key <> "Class" Then
            deserializeColorStop dict(key), obj.Add(dict(key)("Position"))
        'End If
    Next
    
End Function

Public Function serializeColorStop(ByRef obj As ColorStop) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Position", .Position
        dict.Add "Color", .Color
        dict.Add "ThemeColor", .ThemeColor
        dict.Add "TintAndShade", .TintAndShade
    End With

    Set serializeColorStop = dict
    
End Function

Public Function deserializeColorStop(ByRef dict As Dictionary, obj As ColorStop)
    
    On Error Resume Next
    
    With obj
        If dict("ThemeColor") <> 0 Then
            .ThemeColor = dict("ThemeColor")
        Else
            .Color = dict("Color")
        End If
        .TintAndShade = CDbl(dict("TintAndShade"))
    End With
    
End Function

Public Function serializeLinearGradient(ByRef obj As LinearGradient) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Degree", .Degree
        dict.Add "ColorStops", serializeColorStops(.ColorStops)
    End With

    Set serializeLinearGradient = dict
    
End Function

Public Function deserializeLinearGradient(ByRef dict As Dictionary, obj As LinearGradient)
    
    On Error Resume Next
    
    With obj
        .Degree = dict("Degree")
        deserializeColorStops dict("ColorStops"), .ColorStops
        
    End With
    
End Function

' ------------------------------------------------------------------------------------------------------------
'
' Borders
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeBorders(ByRef obj As borders)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
'        dict.Add "Color", .Color
'        dict.Add "ColorIndex", .ColorIndex
'        dict.Add "LineStyle", .LineStyle
'        dict.Add "ThemeColor", .ThemeColor
'        dict.Add "TintAndShade", .TintAndShade
'        dict.Add "Value", .value
'        dict.Add "Weight", .Weight
        
        Dim dictBorders As New Dictionary
        dictBorders.Add "xlLeft", serializeBorder(obj(xlLeft), xlLeft)
        dictBorders.Add "xlRight", serializeBorder(obj(xlRight), xlRight)
        dictBorders.Add "xlTop", serializeBorder(obj(xlTop), xlTop)
        dictBorders.Add "xlBottom", serializeBorder(obj(xlBottom), xlBottom)
        
        dict.Add "Borders", dictBorders
    End With
    
    Set serializeBorders = dict
    
End Function

Public Function deserializeBorders(ByRef dict As Dictionary, ByRef brdrs As borders)

    On Error Resume Next
    
    Dim dictBorder As Dictionary
    
    With brdrs
    
        Set dictBorder = dict("Borders")("xlLeft"): deserializeBorder dictBorder, brdrs(xlLeft)
        Set dictBorder = dict("Borders")("xlRight"): deserializeBorder dictBorder, brdrs(xlRight)
        Set dictBorder = dict("Borders")("xlTop"): deserializeBorder dictBorder, brdrs(xlTop)
        Set dictBorder = dict("Borders")("xlBottom"): deserializeBorder dictBorder, brdrs(xlBottom)
    
    End With
    
End Function

Public Function serializeBorder(ByRef obj As Border, objIndex As XlBordersIndex)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        If .LineStyle = xlNone Or IsNull(.LineStyle) Then
            dict.Add "Active", False
        Else
            dict.Add "Active", True
            dict.Add "BorderIndex", objIndex
            dict.Add "Color", .Color
            dict.Add "ColorIndex", .ColorIndex
            dict.Add "LineStyle", .LineStyle
            dict.Add "ThemeColor", .ThemeColor
            dict.Add "Weight", .Weight
            dict.Add "TintAndShade", .TintAndShade
        End If
    End With
    
    Set serializeBorder = dict
    
End Function

Public Function deserializeBorder(ByRef dict As Dictionary, ByRef brdr As Border)

    On Error Resume Next
    
    If dict("Active") = False Then Exit Function
    
    With brdr
        .ColorIndex = dict("ColorIndex")
        .Color = dict("Color")
        .LineStyle = dict("LineStyle")
        .ThemeColor = dict("ThemeColor")
        .TintAndShade = CDbl(dict("TintAndShade"))
        .Weight = dict("Weight")
    End With
      
End Function

' ------------------------------------------------------------------------------------------------------------
'
' Color Scale
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeColorScale(ByRef obj As ColorScale)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Class", "ColorScale"
        dict.Add "AppliesTo", .AppliesTo.Address
        dict.Add "Type", .Type
        dict.Add "ColorScaleType", .ColorScaleCriteria.Count
        dict.Add "ColorScaleCriteria", serializeColorScaleCriteria(.ColorScaleCriteria)
        dict.Add "Formula", .Formula
        dict.Add "StopIfTrue", .StopIfTrue
        dict.Add "Priority", .Priority
        dict.Add "PTCondition", .PTCondition
        dict.Add "ScopeType", .ScopeType
        
        dict.Add "Font", serializeFont(.Font)
        dict.Add "Interior", serializeInterior(.Interior)
        dict.Add "Borders", serializeBorders(.borders)
    End With
    
    Set serializeColorScale = dict
    
End Function

Public Function deserializeColorScale(ByRef dict As Dictionary, Optional Worksheet As Worksheet)

    Dim cs As ColorScale
    Set cs = ws(Worksheet).Range(dict("AppliesTo")).FormatConditions.AddColorScale( _
            ColorScaleType:=dict("ColorScaleType") _
            )
    
    On Error Resume Next
    With cs
        .ScopeType = dict("ScopeType")
        .Formula = dict("Formula")
        ' .StopIfTrue is read-only property for ColorScale
        .Priority = dict("Priority")
        .NumberFormat = IfEmpty(dict("NumberFormat"), "General")
        ' .PTCondition is read-only property
        
        deserializeFont dict("Font"), .Font
        deserializeBorders dict("Borders"), .borders
        deserializeColorScaleCriteria dict("ColorScaleCriteria"), .ColorScaleCriteria, dict("ColorScaleType")

    End With
    
End Function

Public Function serializeColorScaleCriteria(ByRef obj As ColorScaleCriteria)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    Dim cscriterion As ColorScaleCriterion
    Dim idx As Integer: idx = 0
    
    On Error Resume Next
    For Each cscriterion In obj
        idx = idx + 1
        dict.Add Format(idx, "000"), serializeColorScaleCriterion(cscriterion)
    Next

    Set serializeColorScaleCriteria = dict
    
End Function

Public Function deserializeColorScaleCriteria(ByRef dict As Dictionary, ByRef csa As ColorScaleCriteria, ByVal ColorScaleType)

    On Error Resume Next
    
    With csa
        deserializeColorScaleCriterion dict("001"), .Item(1)
        deserializeColorScaleCriterion dict("002"), .Item(2)
        If ColorScaleType = 3 Then
            deserializeColorScaleCriterion dict("003"), .Item(3)
        End If
    
    End With
    
End Function

Public Function serializeColorScaleCriterion(ByRef obj As ColorScaleCriterion)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Index", .Index
        dict.Add "Type", .Type
        dict.Add "Value", .value
        dict.Add "FormatColor", serializeFormatColor(.FormatColor)
    End With
    
    Set serializeColorScaleCriterion = dict
    
End Function

Public Function deserializeColorScaleCriterion(ByRef dict As Dictionary, ByRef csn As ColorScaleCriterion)
    
    On Error Resume Next
    With csn
        .Type = dict("Type")
        .value = CDbl(dict("Value"))
        
        deserializeFormatColor dict("FormatColor"), .FormatColor
                
    End With
    
End Function

' ------------------------------------------------------------------------------------------------------------
'
' Format Color
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeFormatColor(ByRef obj As FormatColor) As Dictionary

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Class", "FormatColor"
        dict.Add "Color", .Color
        dict.Add "ColorIndex", .ColorIndex
        dict.Add "ThemeColor", .ThemeColor
        dict.Add "TintAndShade", .TintAndShade
    End With
    
    Set serializeFormatColor = dict
    
End Function

Public Function deserializeFormatColor(ByRef dict As Dictionary, ByRef fc As FormatColor)
    
    On Error Resume Next
    With fc
        .TintAndShade = CDbl(dict("TintAndShade"))
        .ThemeColor = dict("ThemeColor")
        .ColorIndex = dict("ColorIndex")
        .Color = dict("Color")
    End With
End Function

' ------------------------------------------------------------------------------------------------------------
'
' IconSetCondition
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeIconSetCondition(ByRef obj As IconSetCondition) As Dictionary

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Class", "IconSetCondition"
        dict.Add "AppliesTo", .AppliesTo.Address
        dict.Add "AppliesToLocal", .AppliesTo.AddressLocal
        dict.Add "Type", .Type
        dict.Add "Formula", .Formula
        
        dict.Add "PercentileValues", .PercentileValues
        dict.Add "Priority", .Priority
        dict.Add "ReverseOrder", .ReverseOrder
        dict.Add "ScopeType", .ScopeType
        dict.Add "ShowIconOnly", .ShowIconOnly
        dict.Add "StopIfTrue", .StopIfTrue
        
        dict.Add "IconSet", .IconSet.ID
        dict.Add "IconCriteria", serializeIconCriteria(.IconCriteria)
    End With
    
    Set serializeIconSetCondition = dict
    
End Function


Public Function deserializeIconSetCondition(ByRef dict As Dictionary, Optional Worksheet As Worksheet)

    Dim isc As IconSetCondition
    Set isc = ws(Worksheet).Range(dict("AppliesTo")).FormatConditions.AddIconSetCondition
    
    On Error Resume Next
    With isc
        .Formula = dict("Formula")
        .ShowIconOnly = dict("ShowIconOnly")
        .PercentileValues = dict("PercentileValues")
        .Priority = dict("Priority")
        .ReverseOrder = dict("ReverseOrder")
        .ScopeType = dict("ScopeType")
        '.StopIfTrue = dict("StopIfTrue")
        '.Type = dict("Type")
        .IconSet = ActiveWorkbook.IconSets(dict("IconSet"))
        deserializeIconCriteria dict("IconCriteria"), .IconCriteria
    End With
    
End Function

Public Function serializeIconCriteria(ByRef obj As IconCriteria)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    Dim icn As IconCriterion
    
    On Error Resume Next
    For Each icn In obj
        dict.Add Format(icn.Index, "000"), serializeIconCriterion(icn)
    Next
    Set serializeIconCriteria = dict
    
End Function

Public Function deserializeIconCriteria(dict As Dictionary, ica As IconCriteria)

    Dim key
    Dim icn As Icon
    
    On Error Resume Next
    For Each key In dict.Keys
        deserializeIconCriterion dict(key), ica(Int(key))
    Next
End Function

Public Function serializeIconCriterion(ByRef obj As IconCriterion)

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Icon", .Icon
        dict.Add "Operator", .Operator
        dict.Add "Type", .Type
        dict.Add "Value", .value
    End With
    
    Set serializeIconCriterion = dict
    
End Function

Public Function deserializeIconCriterion(ByRef dict As Dictionary, icn As IconCriterion)

    On Error Resume Next
    icn.Icon = dict("Icon")
    icn.Operator = dict("Operator")
    icn.Type = dict("Type")
    icn.value = dict("Value")
    
End Function

' ------------------------------------------------------------------------------------------------------------
'
' DataBar
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeDatabar(ByRef obj As Databar) As Dictionary

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Class", "Databar"
        dict.Add "AppliesTo", .AppliesTo.Address
        dict.Add "AppliesToLocal", .AppliesTo.AddressLocal
        dict.Add "Type", .Type
        dict.Add "AxisColor", serializeFormatColor(.AxisColor)
        dict.Add "AxisPosition", IfNull(.AxisPosition, xlDataBarAxisAutomatic)
        dict.Add "BarBorder", serializeDataBarBorder(.BarBorder)
        dict.Add "BarColor", serializeFormatColor(.BarColor)
        dict.Add "BarFillType", .BarFillType
        dict.Add "Direction", .Direction
        dict.Add "Formula", .Formula
        dict.Add "MaxPoint", serializeConditionValue(.MaxPoint)
        dict.Add "MinPoint", serializeConditionValue(.MinPoint)
        dict.Add "NegativeBarFormat", serializeNegativeBarFormat(.NegativeBarFormat)
        dict.Add "PercentMax", .PercentMax
        dict.Add "PercentMin", .PercentMin
        dict.Add "Priority", .Priority
        dict.Add "ScopeType", .ScopeType
        dict.Add "ShowValue", .ShowValue
        dict.Add "StopIfTrue", .StopIfTrue
        dict.Add "Type", .Type
    End With
    
    Set serializeDatabar = dict
    
End Function

Public Function deserializeDatabar(ByRef dict As Dictionary, Optional Worksheet As Worksheet)

    Dim dbar As Databar
    Set dbar = ws(Worksheet).Range(dict("AppliesTo")).FormatConditions.AddDatabar
    
    On Error Resume Next
    With dbar
        deserializeFormatColor dict("AxisColor"), .AxisColor
        .AxisPosition = dict("AxisPosition")
        deserializeDataBarBorder dict("BarBorder"), .BarBorder
        deserializeFormatColor dict("BarColor"), .BarColor
        .BarFillType = dict("BarFillType")
        .Direction = dict("Direction")
        .Formula = dict("Formula")
        deserializeConditionValue dict("MaxPoint"), .MaxPoint
        deserializeConditionValue dict("MinPoint"), .MinPoint
        deserializeNegativeBarFormat dict("NegativeBarFormat"), .NegativeBarFormat
        .PercentMax = dict("PercentMax")
        .PercentMin = dict("PercentMin")
        .Priority = dict("Priority")
        .ScopeType = dict("ScopeType")
        .ShowValue = dict("ShowValue")
        ' .StopIfTrue = dict("StopIfTrue")
        '.Type = dict("Type")
    
    End With
    
End Function

Public Function serializeConditionValue(ByRef obj As ConditionValue) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "Class", "ConditionValue"
        dict.Add "Type", .Type
        dict.Add "Value", .value
    End With
    
    Set serializeConditionValue = dict
    
End Function

Public Function deserializeConditionValue(ByRef dict As Dictionary, ByRef cv As ConditionValue)
    
    On Error Resume Next
    
    With cv
        .Modify dict("Type"), dict("Value")
    End With
    
End Function


Public Function serializeNegativeBarFormat(ByRef obj As NegativeBarFormat) As Dictionary
    
    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)
    
    On Error Resume Next
    
    With obj
        dict.Add "BorderColor", serializeFormatColor(.BorderColor)
        dict.Add "BorderColorType", .BorderColorType
        dict.Add "Color", serializeFormatColor(.Color)
        dict.Add "ColorType", .ColorType
    End With
    
    Set serializeNegativeBarFormat = dict
    
End Function

Public Function deserializeNegativeBarFormat(ByRef dict As Dictionary, ByRef nbf As NegativeBarFormat)
    
    On Error Resume Next
    
    With nbf
        deserializeFormatColor dict("BorderColor"), .BorderColor
        .BorderColorType = dict("BorderColorType")
        deserializeFormatColor dict("Color"), .Color
        .ColorType = dict("ColorType")
    End With
    
End Function

Public Function serializeDataBarBorder(ByRef obj As DataBarBorder) As Dictionary

    Dim dict As New Dictionary: dict.Add "Class", TypeName(obj)

    On Error Resume Next
    With obj
        dict.Add "Class", "DataBarBorder"
        dict.Add "Color", serializeFormatColor(.Color)
        dict.Add "Type", .Type
    End With
    
    Set serializeDataBarBorder = dict

End Function

Public Function deserializeDataBarBorder(ByRef dict As Dictionary, ByRef dbrd As DataBarBorder)
    
    On Error Resume Next
    
    With dbrd
        deserializeFormatColor dict("Color"), .Color
        .Type = dict("Type")
    End With
    
End Function

' ------------------------------------------------------------------------------------------------------------
'
' Helper Functions
'
' ------------------------------------------------------------------------------------------------------------

Private Property Get ws(Optional Worksheet As Worksheet) As Worksheet

    If Worksheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = Worksheet
    End If
    
End Property

Private Function IfNull(arg, nullVal)
    If IsNull(arg) Then
        IfNull = nullVal
    Else
        IfNull = arg
    End If
End Function

Private Function IfEmpty(arg, emptyVal)
    If IsEmpty(arg) Then
        IfEmpty = emptyVal
    Else
        IfEmpty = arg
    End If
End Function
