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
    
    Clipboard.SetString json
    
End Sub


Public Sub Restore()

    Dim json As String
    json = KeyValueStore.GetValue("FormatConditions")

    ' Place a copy in the clipboard
    Clipboard.SetString json

    Dim dict As Dictionary
    Set dict = JsonConverter.ParseJson(json)

    deserializeFormatConditions dict

End Sub

' ------------------------------------------------------------------------------------------------------------
'
' Format Conditions
'
' ------------------------------------------------------------------------------------------------------------


Public Function serializeFormatConditions(ByRef fcs As FormatConditions) As Dictionary

    Dim dict As Dictionary
    Set dict = New Dictionary
    
    On Error Resume Next
    
    Dim fc As Object
    Dim idx As Integer
    idx = 0
    For Each fc In fcs
        idx = idx + 1
        
        Select Case TypeName(fc)
        
        Case Is = "FormatCondition": dict.Add Format(idx, "000_") & TypeName(fc), serializeFormatCondition(fc)
        Case Is = "ColorScale": dict.Add Format(idx, "000_") & TypeName(fc), serializeColorScale(fc)
        Case Is = "IconSetCondition": dict.Add Format(idx, "000_") & TypeName(fc), serializeIconSetCondition(fc)
        Case Is = "Databar": dict.Add Format(idx, "000_") & TypeName(fc), serializeDatabar(fc)
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
    
    For Each key In fcs.Keys
        Set obj = fcs(key)
        
        Select Case TypeName(obj)
        Case Is = "Dictionary"
            Debug.Print obj("AppliesTo")
            Select Case obj("Class")
                Case Is = "FormatCondition": deserializeFormatCondition obj, Worksheet
                Case Is = "ColorScale": deserializeColorScale obj, Worksheet
                Case Is = "IconSetCondition": deserializeIconSetCondition obj, Worksheet
                Case Is = "Databar": deserializeDatabar obj, Worksheet
            End Select
        End Select
    Next
    
End Function

Public Function serializeFormatCondition(ByRef fc As FormatCondition)

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Class", "FormatCondition"
    dict.Add "AppliesTo", fc.AppliesTo.Address
    dict.Add "AppliesToLocal", fc.AppliesTo.AddressLocal
    dict.Add "Type", fc.Type
    dict.Add "Formula1", fc.Formula1
    dict.Add "Formula2", fc.Formula2
    dict.Add "NumberFormat", fc.NumberFormat
    dict.Add "StopIfTrue", fc.StopIfTrue
    dict.Add "DateOperator", fc.DateOperator
    dict.Add "Priority", fc.Priority
    dict.Add "PTCondition", fc.PTCondition
    dict.Add "ScopeType", fc.ScopeType
    dict.Add "Text", fc.Text
    dict.Add "TextOperator", fc.TextOperator
    
    dict.Add "Font", serializeFont(fc.Font)
    dict.Add "Interior", serializeInterior(fc.Interior)
    dict.Add "Borders", serializeBorders(fc.borders)
    
    Set serializeFormatCondition = dict
    
End Function

Public Function deserializeFormatCondition(ByRef dict As Dictionary, Optional Worksheet As Worksheet)

    Dim fc As FormatCondition
    Set fc = ws(Worksheet).Range(dict("AppliesTo")).FormatConditions.Add( _
            Type:=xlExpression, _
            Operator:=dict("Operator"), _
            Formula1:=dict("Formula1"), _
            Formula2:=dict("Formula2") _
            )
    
    On Error Resume Next
    With fc
        .ScopeType = dict("ScopeType")
        .DateOperator = dict("DateOperator")
        .StopIfTrue = dict("StopIfTrue")
        .Priority = dict("Priority")
        .NumberFormat = dict("NumberFormat")
        ' .PTCondition is read-only property
        If Not IsEmpty(dict("Text")) Then .Text = dict("Text")
        If Not IsEmpty(dict("TextOperator")) Then .TextOperator = dict("TextOperator")
        
        deserializeInterior dict("Interior"), .Interior
        deserializeFont dict("Font"), .Font
        deserializeBorders dict("Borders"), .borders
        
    End With

End Function


' ------------------------------------------------------------------------------------------------------------
'
' Font
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeFont(ByRef fnt As Font)

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Name", fnt.Name
    dict.Add "Background", fnt.Background
    dict.Add "Bold", fnt.Bold
    dict.Add "Color", fnt.Color
    dict.Add "ColorIndex", fnt.ColorIndex
    dict.Add "FontStyle", fnt.FontStyle
    dict.Add "Size", fnt.Size
    dict.Add "Strikethrough", fnt.Strikethrough
    dict.Add "Subscript", fnt.Subscript
    dict.Add "Superscript", fnt.Superscript
    dict.Add "ThemeColor", fnt.ThemeColor
    dict.Add "ThemeFont", fnt.ThemeFont
    dict.Add "TintAndShade", fnt.TintAndShade
    dict.Add "Underline", fnt.Underline

    Set serializeFont = dict
    
End Function

Public Function deserializeFont(ByRef dict As Dictionary, fnt As Font)

    On Error Resume Next

    With fnt
        
        .Color = dict("Color")
        .ColorIndex = dict("ColorIndex")
        .FontStyle = dict("FontStyle")
        .Italic = dict("Italic")
        .Name = dict("Name")
        .Size = dict("Size")
        .Strikethrough = dict("Strikethrough")
        .Subscript = dict("Subscript")
        .Superscript = dict("Superscript")
        .ThemeFont = dict("ThemeFont")
        .ThemeColor = dict("ThemeColor")
        .TintAndShade = dict("TintAndShade")
        .Underline = dict("Underline")
    
    End With

End Function

' ------------------------------------------------------------------------------------------------------------
'
' Interior
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeInterior(ByRef intr As Interior)

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Color", intr.Color
    dict.Add "ColorIndex", intr.ColorIndex
    ' dict.Add "Gradient", intr.Gradient
    dict.Add "InvertIfNegative", intr.InvertIfNegative
    dict.Add "Pattern", intr.Pattern
    dict.Add "PatternColor", intr.PatternColor
    dict.Add "PatternColorIndex", IfNull(intr.PatternColorIndex, xlAutomatic)
    
    dict.Add "PatternThemeColor", intr.PatternThemeColor
    dict.Add "PatternTintAndShade", intr.PatternTintAndShade
    dict.Add "ThemeColor", intr.ThemeColor
    dict.Add "TintAndShade", intr.TintAndShade

    Set serializeInterior = dict
    
End Function

Public Function deserializeInterior(ByRef dict As Dictionary, ByRef intr As Interior)

    On Error Resume Next

    Debug.Print JsonConverter.ConvertToJson(dict, " ", 2)

    With intr
    
        
        If dict("Color") <> 0 Then
            .Color = dict("Color")
            .ColorIndex = dict("ColorIndex")
            ' .Gradient = dict("Gradient")
        End If
        
        If dict("Pattern") <> 0 Then
            .Pattern = dict("Pattern")
            Select Case True
            Case .Pattern = XlPattern.xlPatternLinearGradient
            
            Case Else
                .PatternThemeColor = dict("PatternThemeColor")
                .PatternColor = dict("PatternColor")
                .PatternTintAndShade = dict("PatternTintAndShade")
            End Select
        End If
        
        If dict("ThemeColor") <> 0 Then
            .PatternColorIndex = dict("PatternColorIndex")
            .ThemeColor = dict("ThemeColor")
            .TintAndShade = CDbl(dict("TintAndShade"))
        End If
        

        .InvertIfNegative = dict("InvertIfNegative")
    
    End With

End Function

' ------------------------------------------------------------------------------------------------------------
'
' Borders
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeBorders(ByRef brdrs As borders)

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Color", brdrs.Color
    dict.Add "ColorIndex", brdrs.ColorIndex
    dict.Add "LineStyle", brdrs.LineStyle
    dict.Add "ThemeColor", brdrs.ThemeColor
    dict.Add "TintAndShade", brdrs.TintAndShade
    dict.Add "Value", brdrs.value
    dict.Add "Weight", brdrs.Weight
    
    Dim dictBorders As New Dictionary
    dictBorders.Add "xlDiagonalDown", serializeBorder(brdrs(xlDiagonalDown), xlDiagonalDown)
    dictBorders.Add "xlDiagonalUp", serializeBorder(brdrs(xlDiagonalUp), xlDiagonalUp)
    dictBorders.Add "xlEdgeBottom", serializeBorder(brdrs(xlEdgeBottom), xlEdgeBottom)
    dictBorders.Add "xlEdgeLeft", serializeBorder(brdrs(xlEdgeLeft), xlEdgeLeft)
    dictBorders.Add "xlEdgeRight", serializeBorder(brdrs(xlEdgeRight), xlEdgeRight)
    dictBorders.Add "xlEdgeTop", serializeBorder(brdrs(xlEdgeTop), xlEdgeTop)
    dictBorders.Add "xlInsideHorizontal", serializeBorder(brdrs(xlInsideHorizontal), xlInsideHorizontal)
    dictBorders.Add "xlInsideVertical", serializeBorder(brdrs(xlInsideVertical), xlInsideVertical)
    
    dict.Add "Borders", dictBorders
    
    Set serializeBorders = dict
    
End Function

Public Function deserializeBorders(ByRef dict As Dictionary, ByRef brdrs As borders)

    On Error Resume Next
    
    Dim dictBorder As Dictionary
    
    With brdrs
    
        .Color = dict("Color")
        .ColorIndex = dict("ColorIndex")
        .LineStyle = dict("LineStyle")
        .ThemeColor = dict("ThemeColor")
        .TintAndShade = dict("TintAndShade")
        .value = dict("Value")
        .Weight = dict("Weight")
    
        Set dictBorder = dict("Borders")("xlDiagonalDown"): deserializeBorder dictBorder, brdrs(xlDiagonalDown)
        Set dictBorder = dict("Borders")("xlDiagonalUp"): deserializeBorder dictBorder, brdrs(xlDiagonalUp)
        Set dictBorder = dict("Borders")("xlEdgeBottom"): deserializeBorder dictBorder, brdrs(xlEdgeBottom)
        Set dictBorder = dict("Borders")("xlEdgeLeft"): deserializeBorder dictBorder, brdrs(xlEdgeLeft)
        Set dictBorder = dict("Borders")("xlEdgeRight"): deserializeBorder dictBorder, brdrs(xlEdgeRight)
        Set dictBorder = dict("Borders")("xlEdgeTop"): deserializeBorder dictBorder, brdrs(xlEdgeTop)
        Set dictBorder = dict("Borders")("xlInsideHorizontal"): deserializeBorder dictBorder, brdrs(xlInsideHorizontal)
        Set dictBorder = dict("Borders")("xlInsideVertical"): deserializeBorder dictBorder, brdrs(xlInsideVertical)
    
    End With
    
End Function

Public Function serializeBorder(ByRef brdr As Border, brdrIndex As XlBordersIndex)

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    If IsNull(brdr.TintAndShade) Then
        dict.Add "Active", False
    Else
        dict.Add "Active", True
        dict.Add "BorderIndex", brdrIndex
        dict.Add "Color", brdr.Color
        dict.Add "ColorIndex", brdr.ColorIndex
        dict.Add "LineStyle", brdr.LineStyle
        dict.Add "ThemeColor", brdr.ThemeColor
        dict.Add "Weight", brdr.Weight
        dict.Add "TintAndShade", brdr.TintAndShade
    End If
    
    Set serializeBorder = dict
    
End Function

Public Function deserializeBorder(ByRef dict As Dictionary, ByRef brdr As Border)

    On Error Resume Next
    
    If dict("Active") = False Then Exit Function
    
    With brdr
        .Color = dict("Color")
        .ColorIndex = dict("ColorIndex")
        .LineStyle = dict("LineStyle")
        .ThemeColor = dict("ThemeColor")
        .TintAndShade = dict("TintAndShade")
        .Weight = dict("Weight")
    
    End With
      
End Function

' ------------------------------------------------------------------------------------------------------------
'
' Color Scale
'
' ------------------------------------------------------------------------------------------------------------

Public Function serializeColorScale(ByRef cs As ColorScale)

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Class", "ColorScale"
    dict.Add "AppliesTo", cs.AppliesTo.Address
    dict.Add "Type", cs.Type
    dict.Add "ColorScaleType", cs.ColorScaleCriteria.Count
    dict.Add "ColorScaleCriteria", serializeColorScaleCriteria(cs.ColorScaleCriteria)
    dict.Add "Formula", cs.Formula
    dict.Add "StopIfTrue", cs.StopIfTrue
    dict.Add "Priority", cs.Priority
    dict.Add "PTCondition", cs.PTCondition
    dict.Add "ScopeType", cs.ScopeType
    
    dict.Add "Font", serializeFont(cs.Font)
    dict.Add "Interior", serializeInterior(cs.Interior)
    dict.Add "Borders", serializeBorders(cs.borders)
    
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
        .NumberFormat = dict("NumberFormat")
        ' .PTCondition is read-only property
        
        deserializeFont dict("Font"), .Font
        deserializeBorders dict("Borders"), .borders
        deserializeColorScaleCriteria dict("ColorScaleCriteria"), .ColorScaleCriteria, dict("ColorScaleType")

    End With
    
End Function

Public Function serializeColorScaleCriteria(ByRef cscriteria As ColorScaleCriteria)

    Dim dict As New Dictionary
    Dim cscriterion As ColorScaleCriterion
    Dim idx As Integer: idx = 0
    
    On Error Resume Next
    For Each cscriterion In cscriteria
        idx = idx + 1
        dict.Add Format(idx, "000"), serializeColorScaleCriterion(cscriterion)
    Next

    Set serializeColorScaleCriteria = dict
    
End Function

Public Function deserializeColorScaleCriteria(ByRef dict As Dictionary, ByRef csa As ColorScaleCriteria, ByVal ColorScaleType)

    On Error Resume Next
    
    Dim dictBorder As Dictionary
    
    With csa
        deserializeColorScaleCriterion dict("001"), .Item(1)
        deserializeColorScaleCriterion dict("002"), .Item(2)
        If ColorScaleType = 3 Then
            deserializeColorScaleCriterion dict("003"), .Item(3)
        End If
    
    End With
    
End Function

Public Function serializeColorScaleCriterion(ByRef cscriterion As ColorScaleCriterion)

    Dim dict As New Dictionary
    
    On Error Resume Next
    dict.Add "Index", cscriterion.Index
    dict.Add "Type", cscriterion.Type
    dict.Add "Value", cscriterion.value
    dict.Add "FormatColor", serializeFormatColor(cscriterion.FormatColor)
    
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

Public Function serializeFormatColor(ByRef fmtclr As FormatColor) As Dictionary

    Dim dict As New Dictionary
    
    On Error Resume Next
    dict.Add "Class", "FormatColor"
    dict.Add "Color", fmtclr.Color
    dict.Add "ColorIndex", fmtclr.ColorIndex
    dict.Add "ThemeColor", fmtclr.ThemeColor
    dict.Add "TintAndShade", fmtclr.TintAndShade
    
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

Public Function serializeIconSetCondition(ByRef isc As IconSetCondition) As Dictionary

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Class", "IconSetCondition"
    dict.Add "AppliesTo", isc.AppliesTo.Address
    dict.Add "AppliesToLocal", isc.AppliesTo.AddressLocal
    dict.Add "Type", isc.Type
    dict.Add "Formula", isc.Formula
    
    dict.Add "PercentileValues", isc.PercentileValues
    dict.Add "Priority", isc.Priority
    dict.Add "ReverseOrder", isc.ReverseOrder
    dict.Add "ScopeType", isc.ScopeType
    dict.Add "ShowIconOnly", isc.ShowIconOnly
    dict.Add "StopIfTrue", isc.StopIfTrue
    
    dict.Add "IconSet", isc.IconSet.ID
    dict.Add "IconCriteria", serializeIconCriteria(isc.IconCriteria)
    
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

Public Function serializeIconCriteria(ByRef ic As IconCriteria)

    Dim dict As New Dictionary
    Dim icn As IconCriterion
    
    On Error Resume Next
    For Each icn In ic
        dict.Add Format(icn.Index, "000"), serializeIconCriterion(icn)
    Next
    Set serializeIconCriteria = dict
    
End Function

Public Function deserializeIconCriteria(dict As Dictionary, ica As IconCriteria)

    Dim key
    Dim icn As Icon
    For Each key In dict.Keys
        deserializeIconCriterion dict(key), ica(Int(key))
    Next
End Function

Public Function serializeIconCriterion(ByRef icn As IconCriterion)

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Icon", icn.Icon
    dict.Add "Operator", icn.Operator
    dict.Add "Type", icn.Type
    dict.Add "Value", icn.value
    
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

Public Function serializeDatabar(ByRef dbar As Databar) As Dictionary

    Dim dict As New Dictionary
    
    On Error Resume Next
    
    dict.Add "Class", "Databar"
    dict.Add "AppliesTo", dbar.AppliesTo.Address
    dict.Add "AppliesToLocal", dbar.AppliesTo.AddressLocal
    dict.Add "Type", dbar.Type
    dict.Add "AxisColor", serializeFormatColor(dbar.AxisColor)
    dict.Add "AxisPosition", IfNull(dbar.AxisPosition, xlDataBarAxisAutomatic)
    dict.Add "BarBorder", serializeDataBarBorder(dbar.BarBorder)
    dict.Add "BarColor", serializeFormatColor(dbar.BarColor)
    dict.Add "BarFillType", dbar.BarFillType
    dict.Add "Direction", dbar.Direction
    dict.Add "Formula", dbar.Formula
    dict.Add "MaxPoint", serializeConditionValue(dbar.MaxPoint)
    dict.Add "MinPoint", serializeConditionValue(dbar.MinPoint)
    dict.Add "NegativeBarFormat", serializeNegativeBarFormat(dbar.NegativeBarFormat)
    dict.Add "PercentMax", dbar.PercentMax
    dict.Add "PercentMin", dbar.PercentMin
    dict.Add "Priority", dbar.Priority
    dict.Add "ScopeType", dbar.ScopeType
    dict.Add "ShowValue", dbar.ShowValue
    dict.Add "StopIfTrue", dbar.StopIfTrue
    dict.Add "Type", dbar.Type
    
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

Public Function serializeConditionValue(ByRef cv As ConditionValue) As Dictionary
    
    Dim dict As New Dictionary
    
    On Error Resume Next
    
    With cv
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


Public Function serializeNegativeBarFormat(ByRef nbf As NegativeBarFormat) As Dictionary
    
    Dim dict As New Dictionary
    
    On Error Resume Next
    
    With nbf
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

Public Function serializeDataBarBorder(ByRef dbrd As DataBarBorder) As Dictionary

    Dim dict As New Dictionary

    On Error Resume Next
    With dbrd
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
