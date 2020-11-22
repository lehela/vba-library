Attribute VB_Name = "KeyValueStore"
Option Explicit
    
Const loName = "KeyValueStore"

  
Private Sub Example()

    Const key As String = "CF1"
    Dim val As String, res As String
    
    ' 16-bit strings are no problem, even more are ok
    val = Replace(Space(2 ^ 16), " ", "A")
    
    ' Verify the length to console
    Debug.Print Len(val)

    ' Store in Custom Property
    KeyValueStore.SetValue key, val
    
    ' Verify that key exists now
    Debug.Print KeyValueStore.KeyExists(key)
    
    ' Read back the key and verify the length is still the same
    res = KeyValueStore.GetValue(key)
    Debug.Print Len(res)

    ' Delete the key from the CustomProperties
    KeyValueStore.DeleteKey key

    ' Verify that the key is deleted
    Debug.Print KeyValueStore.KeyExists(key)

    
End Sub

' ------------------------------------------------------------------------------------------------------------
' Public Subs
' ------------------------------------------------------------------------------------------------------------

Public Sub KeyValueStore_Init()
    KeyValueStore.Init
End Sub

Public Sub KeyValueStore_ImportFromTable()
    
    If loKeyValueStore Is Nothing Then
        MsgBox "Worksheet does not have a KeyValueStore table", vbCritical
        Exit Sub
    End If
    
    Dim lr As ListRow
    Dim lcs As ListColumns
    Set lcs = loKeyValueStore.ListColumns
    
    Dim key As String, value As String
    
    For Each lr In loKeyValueStore.ListRows
        key = lr.Range(lcs("Key").Index)
        value = lr.Range(lcs("Value").Index)
        If Trim(key) <> "" Then
            SetValue key, value
        End If
    Next
    
    MsgBox "Appended", vbOKOnly

End Sub


Public Sub KeyValueStore_ExportToTable()
    
    Dim kv As CustomProperty
    Dim lo As ListObject
     
    ' Delete KeyValueStore table if exists
    On Error Resume Next: loKeyValueStore.Delete: On Error GoTo 0
    
    ' Delete a selected table if exists
    If Not Selection.ListObject Is Nothing Then
        If MsgBox("Delete currently selected table?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
        Selection.ListObject.Delete
    End If
    
    ' Create KeyValueStore table
    Set lo = ws.ListObjects.Add
    
    lo.Name = loName
    lo.ListColumns.Add
    lo.HeaderRowRange.Cells(1) = "Key"
    lo.HeaderRowRange.Cells(2) = "Value"
    
    ' Populate table
    For Each kv In ws.CustomProperties
        
        With lo.ListRows.Add
            .Range(1, 1) = kv.Name
            .Range(1, 2) = kv.value
        End With
    Next
    
End Sub


' ------------------------------------------------------------------------------------------------------------
' Public Functions
' ------------------------------------------------------------------------------------------------------------

Public Function Init(Optional Worksheet As Worksheet)

    With ws(Worksheet).CustomProperties
        While .Count > 0
            .Item(1).Delete
        Wend
    End With

End Function

Public Function SetValue(ByVal key As String, ByVal value As String, Optional Worksheet As Worksheet)
    
    If KeyExists(key, Worksheet) Then
        KeyValueStore.DeleteKey key, Worksheet
    End If
    
    With ws(Worksheet).CustomProperties
            .Add key, value
    End With

End Function

Public Function DeleteKey(ByVal key As String, Optional Worksheet As Worksheet)
    
    Dim KeyIndex
    KeyIndex = GetIndex(key, Worksheet)
    
    With ws(Worksheet).CustomProperties
    
        If KeyIndex > 0 Then
            .Item(KeyIndex).Delete
        End If
    
    End With
    
End Function

Public Function GetValue(ByVal key As String, Optional Worksheet As Worksheet) As String
    
    GetValue = ""
    
    On Error Resume Next
    GetValue = ws(Worksheet).CustomProperties.Item(GetIndex(key, Worksheet))

End Function


Public Function Keys(Optional Worksheet As Worksheet) As Collection
    Set Keys = New Collection
    Dim kv As CustomProperty
    
    For Each kv In ws(Worksheet).CustomProperties
        Keys.Add kv.Name
    Next
    
End Function

Public Function KeyExists(ByVal key As String, Optional Worksheet As Worksheet) As Boolean
    KeyExists = GetIndex(key, Worksheet) > 0
End Function

Private Function GetIndex(ByVal key As String, Optional Worksheet As Worksheet) As Long

    Dim found As Boolean
    Dim kv As CustomProperty
    
    found = False
    GetIndex = 0
    For Each kv In ws(Worksheet).CustomProperties
        GetIndex = GetIndex + 1
        If kv.Name = key Then
            found = True
            Exit For
        End If
    Next
    
    If Not found Then GetIndex = 0
    
End Function

Private Property Get ws(Optional Worksheet As Worksheet) As Worksheet

    If Worksheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = Worksheet
    End If
    
End Property

Private Property Get loKeyValueStore() As ListObject

    On Error Resume Next
    Set loKeyValueStore = ws.ListObjects(loName)
    On Error GoTo 0

End Property
