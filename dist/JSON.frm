VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JSON 
   Caption         =   "UserForm1"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   OleObjectBlob   =   "JSON.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private jsonData As Object ' object to store JSON data
Private columnName As String
Private tableName As String

'
Private Sub UserForm_Initialize()

    ' setup workflow
    columnName = getColumnName(ActiveCell)
    tableName = Split(columnName, ":")(0)
    
    Dim tableObject As ListObject
    Dim tableRange As Range
    Set tableObject = getTableByName(tableName)
    Set tableRange = tableObject.ListColumns(3).DataBodyRange
    
    
    For Each Cell In tableRange
        If Not IsEmpty(Cell.value) Then
            signatureValue.AddItem Cell.value
        End If
    Next Cell
    
    Set tableRange = tableObject.ListColumns(1).DataBodyRange
    
    For Each Cell In tableRange
        If Not IsEmpty(Cell.value) Then
        
            idValue.AddItem Cell.value
        End If
    Next Cell
    reset
    ' MsgBox "Форма загружена!", vbInformation, tableName
    
End Sub

Private Sub btnSave_Click()
    ActiveCell.value = ConvertToJson(jsonData)
End Sub

Private Sub btnReset_Click()
    reset
End Sub

' reset current data in form to the default value
Private Function reset()

    ' setup workflow
    Dim jsonContent As String
    Dim jsonParser As Object

    ' parse data
    jsonContent = ActiveCell.value
    Set jsonData = ValidatorJson.ParseJson(jsonContent)

    ' show props in list box
    UpdateListBox
    
End Function

' display the value corresponding to the selected key
Private Sub lstKeys_Change()
    On Error Resume Next
    If lstKeys.ListIndex <> -1 Then
        If Not IsNull(jsonData) And Not IsEmpty(jsonData) Then
            idValue.value = jsonData(lstKeys.value)(1)
            textValue.Text = jsonData(lstKeys.value)(2)
        Else
            textValue.Text = "No data found"
        End If
    End If
    On Error GoTo 0
End Sub

' update the id for the signature
Private Sub idValue_Change()
    signatureValue.ListIndex = idValue.ListIndex
End Sub

' update the signature for the id
Private Sub signatureValue_Change()
    idValue.ListIndex = signatureValue.ListIndex
    UpdateCollectionItem
End Sub

' update the value for the current key
Private Sub textValue_Change()
    UpdateCollectionItem
End Sub

'
Private Sub UpdateListBox()
    Dim key As Variant
    lstKeys.Clear
    For Each key In jsonData.Keys
        lstKeys.AddItem key
    Next key
    lstKeys.ListIndex = 0
End Sub

'
Private Function UpdateCollectionItem()
    If lstKeys.ListIndex <> -1 Then
        Set col = New collection
        col.Add ChooseValue(idValue.value, jsonData(lstKeys.value)(1))
        col.Add ChooseValue(textValue.value, jsonData(lstKeys.value)(2))
        Set jsonData(lstKeys.value) = col
    End If
End Function

' select the appropriate value from the two
Private Function ChooseValue(value_1, value_2) As Variant

    Dim isIdEmpty As Boolean
    isIdEmpty = Trim(value_1) = ""
    If isIdEmpty Then
        ChooseValue = value_2
    Else
        ChooseValue = value_1
    End If
End Function
