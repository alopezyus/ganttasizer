Attribute VB_Name = "AppObject"
Private AppE As AppEvents

Public Sub StartEvents()
    Set AppE = New AppEvents
    
    'Initiate Date Picker
    ensureDPManager
    LoadGlobalSettings
End Sub

Public Sub StopEvents()
    Set AppE = Nothing
End Sub

