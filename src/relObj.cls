VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "relObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

'Variables privadas para uso interno de la clase
Private strActID As String
Private lngArrID As Long
Private strRelType As String
Private dblLag As Double

'Devuelve valores de las propiedades de la clase
Property Get ActID() As String
    ActID = strActID
End Property

Property Get ArrID() As Long
    ArrID = lngArrID
End Property

Property Get RelType() As String
    RelType = strRelType
End Property

Property Get Lag() As Double
    Lag = dblLag
End Property

'Escribe valores de las propiedades de la clase
Property Let ActID(ByVal value As String)
    strActID = value
End Property

Property Let ArrID(ByVal value As Long)
    lngArrID = value
End Property

Property Let RelType(ByVal value As String)
    strRelType = value
End Property

Property Let Lag(ByVal value As Double)
    dblLag = value
End Property
