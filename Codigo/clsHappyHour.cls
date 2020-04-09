VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHappyHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
'Component        : clsHappyHour
'Author           : FrankoH298
'Description      : Clase encargada de gestionar los mapas con HappyHour activo.
'Last Modification: 09/04/2020
'***************************************************************

Private Const multiplierEXP As Long = 2
Private Const multiplierGLD As Long = 2

Public mapHappyHourActivated As Long

Private mapasDisponibles() As Long

Public Sub Initialize()
'***************************************************************
'Author           : FrankoH298
'Description      : Cargamos los mapas disponibles para el happyHour
'Last Modification: 09/04/2020
'***************************************************************
    Dim MapasIniManager As clsIniManager
    Dim CantidadMapas As Long
    Dim loopC As Long
    
    Call MapasIniManager.Initialize(App.Path & "\Dat\HappyHMaps.ini")
    
    CantidadMapas = MapasIniManager.GetValue("CONFIG", "CANTIDAD")
    ReDim mapasDisponibles(1 To CantidadMapas) As Long
    
    For loopC = 1 To CantidadMapas
        mapasDisponibles(loopC) = MapasIniManager.GetValue("MAPAS", "MAPA" & loopC)
    Next loopC
    
    MapasIniManager = Nothing
    
End Sub

Public Sub ToggleHappyHour(ByVal Random As Boolean)
    If mapHappyHourActivated = 0 Then
        If RandomNumber(0, 4) = 1 And Random Then
            mapHappyHourActivated = mapasDisponibles(RandomNumber(LBound(mapasDisponibles), UBound(mapasDisponibles)))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> HappyHour iniciado en el mapa " + CStr(mapHappyHourActivated), FontTypeNames.FONTTYPE_GM))
        End If
    Else
        mapHappyHourActivated = 0
    End If
End Sub
Public Property Get getMultiplierEXP() As Long
    getMultiplierEXP = multiplierEXP
End Property

Public Property Get getMultiplierGLD() As Long
    getMultiplierGLD = multiplierGLD
End Property
