Attribute VB_Name = "modTimers"
Option Explicit
Type tMainLoop
    MaxInt As Long
    LastCheck As Long
End Type
Private Const NumTimers As Byte = 2 '//Aca la cantidad de timers.

Private MainLoops(1 To NumTimers) As tMainLoop

Private Enum eTimers
    GameTimer
    packetResend
End Enum

Public Sub MainLoop()
    Dim LoopC As Long
    MainLoops(eTimers.GameTimer).MaxInt = 40
    MainLoops(eTimers.packetResend).MaxInt = 10

    Do While (frmMain.Visible) Or (frmCargando.Visible)
        For LoopC = 1 To NumTimers
            With MainLoops(LoopC)
                If GetTickCount - .LastCheck >= .MaxInt Then
                    Call MakeProcces(LoopC)
                End If
            End With
            DoEvents
        Next LoopC
        DoEvents
    Loop
End Sub

Private Sub MakeProcces(ByVal Index As Integer)
    Select Case Index
        Case eTimers.GameTimer
            Call frmMain.GameTimer_Timer

        Case eTimers.packetResend
            Call frmMain.packetResend_Timer
    End Select
    MainLoops(Index).LastCheck = GetTickCount
End Sub
