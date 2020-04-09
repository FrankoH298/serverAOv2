Attribute VB_Name = "modTimers"
Option Explicit
Type tMainLoop
    MAXINT As Long
    LastCheck As Long
End Type

Private Const NumTimers As Long = 4 '//Aca la cantidad de timers.

Private MainLoops(1 To NumTimers) As tMainLoop

Private Enum eTimers
    GameTimer = 1
    packetResend = 2
    TIMER_AI = 3
    Auditoria = 4
End Enum

Public prgRun As Boolean


Public Sub MainLoop()
    Dim loopC As Long
    MainLoops(eTimers.GameTimer).MAXINT = 40
    MainLoops(eTimers.packetResend).MAXINT = 10
    MainLoops(eTimers.TIMER_AI).MAXINT = 380
    MainLoops(eTimers.Auditoria).MAXINT = 1000
    
    prgRun = True
    
    Do While prgRun
        For loopC = 1 To NumTimers
            With MainLoops(loopC)
                If timeGetTime > .LastCheck Then
                    Call MakeProcces(loopC)
                End If
            End With
            DoEvents
        Next loopC
        DoEvents
    Loop
End Sub

Private Sub MakeProcces(ByVal index As Integer)
    Select Case index
    
        Case eTimers.GameTimer
            Call frmMain.GameTimer_Timer

        Case eTimers.packetResend
            Call frmMain.packetResend_Timer
            
        Case eTimers.TIMER_AI
            Call frmMain.TIMER_AI_Timer
            
        Case eTimers.Auditoria
            Call frmMain.Auditoria_Timer
            
    End Select
    
    MainLoops(index).LastCheck = timeGetTime + MainLoops(index).MAXINT
End Sub
