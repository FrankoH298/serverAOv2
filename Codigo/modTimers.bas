Attribute VB_Name = "modTimers"
Option Explicit
Type tMainLoop
    MAXINT As Long
    LastCheck As Long
End Type
Private Const NumTimers As Byte = 5 '//Aca la cantidad de timers.

Private MainLoops(1 To NumTimers) As tMainLoop

Private Enum eTimers
    GameTimer = 1
    packetResend = 2
    TIMER_AI = 3
    Auditoria = 4
    eTimerFlush = 5
End Enum

Public prgRun As Boolean


Public Sub MainLoop()
    Dim LoopC As Long
    MainLoops(eTimers.GameTimer).MAXINT = 40
    MainLoops(eTimers.packetResend).MAXINT = 10
    MainLoops(eTimers.TIMER_AI).MAXINT = 380
    MainLoops(eTimers.Auditoria).MAXINT = 1000
    MainLoops(eTimers.eTimerFlush).MAXINT = 12
    
    prgRun = True
    
    Do While prgRun
        For LoopC = 1 To NumTimers
            With MainLoops(LoopC)
                If timeGetTime > .LastCheck Then
                    Call MakeProcces(LoopC)
                End If
            End With
            DoEvents
        Next LoopC
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
            
        Case eTimers.eTimerFlush
            Call TimerFlush
    End Select
    
    MainLoops(index).LastCheck = timeGetTime + MainLoops(index).MAXINT
End Sub
Private Sub TimerFlush()
    Dim i As Long
    For i = 1 To MaxUsers

        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then

                Dim Ret As Long

                Ret = WsApiEnviar(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))

                If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
                    ' Close the socket avoiding any critical error
                    Call CloseSocketSL(i)
                    Call Cerrar_Usuario(i)
                End If
            End If
        End If

    Next i
    DoEvents
End Sub
