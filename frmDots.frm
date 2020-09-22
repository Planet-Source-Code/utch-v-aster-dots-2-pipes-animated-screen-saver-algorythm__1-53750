VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrCircle 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2610
      Top             =   3630
   End
   Begin VB.PictureBox pScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   390
      ScaleHeight     =   735
      ScaleWidth      =   1125
      TabIndex        =   0
      Top             =   600
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim C As Integer
Dim Stage As Integer
Dim nCircles As Integer

Private Sub Form_Load()
    
    Stage = c_LOADING
    TmrCircle.Enabled = True
    nCircles = 250
    
End Sub

Private Sub Form_Resize()
    
    pScreen.Move 0, 0, ScaleWidth, ScaleHeight
    
End Sub

Private Sub pScreen_Click()
    End
End Sub

Private Sub TmrCircle_Timer()

    If Stage = c_LOADING Then
        Dim X As Long
        Dim Y As Long
        Dim R As Single
        Dim Color As Long
    
        R = Rand(400, 400)
        
        X = Rand(0, ScaleWidth)
        Y = Rand(0, ScaleHeight)
        
        Color = RandomColor()
        Call DrawCircle(X, Y, R, Color)
    
        Circles(C).Color = Color
        
10:
        Circles(C).Radius = R
        
        If R < 0.5 Then GoTo 10
        
        Circles(C).X = X
        Circles(C).Y = Y
        Circles(C).Speed = Rand(25, 25)
        Circles(C).XSlope = Rand(1, 5)
        Circles(C).YSlope = Rand(1, 5)
        Circles(C).Dead = False
        
        If Rand(1, 2) Mod 2 = 1 Then Circles(C).XSlope = Circles(C).XSlope * -1
        If Rand(1, 2) Mod 2 = 1 Then Circles(C).YSlope = Circles(C).YSlope * -1
    
        C = C + 1
        
        If C = nCircles Then
            
            Call Pause(1)
            Stage = c_MOVING
            C = 0
            
        End If
    Else
        
        Dim OnScreen As Boolean
        Static DeadCount As Integer
        
        'pScreen.Cls
    
        For C = 0 To nCircles
            
            OnScreen = True
            
            If Not (Circles(C).Dead) Then
                
                Circles(C).X = (Circles(C).X + Circles(C).XSlope * Circles(C).Speed)
                Circles(C).Y = (Circles(C).Y + Circles(C).YSlope * Circles(C).Speed)
                
                If Circles(C).X <= Circles(C).Radius * -1 Then OnScreen = False
                If Circles(C).X >= ScaleWidth + Circles(C).Radius Then OnScreen = False
                If Circles(C).Y <= Circles(C).Radius * -1 Then OnScreen = False
                If Circles(C).Y >= ScaleHeight + Circles(C).Radius Then OnScreen = False
                
                If Circles(C).Dead = False Then
                    If Not OnScreen Then
                        DeadCount = DeadCount + 1
                        Circles(C).Dead = True
                    End If
                End If
                
                
                Call DrawCircle(Circles(C).X, Circles(C).Y, Circles(C).Radius, Circles(C).Color)
            
            End If
            
        Next
        
        If DeadCount = nCircles Then
            
            Call Pause(1)
            Stage = c_LOADING
            C = 0
            DeadCount = 0
            
        End If
    
    End If
End Sub

Sub DrawCircle(X As Long, Y As Long, R As Single, Color As Long)
    
    If X = 0 Or Y = 0 Or R = 0 Then Exit Sub
    
    pScreen.FillStyle = 0
    pScreen.FillColor = Color
    pScreen.Circle (X, Y), R
    
    pScreen.FillStyle = 1
    pScreen.Circle (X, Y), R - 0.1, vbBlack 'pScreen.BackColor
    
End Sub
