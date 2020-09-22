Attribute VB_Name = "Module1"
Option Explicit

Public Const c_LOADING = 0
Public Const c_MOVING = 1

Public Type jCircle
    Color As Long
    Dead As Boolean
    Radius As Single
    Speed As Integer
    X As Long
    XSlope As Long
    Y As Long
    YSlope As Long
End Type
Public Circles(0 To 999) As jCircle

Public Declare Function GetTickCount Lib "kernel32" () As Long

Function InDesignMode() As Boolean
    
    On Error GoTo err
    
    Debug.Print 1 / 0
    
    InDesignMode = False
    
    Exit Function

err:
    
    InDesignMode = True
    
End Function

Public Function Pause(Value As Long)
    
    Value = Value * 1000
    
    Dim PreTick As Long
    PreTick = GetTickCount

    Do Until GetTickCount() >= PreTick + Value
        DoEvents
    Loop

End Function

Public Function Rand(Min As Integer, Max As Integer) As Integer
10:
    Rand = Int((Rnd * Max) + Min)
    If Rand < Min Or Rand > Max Then GoTo 10
End Function

Public Function RandomColor()
    
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    
    Randomize
    
    Red = Rand(1, 255)
    Green = Rand(1, 255)
    Blue = Rand(1, 255)

    RandomColor = RGB(Red, Green, Blue)

End Function
