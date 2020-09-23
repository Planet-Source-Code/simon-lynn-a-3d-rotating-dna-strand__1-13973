VERSION 5.00
Begin VB.Form frmDNA 
   BackColor       =   &H00000000&
   Caption         =   "3D DNA"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3075
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   3075
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "frmDNA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DEG = 1.74532925199433E-02  ' Constant for converting degrees to radians

Private Finish As Boolean

Private Sub Form_Load()

    Show
    Rotate  ' Do the main loop

End Sub

Private Sub Delay(Amount)

' This sub creates a delay based on the Timer function.
' I find it more reliable than the timer control,
' and it also allows for smaller intervals.

Dim Start!

    Start = Timer   ' Start = number of seconds since midnight
    Do Until Timer >= Start + Amount ' Stop when amount has elapsed
        DoEvents
    Loop

End Sub

Private Sub Orbit(Angle!, CntX!, CntY!, Radius!, Ratio!, Colour&)

' This plots two points on the circumference of the circle or
' ellipse (depending on the ratio setting), one at the specified
' angle, the other 180 degrees from the first, and draws a line
' between them, of specified colour.
' The ratio argument shortens the circle in the Y direction, so that
' Y Radius = Radius / Ratio.

Dim PntX1!, PntY1!, PntX2!, PntY2!

    PntX1 = CntX - (Sin(Angle * DEG) * Radius)
    PntY1 = CntY - (Cos(Angle * DEG) * (Radius / Ratio))
    PntX2 = CntX - (Sin((Angle + 180) * DEG) * Radius)
    PntY2 = CntY - (Cos((Angle + 180) * DEG) * (Radius / Ratio))
        
    PSet (PntX1, PntY1), vbRed
    PSet (PntX2, PntY2), vbBlue
    
    DrawWidth = 1
    Line (PntX1, PntY1)-(PntX2, PntY2), Colour ' Draws a line between the two points
    DrawWidth = 3

End Sub

Private Sub Rotate()

' Here's a brief explanation of how this works:
' The orbit sub-routine is used to create elliptical orbits
' with two points going round and a line in between. The
' strand is simply an number of these orbits, but each one
' is offset by a few degrees to the last, and a slightly
' further down the form.

Static Ang
Dim i%, Ratio!, Rungs%, Step!, Gap!

    Do
        DoEvents
    
        Cls
        
        Ang = Ang + 10
        
    '##################################################
    ' Change these variables to change the appearance of the strand:
        
        Ratio = -3      ' Causes a change in the apparent angle of the strand.
        Rungs = 25      ' How many sets of points to have
        Step = 8        ' How many degrees to increase per rung
        Gap = 120       ' The distance between rungs
        
    '##################################################
          
        For i = 0 To Rungs - 1
            Orbit Ang + (i * Step), Width / 2, 600 + (i * Gap), _
                500, Ratio, RGB(255 - (i * (255 / Rungs)), 0, i * (255 / Rungs))
        Next
        
        Delay 0.01      ' Slow down the loop
        
    Loop Until Finish
    
    End
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Finish = True ' Stops the loop before ending

End Sub
