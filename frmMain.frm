VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " Let's Talk About SPEED !"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrTenSecondsOnly 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   285
      Top             =   465
   End
   Begin VB.Menu mnuFile 
      Caption         =   " &File"
      Begin VB.Menu mnuExit 
         Caption         =   " E&xit"
      End
   End
   Begin VB.Menu mnuFormSize 
      Caption         =   " Form &Size"
      Begin VB.Menu mnuSize 
         Caption         =   " 1 - 640 x 480"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   " 2 - 800 x 600"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuSize 
         Caption         =   " 3 - 1024 x 768"
         Index           =   3
      End
      Begin VB.Menu mnuSize 
         Caption         =   " 4 - 1152 x 864"
         Index           =   4
      End
      Begin VB.Menu mnuSize 
         Caption         =   " 5 - 1280 x 1024"
         Index           =   5
      End
      Begin VB.Menu mnuSize 
         Caption         =   " 6 - 1600 x 1200"
         Index           =   6
      End
   End
   Begin VB.Menu mnuIterationDeep 
      Caption         =   " &Iterations"
      Begin VB.Menu mnuIterations 
         Caption         =   " 1 - 3"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuIterations 
         Caption         =   " 2 - 4"
         Index           =   2
      End
      Begin VB.Menu mnuIterations 
         Caption         =   " 3 - 5"
         Index           =   3
      End
      Begin VB.Menu mnuIterations 
         Caption         =   " 4 - 6"
         Index           =   4
      End
      Begin VB.Menu mnuIterations 
         Caption         =   " 5 - 7"
         Index           =   5
      End
      Begin VB.Menu mnuIterations 
         Caption         =   " 6 - 8"
         Index           =   6
      End
      Begin VB.Menu mnuIterations 
         Caption         =   " 7 - 9"
         Index           =   7
      End
   End
   Begin VB.Menu mnuStart 
      Caption         =   " &Go for 10 seconds !"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   " ?"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmMain.frm
'

'   Last update 02/01/2005  Light Templer / Carles P.V.


' This form is just a demonstration for the fast gradient sub.
' The recursive called sub RecursiveDrawGradientAndDivideRectangle()
' does all the work. By subdividing the forms client area in four
' rectangles and drawing a top-down gradient on each with randomly
' generated colors we get a mix between larger and smaller gradients.
' After ten seconds the timer stops the algorithm to see the number
' of so far drawn gradients.

' The magic is in the bas modul:  The sub 'DrawTopDownGradient()'  is a
' piece of code carefully optimized for speed by the PSC comunity, but
' mainly by Carles P.V. .
' Please have a closer look at it - its well worth!

' NOTE: This gradient sub should run without any problems on every
' Windows version from Win 95 upto Windows XP SP2 without any GDI
' memory leaks.

' btw:  'showRGBgradientsInitScreen()'  and  'SetNewMenuCheck()' are nice ones - have a look ... ;)

' Any comments / suggestions to improvements?
' Please send me an email:  schwepps_bitterlemon@gmx.de



Option Explicit


' *************************************
' *        API DECLARES               *
' *************************************
Private Declare Function API_GetClientRect Lib "user32" Alias "GetClientRect" _
        (ByVal hwnd As Long, _
         lpRect As tpAPI_RECT) As Long


' *************************************
' *        PRIVATES                   *
' *************************************
Private flgStopAll      As Boolean
Private lIterationDeep  As Long
Private lIterations     As Long
Private lGradients      As Long
'
'
'


Private Sub mnuStart_Click()
    ' Start the output
    
    Dim FormArea    As tpAPI_RECT
        
    
    ' Don't start it twice or change anything ;)
    mnuStart.Enabled = False
    mnuFormSize.Enabled = False
    mnuIterationDeep.Enabled = False
    
    ' Get client area of form in pixel
    API_GetClientRect Me.hwnd, FormArea
    With FormArea
        .lRight = .lRight - .lLeft
        .lLeft = 0
        .lBottom = .lBottom - .lTop
        .lTop = 0
    End With
    
    ' Reset all
    flgStopAll = False
    lGradients = 0
    tmrTenSecondsOnly.Enabled = True
    
    ' Start the main loop
    Do While flgStopAll = False
        RecursiveDrawGradientAndDivideRectangle FormArea
    Loop
    
    ' Give us the result
    MsgBox " Ok, just drawn " & Format$(lGradients, "#,###") & " gradients in 10 seconds !", _
            vbExclamation + vbMsgBoxSetForeground, " Ready"
    
    ' Make it restartable
    mnuStart.Enabled = True
    mnuFormSize.Enabled = True
    mnuIterationDeep.Enabled = True
    
    showRGBgradientsInitScreen
    
End Sub


Private Sub RecursiveDrawGradientAndDivideRectangle(DrawingArea As tpAPI_RECT)
    ' Main worker sub - the name speaks for itself ;)
    
    Dim SubArea As tpAPI_RECT
    
    ' Ten seconds over? Stop!
    If flgStopAll = True Then
    
        Exit Sub
    End If
    
    ' Increment iteration level
    lIterations = lIterations + 1
    
    ' Draw the gradient with random colors
    DrawTopDownGradient Me.hdc, DrawingArea, RGB(256 * Rnd, 256 * Rnd, 256 * Rnd), RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
    lGradients = lGradients + 1
    DoEvents
        
    ' Stop when selected number of iterations are reached
    If lIterations = lIterationDeep Then
        lIterations = lIterations - 1
        
        Exit Sub
    End If
    
    ' Divide area in four parts with same size and go in recursion for all of them
    With SubArea
        
        ' Left Top
        .lLeft = DrawingArea.lLeft
        .lTop = DrawingArea.lTop
        .lRight = .lLeft + (DrawingArea.lRight - DrawingArea.lLeft) / 2
        .lBottom = .lTop + (DrawingArea.lBottom - DrawingArea.lTop) / 2
        RecursiveDrawGradientAndDivideRectangle SubArea
        
        ' Right Top
        .lLeft = .lRight
        .lRight = DrawingArea.lRight
        RecursiveDrawGradientAndDivideRectangle SubArea
        
        ' Left Bottom
        .lLeft = DrawingArea.lLeft
        .lTop = .lBottom - 1
        .lRight = .lLeft + (DrawingArea.lRight - DrawingArea.lLeft) / 2
        .lBottom = DrawingArea.lBottom
        RecursiveDrawGradientAndDivideRectangle SubArea
        
        ' Right Bottom
        .lLeft = .lRight
        .lRight = DrawingArea.lRight
        RecursiveDrawGradientAndDivideRectangle SubArea
        
    End With
    
    ' We leave this iteration deepness level so decrement
    lIterations = lIterations - 1
    
End Sub


Private Sub Form_Load()
    
    ' Setting defaults
    mnuSize_Click 2                 ' 800 x 600 form size
    mnuIterations_Click 4           ' 7 iteration levels deep
    
    Me.Show
    DoEvents
    showRGBgradientsInitScreen      ' To proof the the proper colors (one olderversion twisted the R,G,B ... ;)

    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    End
    
End Sub


Private Sub showRGBgradientsInitScreen()
    ' Just an init screen to proof correct color handling.
    ' Look at it only if you like this old programmers game: What can be done with very less (lines) of tricky code ... ;)
    ' (Be aware: This game is NOT suitable for good, readable and fast code ;))) ! )

    Dim i As Long, TheRect As tpAPI_RECT, FormArea As tpAPI_RECT
    
    FormArea = GetFormArea()
    DrawTopDownGradient Me.hdc, FormArea, 0, RGB(255, 255, 255)
    
    TheRect.lTop = FormArea.lRight / 16: TheRect.lBottom = FormArea.lBottom - TheRect.lTop
    For i = 0 To 2
        TheRect.lLeft = TheRect.lTop + i * (TheRect.lTop * 5)
        TheRect.lRight = TheRect.lLeft + (TheRect.lTop * 4)
        DrawTopDownGradient Me.hdc, TheRect, RGB(255, 255, 255), RGB(-255 * (i = 0), -255 * (i = 1), -255 * (i = 2))
    Next i

End Sub

Private Sub tmrTenSecondsOnly_Timer()
    
    ' After 10 seconds we stop running
    flgStopAll = True
    tmrTenSecondsOnly.Enabled = False

End Sub


Private Sub mnuSize_Click(Index As Integer)
    
    With Me
        ' Set new size (shrinked to 90% of screen width to fit into screen)
        .Width = .ScaleX(Choose(Index, 640, 800, 1024, 1152, 1200, 1600) * 0.9, vbPixels, vbTwips)
        .Height = .ScaleY(Choose(Index, 480, 600, 768, 864, 1024, 1200) * 0.9, vbPixels, vbTwips)
        
        ' Center form on screen
        .Left = (Screen.Width / 2) - (.Width / 2)
        .Top = (Screen.Height / 2) - (.Height / 2)
    End With
    
    ' Set new checkmark
    SetNewMenuCheck mnuSize, Index
    
    Cls
    showRGBgradientsInitScreen
    
End Sub


Private Sub mnuIterations_Click(Index As Integer)
    
    ' Get iteration deepness level
    lIterationDeep = Choose(Index, 3, 4, 5, 6, 7, 8, 9)
    
    ' Set new checkmark
    SetNewMenuCheck mnuIterations, Index
    
End Sub


Private Sub mnuExit_Click()
    
    ' Exit application
    Unload Me

End Sub


Private Sub mnuAbout_Click()

    MsgBox " This is just a small demonstration for the very fast (fastest?)" + vbCrLf + _
            "gradient sub  'DrawTopDownGradient()'  without any further use." + vbCrLf + _
            "Using this shell comparing gradient methods should be easy." + vbCrLf + vbCrLf + _
            "  Light Templer / Carles P.V.  in Nov/Dec 2004 for PSC community", _
            vbInformation + vbMsgBoxSetForeground, _
            " About 'Let's Talk About SPEED !'"

End Sub

Private Function GetFormArea() As tpAPI_RECT
    ' Get client area of form in pixel
    
    API_GetClientRect Me.hwnd, GetFormArea
    With GetFormArea
        .lRight = .lRight - .lLeft
        .lLeft = 0
        .lBottom = .lBottom - .lTop
        .lTop = 0
    End With
    
End Function

Private Sub SetNewMenuCheck(MenuRow As Variant, ByVal lNewCheckIndex As Long)
    ' Clear all checkmarks in an indexed menu and set checkmark to the new you wanted to be checked
    
    Dim MenuEntry As Menu
    
    For Each MenuEntry In MenuRow
        MenuEntry.Checked = IIf(MenuEntry.Index = lNewCheckIndex, True, False)
    Next MenuEntry
    
End Sub


' #*#
