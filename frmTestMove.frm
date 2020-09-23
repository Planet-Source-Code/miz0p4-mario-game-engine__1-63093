VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmTestMove 
   BackColor       =   &H00FF5225&
   Caption         =   "Mario101"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   ForeColor       =   &H00979797&
   Icon            =   "frmTestMove.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00939393&
      Caption         =   "Control Panel"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Tag             =   "0"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Timer tmrRun 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   2040
         Tag             =   "0"
         Top             =   240
      End
      Begin VB.Timer TjumpD 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1560
         Top             =   240
      End
      Begin VB.Timer TJump 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1080
         Tag             =   "0"
         Top             =   240
      End
      Begin VB.Timer TWalkingRight 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   600
         Tag             =   "0"
         Top             =   240
      End
      Begin VB.Timer TWalkingLeft 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Tag             =   "0"
         Top             =   240
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      CausesValidation=   0   'False
      Height          =   735
      Left            =   7680
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   4695
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8281
      _cy             =   1296
   End
   Begin VB.Label lblDiscbeTurnBig 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Mario Big"
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line Line10 
      X1              =   3000
      X2              =   1800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line9 
      X1              =   840
      X2              =   840
      Y1              =   960
      Y2              =   600
   End
   Begin VB.Line Line8 
      X1              =   360
      X2              =   360
      Y1              =   600
      Y2              =   1320
   End
   Begin VB.Label lblD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblDiscribeStext 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Mario Small"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   2160
      X2              =   840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run  [no imge yet]"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Line Line7 
      X1              =   1800
      X2              =   360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblDiscribeLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "Go Left"
      Height          =   255
      Left            =   9960
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Line Line6 
      X1              =   9960
      X2              =   10560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblDiscribeRight 
      BackStyle       =   0  'Transparent
      Caption         =   "Go Right"
      Height          =   255
      Left            =   12360
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Line Line5 
      X1              =   12240
      X2              =   13080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblDiscribeJump 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump"
      Height          =   255
      Left            =   11760
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   11640
      X2              =   12120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblDiscribeDuck 
      BackStyle       =   0  'Transparent
      Caption         =   "Duck"
      Height          =   255
      Left            =   11400
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   11880
      X2              =   11400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   11400
      X2              =   11400
      Y1              =   1440
      Y2              =   1200
   End
   Begin VB.Label lblUA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "/\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11160
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblRA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11760
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblDA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "\/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11160
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblLA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10560
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   2385
      Left            =   0
      Picture         =   "frmTestMove.frx":058A
      Top             =   2760
      Width           =   14460
   End
   Begin VB.Image MarioSprit 
      Height          =   480
      Left            =   6480
      Picture         =   "frmTestMove.frx":70A02
      Tag             =   "2"
      Top             =   2280
      Width           =   240
   End
End
Attribute VB_Name = "frmTestMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    strMSize = "Sprits\Big\"
    intSpeed = 70
    intJhight = -100
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
       ''''''''''''' This is to make sure no keys will
       ' Case Else  ' Change whats being held down
       '''''''''''''
       
        Case 37 ' Left Arrow Key
            TWalkingLeft.Enabled = True
            MarioSprit.Tag = 1
            
        Case 38 ' Up Arrow Key
            TJump.Enabled = True
            WindowsMediaPlayer1.URL = (App.Path & "\" & "Sounds\smw_jump.wav")
        
        Case 39 ' Right Arrow Key
            TWalkingRight.Enabled = True
            MarioSprit.Tag = 2
            
        Case 40 ' Down Arrow Key
            If MarioSprit.Tag = 2 Then
                MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioDownRight.gif")
            ElseIf MarioSprit.Tag = 1 Then
                MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioDownLeft.gif")
            End If
            
        Case 65 ' running
            tmrRun.Enabled = True
        Case 68
            strMSize = "Sprits\Big\"
        Case 83
            strMSize = "Sprits\Small\"
        
    End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
       ''''''''''''' This is to make sure no keys will
       ' Case Else  ' Change whats being held down
       '''''''''''''
        
        Case 37 ' Left Arrow Key
            TWalkingLeft.Enabled = False
            MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandLeft.gif")
            
        Case 38 ' Up Arrow Key
        'don't need anything, timer must finish so mario will fall.
        
        Case 39 ' Right Arrow Key
            TWalkingRight.Enabled = False
            MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandRight.gif")
            
        Case 40 ' Down Arrow Key
            If MarioSprit.Tag = 2 Then
                MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandRight.gif")
            ElseIf MarioSprit.Tag = 1 Then
                MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandLeft.gif")
            End If
            
        Case 65 ' keybord letter a, for running
            tmrRun.Enabled = False
            intSpeed = 70
            tmrRun.Tag = 0
    End Select
    
End Sub

Private Sub tmrRun_Timer() 'this event is to make mario start runing after some time
'this section needs to be worked on. ifstatments need to be modified
'to know if the user is pressing a key, to know when to start the timer...
    tmrRun.Tag = tmrRun.Tag + 1 'seting up the timer's tag(istead of a var) to count
    If tmrRun.Tag < 64 Then 'less then 64, do normal running
        intSpeed = 160
    ElseIf tmrRun.Tag > 65 Then ' higher then 65 do  high speed runing
        intSpeed = 220
        'MarioSprit.Pictur
    End If

End Sub

Private Sub TWalkingLeft_Timer()

    TWalkingRight.Tag = TWalkingRight.Tag + 1 'Adding 1 to tag, to creat a sprite chart.
    MarioSprit.Left = MarioSprit.Left - intSpeed 'walking left
    
    If TWalkingRight.Tag = 4 Then 'Reseting to keep walking
        TWalkingRight.Tag = 0
    End If

        '|| Walking Left Sprits ||'
    If TWalkingRight.Tag = 1 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStepLeft1.gif")
    ElseIf TWalkingRight.Tag = 2 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStepLeft2.gif")
    ElseIf TWalkingRight.Tag = 3 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandLeft.gif")
    End If

End Sub

Private Sub TWalkingRight_Timer()

    TWalkingRight.Tag = TWalkingRight.Tag + 1 'Adding 1 to tag, to creat a sprite chart.
    MarioSprit.Left = MarioSprit.Left + intSpeed

    If TWalkingRight.Tag = 4 Then 'Reseting to keep walking
        TWalkingRight.Tag = 0
    End If

        '|| Walking Right Sprits ||'
    If TWalkingRight.Tag = 1 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStepRight1.gif")
    ElseIf TWalkingRight.Tag = 2 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStepRight2.gif")
    ElseIf TWalkingRight.Tag = 3 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandRight.gif")
    End If

End Sub

Private Sub TJump_Timer()
    MarioSprit.Top = MarioSprit.Top + intJhight
    
    If MarioSprit.Tag = 2 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioJumpUpRight.gif")
        If MarioSprit.Top <= 1300 Then
            intJhight = 100
            TjumpD.Enabled = True
        End If
    End If
    
    If MarioSprit.Tag = 1 Then
            MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioJumpUpLeft.gif")
        If MarioSprit.Top <= 1300 Then
            intJhight = 100
            TjumpD.Enabled = True
        End If
    End If
End Sub

Private Sub TjumpD_Timer()
    If MarioSprit.Tag = 2 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioJumpDownRight.gif")
            If MarioSprit.Top >= 2280 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandRight.gif")
            MarioSprit.Top = 2280
            TJump.Enabled = False
            TjumpD.Enabled = False
            intJhight = -100
        End If
    End If
    
    If MarioSprit.Tag = 1 Then
        MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioJumpDownLeft.gif")
            If MarioSprit.Top >= 2280 Then
                MarioSprit.Picture = LoadPicture(App.Path & "\" & strMSize + "MarioStandLeft.gif")
            MarioSprit.Top = 2280
            TJump.Enabled = False
            TjumpD.Enabled = False
            intJhight = -100
        End If
    End If
    
End Sub
