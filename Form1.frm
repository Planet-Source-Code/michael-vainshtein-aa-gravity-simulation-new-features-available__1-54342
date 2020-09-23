VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Gravity simulation"
   ClientHeight    =   7245
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Launch (click on the form to set launch point)"
      Height          =   465
      Left            =   6945
      TabIndex        =   8
      Top             =   60
      Width           =   2625
   End
   Begin VB.Timer tmrT 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   8505
      Top             =   3570
   End
   Begin VB.PictureBox picWind 
      AutoRedraw      =   -1  'True
      Height          =   1515
      Left            =   6720
      ScaleHeight     =   1455
      ScaleWidth      =   2970
      TabIndex        =   2
      Top             =   2640
      Width           =   3030
   End
   Begin MSComctlLib.Slider sldGrav 
      Height          =   300
      Left            =   6825
      TabIndex        =   0
      Top             =   825
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      Max             =   30
      SelStart        =   1
      Value           =   1
   End
   Begin MSComctlLib.Slider sldElast 
      Height          =   300
      Left            =   6870
      TabIndex        =   4
      Top             =   1455
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.Image OBJ 
      Height          =   480
      Left            =   330
      Picture         =   "Form1.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6990
      TabIndex        =   9
      Top             =   3825
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Plastic"
      Height          =   195
      Left            =   9360
      TabIndex        =   7
      Top             =   1770
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Elastic"
      Height          =   195
      Left            =   6840
      TabIndex        =   6
      Top             =   1755
      Width           =   465
   End
   Begin VB.Label Label3 
      Caption         =   "Collision Type:"
      Height          =   210
      Left            =   6945
      TabIndex        =   5
      Top             =   1275
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Wind: (Click on the picture box to set wind speed and direction)"
      Height          =   480
      Left            =   6675
      TabIndex        =   3
      Top             =   2130
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "Gravity Force:"
      Height          =   210
      Left            =   6900
      TabIndex        =   1
      Top             =   615
      Width           =   1725
   End
   Begin VB.Label StartP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   48
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   345
      TabIndex        =   10
      Top             =   450
      Width           =   420
   End
   Begin VB.Menu mnuSH 
      Caption         =   "Show\Hide controls"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GForce, Vy, Vx, windX, windY, K, StartX, StartY

Private Sub cmdStart_Click()
    Me.Cls
    OBJ.Move StartX - OBJ.Width / 2, StartY - OBJ.Height / 2
    StartP.Move StartX - StartP.Width / 2, StartY - StartP.Height / 2
    If cmdStart.Caption <> "Stop" Then
        tmrT = True
        cmdStart.Caption = "Stop"
        Vx = 0
        Vy = 0
    Else
        tmrT = False
        cmdStart.Caption = "Launch"
    End If
End Sub

Private Sub Form_Load()
    sldGrav_Click
    sldElast_Click
    StartX = 435
    StartY = 435
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartX = X
    StartY = Y
    StartP.Move StartX - StartP.Width / 2, StartY - StartP.Height / 2
End Sub


Private Sub Form_Resize()
    cmdStart.Move ScaleWidth - cmdStart.Width - 250
    sldGrav.Move ScaleWidth - sldGrav.Width - 100
    sldElast.Left = sldGrav.Left
    Label1.Move sldGrav.Left
    Label2.Move ScaleWidth - Label2.Width - 50
    Label3.Move sldElast.Left
    Label4.Move sldElast.Left
    Label5.Move ScaleWidth - Label4.Width - 50
    picWind.Move ScaleWidth - picWind.Width - 50
End Sub

Private Sub mnuSH_Click()
    cmdStart.Visible = Not cmdStart.Visible
    sldGrav.Visible = Not sldGrav.Visible
    sldElast.Visible = Not sldElast.Visible
    Label1.Visible = Not Label1.Visible
    Label2.Visible = Not Label2.Visible
    Label3.Visible = Not Label3.Visible
    Label4.Visible = Not Label4.Visible
    Label5.Visible = Not Label5.Visible
    picWind.Visible = Not picWind.Visible
End Sub

Private Sub picWind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picWind.Tag = "1"
End Sub

Private Sub picWind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const LineNum = 8
    Const LowK = 8000
    Dim A
    If picWind.Tag = "1" Then
        picWind.Cls
        picWind.DrawWidth = 2
        For A = 0 To LineNum - 1
            picWind.Line (picWind.Width / 2 + 100 * (A - LineNum / 2), A * picWind.Height / LineNum)-(A * picWind.Width / LineNum / 3 + X, Y + A * picWind.Height / LineNum)
        Next
        windX = X / LowK
        windY = Y / LowK
        lblData = windX / LowK & vbNewLine & windY / LowK
    End If
End Sub

Private Sub picWind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picWind.Tag = "0"
End Sub

Private Sub sldElast_Click()
    K = sldElast.Value
End Sub

Private Sub sldGrav_Click()
    GForce = -sldGrav.Value
End Sub



Private Sub tmrT_Timer()
Static LastTimeHitBorder As Boolean
sldGrav_Click
sldElast_Click
    With OBJ
        'Gravity effect
        Vy = Vy + GForce - windY
        .Top = .Top - Vy
        
        'Wind effect
        Vx = Vx + windX
        .Left = .Left + Vx
        
    If .Top + .Height >= Me.ScaleHeight Then
        Vy = -Vy + Vy * K / 100
        
        .Top = Me.ScaleHeight - .Height
        If LastTimeHitBorder = True Then tmrT.Enabled = False
        If Vy = 0 Then tmrT.Enabled = False:       cmdStart.Caption = "Launch"
        LastTimeHitBorder = True
    End If
    If .Top < 0 Then
        Vy = -Vy + Vy * K / 100
        
        .Top = 0
        If LastTimeHitBorder = True Then tmrT.Enabled = False
        If Vy = 0 Then tmrT.Enabled = False:       cmdStart.Caption = "Launch"
        LastTimeHitBorder = True
    End If
    If .Left + .Width >= Me.ScaleWidth Then
        Vx = -Vx + Vx * K / 100
        
        .Left = Me.ScaleWidth - .Width
        If LastTimeHitBorder = True Then tmrT.Enabled = False
        If Vx = 0 Then tmrT.Enabled = False:       cmdStart.Caption = "Launch"
        LastTimeHitBorder = True
    End If
    If .Left <= 0 Then
        Vx = -Vx + Vx * K / 100
        
        .Left = 0
        If LastTimeHitBorder = True Then tmrT.Enabled = False
        If Vx = 0 Then tmrT.Enabled = False:       cmdStart.Caption = "Launch"
        LastTimeHitBorder = True
    End If
    
    LastTimeHitBorder = False
    End With
    
    Me.PSet (OBJ.Left + OBJ.Width / 2, OBJ.Top + OBJ.Height / 2)
    lblData = Vy & vbNewLine & Vx
End Sub
