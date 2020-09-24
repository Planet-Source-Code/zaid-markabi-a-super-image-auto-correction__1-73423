VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Auto-Correct         -  by   Zaid Markabi           -   zaidmarkabi@yahoo.com            -   yazanmarkabi.com"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   12855
      TabIndex        =   7
      Top             =   4440
      Width           =   12855
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fog"
         Height          =   375
         Index           =   5
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Burn"
         Height          =   375
         Index           =   2
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Night"
         Height          =   375
         Index           =   4
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Water"
         Height          =   375
         Index           =   1
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sunlight"
         Height          =   375
         Index           =   3
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   615
         Left            =   11760
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   7440
         Max             =   10
         Min             =   1
         TabIndex        =   18
         Top             =   480
         Value           =   5
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   6720
         Max             =   10
         Min             =   1
         TabIndex        =   17
         Top             =   1320
         Value           =   6
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   8760
         Max             =   10
         Min             =   1
         TabIndex        =   16
         Top             =   1320
         Value           =   6
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   10800
         Max             =   10
         Min             =   1
         TabIndex        =   15
         Top             =   1320
         Value           =   6
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   9480
         Max             =   4
         Min             =   1
         TabIndex        =   14
         Top             =   480
         Value           =   2
         Width           =   1935
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Natural"
         Height          =   375
         Index           =   0
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Auto-Correct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   240
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Adjust Grade :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7320
         TabIndex        =   27
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Adjust Red :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   6600
         TabIndex        =   26
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Adjust Green :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   8640
         TabIndex        =   25
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Adjust Blue :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   10680
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Loop (slower) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9360
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   5520
         X2              =   5520
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Auto Effects (Standard)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2160
         TabIndex        =   22
         Top             =   120
         Width           =   2460
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CUSTOM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   330
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<"
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   3960
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   6480
      ScaleHeight     =   4185
      ScaleWidth      =   6465
      TabIndex        =   4
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save .."
         Height          =   255
         Left            =   5760
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   " After "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   4215
      ScaleWidth      =   6300
      TabIndex        =   1
      Top             =   120
      Width           =   6300
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   720
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   " Original "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Image Auto-Correction - 2010.9.10
' this tool can be used to Correct image colors, Refresh old photos, Adjust colors.
'
' This code had written by
'  Zaid Markabi , Arabic Syrian student.
'
' Email :   zaidmarkabi@yahoo.com
' Website : yazanmarkabi.com  or  yazanmarkabi.webs.com
'

Private Sub cmdApply_Click()
Fast_Adjust Picture1, Picture2, HScroll1.Value, HScroll2.Value, HScroll3.Value, HScroll4.Value, HScroll5.Value
End Sub

Private Sub cmdMove_Click()
Picture1.Picture = Picture2.Image
End Sub

Private Sub cmdSave_Click()
SavePicture Picture2.Image, "c:\Output_AutoCorrect.bmp"
MsgBox "Image had saved to this path :" + vbCrLf + vbCrLf + "C:\Output_AutoCorrect.bmp", vbOKOnly + vbInformation, "Save"
End Sub

Private Sub cmdTest_Click(Index As Integer)
Select Case Index
Case Is = 0
    Fast_Adjust Picture1, Picture2, 5, 9, 9, 5, 2
Case Is = 1
    Fast_Adjust Picture1, Picture2, 6, 3, 7, 9, 2
Case Is = 2
    Fast_Adjust Picture1, Picture2, 5, 9, 5, 5, 3
Case Is = 3
    Fast_Adjust Picture1, Picture2, 5, 9, 9, 9, 3
Case Is = 4
    Fast_Adjust Picture1, Picture2, 5, 1, 1, 1, 2
Case Is = 5
    Fast_Adjust Picture1, Picture2, 2, 16, 16, 16, 1
Case Is = 6
    Fast_Adjust Picture1, Picture2, 5, 6, 6, 6, 2
End Select
End Sub

Private Sub Command1_Click()
On Error GoTo Err
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
Picture2.Picture = Picture1.Picture
Err:
End Sub

Private Sub Form_Load()
cmdTest_Click 0
End Sub

Private Sub HScroll1_Change()
cmdApply_Click
End Sub

Private Sub HScroll2_Change()
cmdApply_Click
End Sub

Private Sub HScroll3_Change()
cmdApply_Click
End Sub

Private Sub HScroll4_Change()
cmdApply_Click
End Sub

Private Sub HScroll5_Change()
cmdApply_Click
End Sub
