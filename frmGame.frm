VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StickWarz 3000"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox SM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   1
      Left            =   6960
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox SS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   1
      Left            =   8040
      Picture         =   "frmGame.frx":C042
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   120
      ScaleHeight     =   273
      ScaleMode       =   0  'User
      ScaleWidth      =   343.045
      TabIndex        =   9
      Top             =   480
      Width           =   5100
      Begin VB.PictureBox pEN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   2760
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   19
         Top             =   3915
         Width           =   855
      End
      Begin VB.PictureBox pEN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   2760
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   18
         Top             =   45
         Width           =   855
      End
      Begin VB.PictureBox pHP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   1320
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   16
         Top             =   3915
         Width           =   855
      End
      Begin VB.PictureBox pHP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   1320
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   15
         Top             =   45
         Width           =   855
      End
      Begin VB.PictureBox Board 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawWidth       =   5
         ForeColor       =   &H80000008&
         Height          =   3630
         Left            =   120
         Picture         =   "frmGame.frx":18084
         ScaleHeight     =   240
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   320
         TabIndex        =   10
         Top             =   240
         Width           =   4830
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "EN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   20
         Top             =   3870
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "EN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   17
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   14
         Top             =   3870
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   13
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   3870
         Width           =   735
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox FM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6120
      Picture         =   "frmGame.frx":504C6
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox FS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6120
      Picture         =   "frmGame.frx":50808
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox GS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   6240
      Picture         =   "frmGame.frx":50B4A
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox GM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   6240
      Picture         =   "frmGame.frx":50E8C
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox bSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   -120
      Picture         =   "frmGame.frx":511CE
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox bMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   1080
      Picture         =   "frmGame.frx":89610
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Game"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox SM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   0
      Left            =   6960
      Picture         =   "frmGame.frx":C1A52
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox SS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   0
      Left            =   8040
      Picture         =   "frmGame.frx":CDA94
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FPS As Long
Dim sFPS As Long
Dim tSpeed As Long

Function MainLoop()
Dim C As Long
tSpeed = 5000
Running = True
Do Until Running = False
    If C >= tSpeed Then
        Board.Cls
        BitBlt Board.hDC, 0, 0, 320, 240, bMask.hDC, 0, 0, vbSrcAnd
        BitBlt Board.hDC, 0, 0, 320, 240, bSprite.hDC, 0, 0, vbSrcInvert
        GetKeys P(1), vbKeyW, vbKeyA, vbKeyD, vbKeyZ, vbKeyX, vbKeyC, vbKeyS, vbKeyQ
        GetKeys P(2), vbKeyI, vbKeyJ, vbKeyL, vbKeyM, 188, 190, vbKeyK, vbKeyU
        MovePlayers
        MoveShots
        C = 0
    Else
        C = C + 1
    End If
    DoEvents
Loop
End Function

Private Sub cmdNew_Click()
NewPlayer P(1), 2, 2, 1, 1
NewPlayer P(2), 320 - 34, 2, 2, 1
MainLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
Running = False
End Sub
