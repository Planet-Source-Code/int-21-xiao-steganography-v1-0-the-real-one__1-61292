VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Xiao Steganography 1.0"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   650
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   1950
      Begin VB.Label Label4 
         Caption         =   "Exit"
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmMain.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   260
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   105
         MouseIcon       =   "frmMain.frx":06DC
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":082E
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   650
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   1950
      Begin VB.Image imgExtract 
         Height          =   315
         Left            =   100
         MouseIcon       =   "frmMain.frx":0BED
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":0D3F
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lbExtract 
         AutoSize        =   -1  'True
         Caption         =   "Extract Files"
         Height          =   195
         Left            =   840
         MouseIcon       =   "frmMain.frx":1014
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame frAddFiles 
      Height          =   650
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   1950
      Begin VB.Label lbAdd 
         AutoSize        =   -1  'True
         Caption         =   "Add Files"
         Height          =   195
         Left            =   840
         MouseIcon       =   "frmMain.frx":1166
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   300
         Width           =   765
      End
      Begin VB.Image ImgAdd 
         Height          =   315
         Left            =   100
         MouseIcon       =   "frmMain.frx":12B8
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":140A
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Label lbAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   4440
      MouseIcon       =   "frmMain.frx":161C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgAbout 
      Height          =   240
      Left            =   5160
      MouseIcon       =   "frmMain.frx":176E
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":18C0
      Top             =   2880
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2505
      Left            =   120
      Top             =   240
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Developed by Int21"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image ImgLogo 
      Height          =   2505
      Left            =   120
      Picture         =   "frmMain.frx":1E4A
      Top             =   240
      Width           =   3180
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Xiao Steganography 1.0
'Developed by Int21
'This is my 1st version about this, i was working hard in this for 2 week(fulltime)
'* This version was tested only with bitmap file(.bmp) to carried the attach files
'* I'm not responsible if you try with another file
'* The files tested to be attached was plaintext(.txt) and jpg,gif,bmp,png.
'* NO was tested with another type file in this version

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub imgAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub ImgAdd_Click()
    frmAdd.Show vbModeless, Me
End Sub

Private Sub Label4_Click()
    Unload Me
End Sub

Private Sub lbAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub lbAdd_Click()
    frmAdd.Show vbModeless, Me
End Sub

Private Sub lbExtract_Click()
    frmRead.Show vbModeless, Me
End Sub

Private Sub imgExtract_Click()
    frmRead.Show vbModeless, Me
End Sub
