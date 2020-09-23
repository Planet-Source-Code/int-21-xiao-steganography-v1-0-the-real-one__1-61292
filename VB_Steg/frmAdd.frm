VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add file to image"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Files 
      Left            =   5880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      Begin VB.PictureBox PicCont 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         ScaleHeight     =   205
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   5
         Top             =   240
         Width           =   3615
         Begin VB.HScrollBar HS 
            Height          =   255
            Left            =   0
            Max             =   1
            Min             =   1
            TabIndex        =   12
            Top             =   2835
            Value           =   1
            Width           =   3255
         End
         Begin VB.VScrollBar VS 
            Height          =   2775
            Left            =   3315
            Max             =   1
            Min             =   1
            TabIndex        =   11
            Top             =   0
            Value           =   1
            Width           =   255
         End
         Begin VB.PictureBox TheImage 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   0
            ScaleHeight     =   183
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   215
            TabIndex        =   6
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   405
         Left            =   3960
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label cmdTarget 
         AutoSize        =   -1  'True
         Caption         =   "Load Image Target"
         Height          =   315
         Left            =   4440
         MouseIcon       =   "frmAdd.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2850
         Width           =   1755
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4080
         MouseIcon       =   "frmAdd.frx":06DC
         MousePointer    =   99  'Custom
         Picture         =   "frmAdd.frx":082E
         Top             =   2835
         Width           =   240
      End
   End
   Begin VB.Frame frStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Label lbProc 
         AutoSize        =   -1  'True
         Caption         =   "Wait while the files are joined with the image..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   5115
      End
   End
   Begin VB.Frame frStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Image Image7 
         Height          =   240
         Left            =   2880
         Picture         =   "frmAdd.frx":0C32
         Top             =   1725
         Width           =   240
      End
      Begin VB.Label cmdSee 
         AutoSize        =   -1  'True
         Caption         =   "See final image"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3240
         MouseIcon       =   "frmAdd.frx":0FF1
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1755
         Width           =   1335
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   2760
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lbDone 
         AutoSize        =   -1  'True
         Caption         =   "The process has finished."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2745
      End
   End
   Begin VB.Frame frStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ListView lvwFilesAdded 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Type"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size (K)"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   240
         Picture         =   "frmAdd.frx":1143
         Top             =   2925
         Width           =   240
      End
      Begin VB.Label cmdAddFile 
         AutoSize        =   -1  'True
         Caption         =   "Add File"
         Height          =   195
         Left            =   600
         MouseIcon       =   "frmAdd.frx":1A45
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   2955
         Width           =   1155
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   120
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Image Image5 
         Height          =   210
         Left            =   2040
         Picture         =   "frmAdd.frx":1B97
         Top             =   2925
         Width           =   225
      End
      Begin VB.Label cmdRemove 
         AutoSize        =   -1  'True
         Caption         =   "Remove File"
         Height          =   195
         Left            =   2400
         MouseIcon       =   "frmAdd.frx":2431
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   2955
         Width           =   1170
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   1920
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lbRema 
         AutoSize        =   -1  'True
         Caption         =   "Remaining Bytes:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1515
      End
   End
   Begin VB.Label lbStep 
      AutoSize        =   -1  'True
      Caption         =   "Step 1: Select the target image:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   3465
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   2520
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label cmdFinish 
      AutoSize        =   -1  'True
      Caption         =   "Finish"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      MouseIcon       =   "frmAdd.frx":2583
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   4155
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   2640
      Picture         =   "frmAdd.frx":26D5
      Top             =   4125
      Width           =   240
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   1320
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label cmdNext 
      AutoSize        =   -1  'True
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1680
      MouseIcon       =   "frmAdd.frx":2A94
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4155
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1440
      Picture         =   "frmAdd.frx":2BE6
      Top             =   4155
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   120
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label cmdBack 
      AutoSize        =   -1  'True
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      MouseIcon       =   "frmAdd.frx":3170
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4155
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   165
      Left            =   240
      Picture         =   "frmAdd.frx":32C2
      Top             =   4170
      Width           =   150
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSteps
Dim currentStep& 'Counter to know the current step in process
Const MAX_STEPS = 3 'Max step in this module

Dim strImageFile$ 'the main image where the files will be attach

Dim m_Millimeter&

Dim WithEvents clsStega As ClsStegano 'The magic class
Attribute clsStega.VB_VarHelpID = -1

'If you wants capture some error, do it here
Private Sub clsStega_SomeError(strDescription As String)
    MsgBox "Some error: " & strDescription
End Sub
'The current process is show it here
Private Sub clsStega_StatusChanged(prcDone As Long, strStatus As String)
    frmLoad.ProgBar.Value = prcDone
    frmLoad.lbStatus = strStatus
End Sub

Private Sub cmdAddFile_Click()
Dim theData$, theKey$, sTitle$
Dim Itlvw As ListItem
    'browse to get the files to be add
    Files.Filter = "PlainText|*.txt|Image Type|*.gif;*.jpg;*.bmp;*.png"
    Files.FileName = ""
    Files.ShowOpen
    theData = Files.FileName
    sTitle = VBA.Left$(Files.FileTitle, Len(Files.FileTitle) - 4)
    If theData <> "" Then
        theKey = "f0" & lvwFilesAdded.ListItems.Count + 1 'generate a unique key for this file
        If clsStega.AddFile(theData, sTitle, theKey) Then 'if was added... continue
            Set Itlvw = lvwFilesAdded.ListItems.Add(, theKey, VBA.Right$(theData, 3))
            Itlvw.SubItems(1) = sTitle
            Itlvw.SubItems(2) = FileLen(theData)
            lbRema = "Remaining Bytes: " & clsStega.BytesTotal - clsStega.BytesAdded 'Limit bytes to attach
            cmdNext.Enabled = True
        End If
    End If
End Sub

Private Sub cmdBack_Click()
    cmdNext.Enabled = True
    cmdFinish.Enabled = False
If currentStep >= 0 Then
    frStep(currentStep).Visible = False
    currentStep = currentStep - 1
    lbStep = vSteps(currentStep)
    If currentStep < 0 Then currentStep = 0
    
    frStep(currentStep).Visible = True
    If currentStep = 0 Then cmdBack.Enabled = False
End If
End Sub

Private Sub cmdFinish_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    cmdNext.Enabled = False
    cmdBack.Enabled = True
    frStep(currentStep).Visible = False
    currentStep = currentStep + 1
    lbStep = vSteps(currentStep)
    frStep(currentStep).Visible = True
    If currentStep >= MAX_STEPS Then
        cmdBack.Enabled = False
        cmdNext.Enabled = False
        cmdFinish.Enabled = True
    End If
    If currentStep = 2 Then 'Process data
        frmLoad.Show vbModeless, Me
        clsStega.EncodeIt 'do the magic
        
        Files.FileName = ""
        Files.Filter = "Image File|*.bmp"
        Files.ShowSave ' save the new image with the files added
        If Files.FileName <> "" Then
            clsStega.OutputImageFile = Files.FileName
            clsStega.Save2Image
            cmdSee.Enabled = True
            cmdNext_Click 'Next step, Finish
        Else
            MsgBox "Save image was cancel!"
        End If
        Unload frmLoad
    End If
End Sub

Private Sub cmdRemove_Click()
Dim ItSel As ListItem
    Set ItSel = lvwFilesAdded.SelectedItem
    If Not ItSel Is Nothing Then
        clsStega.RemoveFile ItSel.Key
        lbRema = "Remaining Bytes: " & clsStega.BytesTotal - clsStega.BytesAdded
        lvwFilesAdded.ListItems.Remove ItSel.Index
    End If
    If lvwFilesAdded.ListItems.Count = 0 Then cmdNext.Enabled = False
End Sub

Private Sub cmdSee_Click()
    Shell "explorer " & Files.FileName
End Sub

Private Sub cmdTarget_Click()
Dim ToW&, ToH&, InW&, InH&, mPad&
    mPad = 2


    Files.Filter = "Image File|*.bmp"
    Files.ShowOpen
    strImageFile = Files.FileName
    
    If strImageFile <> "" And VBA.Right$(strImageFile, 4) = ".bmp" Then
        TheImage.Picture = LoadPicture(strImageFile)
        clsStega.ImageFile = strImageFile
        'calculate the size the image to adjust in screen
        ToW = Me.ScaleWidth - mPad - mPad
        ToH = Me.ScaleHeight - mPad - mPad
        InW = TheImage.Picture.Width / m_Millimeter
        InH = TheImage.Picture.Height / m_Millimeter
        
        VS.Max = InH - HS.Top + mPad
        HS.Max = InW - VS.Left + mPad
        VS.LargeChange = ToH - HS.Height
        HS.LargeChange = ToW - VS.Width
        
        lbRema = "Remaining Bytes: " & clsStega.BytesTotal
        
        cmdNext.Enabled = True
    Else
        MsgBox "Error: trying to add unsupport image", vbCritical, "Bad image"
    End If
End Sub

Private Sub Form_Load()
    Set clsStega = New ClsStegano
    currentStep = 0
    'for future use
    'vHex = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", " E", " F")
    'vBin = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", "1000", "1001", "1010", "1011", "1100", "1101", " 1110", " 1111")

    m_Millimeter = ScaleX(100, vbPixels, vbMillimeters)
    'fill the name for each step process
    vSteps = Array("Step 1: Select the target image:", "Step 2: Select the file must be add into the target image:", "Step 3: Joining the files with the target image:", "Finished: Files joined")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsStega = Nothing
End Sub

Private Sub HS_Change()
    TheImage.Move -HS.Value
End Sub

Private Sub HS_Scroll()
    TheImage.Move -HS.Value
End Sub

Private Sub Image1_Click()
    Call cmdTarget_Click
End Sub

Private Sub VS_Change()
    TheImage.Move TheImage.Left, -VS.Value
End Sub

Private Sub VS_Scroll()
    TheImage.Move TheImage.Left, -VS.Value
End Sub
