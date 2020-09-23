VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "VB PreProcessor"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   780
      Width           =   8250
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Temporary Project"
      Height          =   495
      Left            =   5085
      TabIndex        =   3
      Top             =   4245
      Width           =   3165
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   9165
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5010
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Enter the filename to compile up above."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy to Temporary Folder and Compile"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   4245
      Width           =   3195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Methods
Private Sub Command1_Click()
    Dim i As Long, FoundIt As Boolean
    CompileProject Combo1.Text
    For i = 0 To Combo1.ListCount - 1
        If StrComp(Combo1.Text, Combo1.List(i), vbTextCompare) = 0 Then
            FoundIt = True
            Exit For
        End If
    Next i
    If FoundIt = False Then
        Combo1.AddItem Combo1.Text
    End If
End Sub

Private Sub Command2_Click()
    OpenTempProject
End Sub

Private Sub Form_Load()
    'Load the settings
    Dim s As String, s2() As String, i As Long
    Dim LastVBP As String
    s = GetSetting(App.Title, "Preferences", "VBP", App.Path & "\Simple Project\Project1.vbp")
    LastVBP = GetSetting(App.Title, "Preferences", "LastVBP", App.Path & "\Simple Project\Project1.vbp")
    s2 = Split(s, vbCrLf)
    For i = 0 To UBound(s2, 1)
        Combo1.AddItem s2(i)
    Next i
    Combo1.Text = LastVBP
    Text1.BackColor = BackColor
    LoadFile App.Path & "\Read Me.txt", s
    Text1.Text = s
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save the settings
    Dim s As String, i As Long
    For i = 0 To Combo1.ListCount - 1
        s = s & Combo1.List(i) & vbCrLf
    Next i
    s = Left$(s, Len(s) - 2)
    SaveSetting App.Title, "Preferences", "VBP", s
    SaveSetting App.Title, "Preferences", "LastVBP", Combo1.Text
End Sub
