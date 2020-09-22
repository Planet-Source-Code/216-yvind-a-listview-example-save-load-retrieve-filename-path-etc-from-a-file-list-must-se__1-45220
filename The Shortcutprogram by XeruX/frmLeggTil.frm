VERSION 5.00
Begin VB.Form frmLeggTil 
   Caption         =   "Add a shortcut"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Description: (Action, RPG, Strategi etc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   3495
   End
End
Attribute VB_Name = "frmLeggTil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
' This adds the selected file from File1 into lst (the listview)
Dim sFileName As String:    sFileName = File1.List(File1.ListIndex) ' To get the filename
Dim sDirectory As String:   sDirectory = File1.Path & "\" ' To get the path
Dim sDescription As String: sDescription = txtDescription.Text ' To get the description from txtDescription.
If (Len(sFileName) = 0) Then Exit Sub
Dim objLvi As MSComctlLib.ListItem: Set objLvi = frmMain.lst.ListItems.Add()
objLvi.Text = sFileName
objLvi.SubItems(1) = sDirectory ' Inserts the directory into SubItem(1)
objLvi.SubItems(2) = FileLen(sDirectory & sFileName) & " bytes" ' To get the size of the file
If txtDescription = "" Then
txtDescription.Text = ""
Else
objLvi.SubItems(3) = sDescription
End If
Set objLvi = Nothing
frmMain.Caption = "Currently [" & frmMain.lst.ListItems.Count & "] files in list."
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub Form_Load()
' This makes every other filetype than .exe invisible:
File1.Pattern = "*.exe"
End Sub

