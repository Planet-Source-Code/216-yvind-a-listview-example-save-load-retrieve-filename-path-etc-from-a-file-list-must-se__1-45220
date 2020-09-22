VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Shortcut"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8040
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Remove all"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09FE
            Key             =   "Extract"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FBC
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":155F
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B07
            Key             =   "New"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   1588
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7305
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "20:49"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin MSComctlLib.ListView lst 
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6588
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Description"
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Select the type of .exe file (gamex, programs, etc)"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Begin VB.Menu mnuActionsAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuActionsSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuActionsOpen 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
MsgBox "Do you want to save changes?", vbYesNo, "Quit"
If vbYes Then: mnuActionsSave_Click
If vbNo Then: End
End Sub

Private Sub cmdLoad_Click()
' To open the selected program in the listview
Shell lst.SelectedItem.SubItems(1) & lst.SelectedItem.Text, vbNormalFocus
End Sub


Private Sub cmdRemove_Click()
    ' To remove the selected item in listview:
lst.ListItems.Remove (lst.SelectedItem.Index)
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End Sub


Private Sub cmdRemoveAll_Click()
' Clears the listview:
lst.ListItems.Clear
End Sub

Private Sub Combo1_Click()

    Select Case Combo1
    
    Case "Gamez"
    ' We need to clear the lsit so that the items which is in the last selected
    ' list wont get mixed with the newest selected list:
    
    lst.ListItems.Clear
    ' Then to load the selected list from the combobox:
    LoadLW lst, App.Path + "\Gamez.txt"
    ' Updates the caption with the number which says how many items there is in the list:
    frmMain.Caption = "Currently [" & frmMain.lst.ListItems.Count & "] files in list."
    Case "Programs"
    lst.ListItems.Clear
    LoadLW lst, App.Path + "\Programs.txt"
    ' Updates the caption with the number which says how many items there is in the list:
    frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
    Case "Mixed"
    lst.ListItems.Clear
    LoadLW lst, App.Path + "\Mixed.txt"
    ' Updates the caption with the number which says how many items there is in the list:
    frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
    Case "Everything"
    lst.ListItems.Clear
    LoadLW lst, App.Path + "\Everything.txt"
    LoadLW lst, App.Path + "\Mixed.txt"
    LoadLW lst, App.Path + "\Programs.txt"
    LoadLW lst, App.Path + "\Gamez.txt"
    ' Updates the caption with the number which says how many items there is in the list:
    frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
    End Select
    
End Sub

Private Sub Form_Load()
' Adds the categories into the combobox
Combo1.AddItem "Gamez"
Combo1.AddItem "Programs"
Combo1.AddItem "Mixed"
Combo1.AddItem "Everything"

End Sub

Private Sub mnuActionsAdd_Click()
frmLeggTil.Show
End Sub

Private Sub mnuActionsOpen_Click()
If Combo1 = "Gamez" Then
lst.ListItems.Clear
 ' Loads the selected categori from the combobox from a .txt file:
LoadLW lst, App.Path + "\Gamez.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "Programs" Then
lst.ListItems.Clear
LoadLW lst, App.Path + "\Programs.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "Mixed" Then
lst.ListItems.Clear
LoadLW lst, App.Path + "\Mixed.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "Everything" Then
lst.ListItems.Clear
LoadLW lst, App.Path + "\Everything.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "" Then
lst.View = 3
lst.ListItems.Clear
    LoadLW lst, App.Path + "\Everything.txt"
    ' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If
End Sub

Private Sub mnuActionsSave_Click()
If Combo1 = "Gamez" Then
SaveLW lst, App.Path + "\Gamez.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "Programs" Then
SaveLW lst, App.Path + "\Programs.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "Mixed" Then
SaveLW lst, App.Path + "\Mixed.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "Everything" Then
SaveLW lst, App.Path + "\Everything.txt"
' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

If Combo1 = "" Then
    SaveLW lst, App.Path + "\Everything.txt"
    ' Updates the caption with the number which says how many items there is in the list:
frmMain.Caption = "Currently [" & lst.ListItems.Count & "] files in list."
End If

End Sub

Private Sub mnuFileExit_Click()
MsgBox "Do you want to save changes?", vbYesNo, "Quit"
If vbYes Then: mnuActionsSave_Click
If vbNo Then: End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "Add"
            frmLeggTil.Show
        Case "Save"
If Combo1 = "Gamez" Then
 ' Saves the items into a .txt file:
SaveLW lst, App.Path + "\Gamez.txt"
End If

If Combo1 = "Programs" Then
SaveLW lst, App.Path + "\Programs.txt"
End If

If Combo1 = "Mixed" Then
SaveLW lst, App.Path + "\Mixed.txt"
End If

If Combo1 = "Everything" Then
SaveLW lst, App.Path + "\Everything.txt"
End If

If Combo1 = "" Then
    SaveLW lst, App.Path + "\Everything.txt"
End If
        Case "Open"
        ' To load the selected program in the listview:
        Shell lst.SelectedItem.SubItems(1) & lst.SelectedItem.Text, vbNormalFocus
    End Select
End Sub

