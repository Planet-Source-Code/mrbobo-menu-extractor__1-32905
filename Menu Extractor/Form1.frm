VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Menu Extractor"
   ClientHeight    =   4305
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10440
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Caption"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Shortcut"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Key"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Index"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Checked"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Enabled"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Visible"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "HelpContextID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "WindowList"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Negotiate"
         Object.Width           =   882
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   4080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   5
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExtract 
         Caption         =   "&Extract Menu"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileInsert 
         Caption         =   "&Insert Menu in..."
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Menu As..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBU 
         Caption         =   "&Always make backup when inserting"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'©2002 PSST Software and MrBobo
'Press help for functions descriptions
Private Enum SaveType 'Used to create appropriate headers for new items
    mForm = 1
    mMDIForm = 2
    mUserControl = 3
End Enum
Private Type bVBmenu 'Used to hold Menu properties
    Caption As String
    Key As String
    Index As Long
    WindowList As String
    Checked As String
    Enabled As String
    Visible As String
    HelpContextID As Long
    Shortcut As String
    Negotiate As Long
End Type
Dim BeforeString As String 'Existing header ABOVE menu data
Dim AfterString As String 'Existing header(and code) BELOW menu data
Dim MenuString As String 'Menu data
Dim FullString As String 'Entire file

Private Sub Form_Load()
    Dim mycommand As String
    mycommand = Command()
    If mycommand <> "" Then
        MenuString = ExtractMenu(mycommand) 'Get the Menu data
        ExtractoFlange 'Fill the Listview with Menu data
    End If
    AssocBMU 'establish file association to .bmu files
    'Did we want Backups last time we ran ?
    mnuFileBU.Checked = (GetSetting("PSST SOFTWARE\" + App.EXEName, "Settings", "BackUp", "1") = "1")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save backup status
    SaveSetting "PSST SOFTWARE\" + App.EXEName, "Settings", "BackUp", IIf(mnuFileBU.Checked, "1", "0")
End Sub

Private Sub mnuFile_Click()
    'Disable save options if no file loaded
    mnuFileSave.Enabled = Len(MenuString) > 0
    mnuFileInsert.Enabled = Len(MenuString) > 0
End Sub

Private Sub mnuFileBU_Click()
    mnuFileBU.Checked = Not mnuFileBU.Checked
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExtract_Click()
    With cmnDlg
        .Filter = "VB6 Forms (*.frm)|*.frm|VB6 Control (*.ctl)|*.ctl|Menu Template (*.bmu)|*.bmu"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        MenuString = ExtractMenu(.FileName) 'Get the Menu data
        ExtractoFlange 'Fill the Listview with Menu data
    End With
End Sub
Private Function ExtractMenu(mFile As String) As String
    Dim z As Long, st As Long
    FullString = OneGulp(mFile)
    z = InStr(FullString, "Attribute VB_Name =") 'End of Menu data in file header
    If z = 0 Then GoTo woops
    AfterString = Mid(FullString, z, Len(FullString) - z + 1) 'Remember for later
    z = InStr(1, FullString, "Begin VB.Menu") 'Start of Menu data in header
    If z = 0 Then GoTo woops
    St2 = InStrRev(FullString, vbCrLf, z) 'Move to start of line
    BeforeString = Left(FullString, St2 - 1) 'Remember for later
    ExtractMenu = Mid(FullString, St2 + 2, Len(FullString) - Len(AfterString) - z - 3) 'Menu data
    Exit Function
woops:
    ExtractMenu = "" 'No menu data found
End Function

Private Sub mnuFileInsert_Click()
    Dim sfile As String, mExt As String, tmpMenu As String, tmpBeforeString As String, tmpAfterString As String
    Dim z As Long, temp As String
    On Error GoTo woops
Restart:
    With cmnDlg
        .DialogTitle = "Insert Menu into..."
        .Filter = "VB6 Form (*.frm)|*.frm|VB6 MDIForm (*.frm)|*.frm|VB6 Control (*.ctl)|*.ctl"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        sfile = .FileName
        temp = ChangeExt(.FileTitle)
        'If the file is a new file VB insists on correct naming to function
        If Not Left(temp, 1) Like "[a-zA-Z]" Then 'First letter in name cannot be a number
            MsgBox "This name is invalid for a VB6 file." + vbCrLf + "The first letter cannot be a number." + vbCrLf + "Please use a valid name."
            GoTo Restart
        End If
        For z = 1 To Len(temp) 'All of the name must be Alphanumeric
            If Not Mid(temp, z, 1) Like "[a-zA-Z0-9]" Then
                MsgBox "This name is invalid for a VB6 file." + vbCrLf + "Please use a valid name."
                GoTo Restart
            End If
        Next
        Select Case .FilterIndex
            Case 1, 2 'Standard Form or MDIform
                mExt = "frm"
            Case 3 'Usercontrol
                mExt = "ctl"
        End Select
        'In case the user left off an extension
        'or used an invalid extension...
        If InStr(sfile, ".") = 0 Then
            sfile = sfile + "." + mExt
        Else
            sfile = ChangeExt(sfile, mExt)
        End If
        'Insert to existing file
        If FileExists(sfile) Then
            'Remember current data
            tmpBeforeString = BeforeString
            tmpMenu = MenuString
            tmpAfterString = AfterString
            'Open target file and split into Before,Menu,After
            ExtractMenu sfile
            'Do a backup ?
            If mnuFileBU.Checked Then FileSave FullString, SafeSave(ChangeExt(sfile, "bak"))
            'Save
            FileSave BeforeString + vbCrLf + tmpMenu + vbCrLf + "End" + vbCrLf + AfterString, sfile
            'Return variables back to initially extracted menu data
            'in case we want to make another insertion or save
            BeforeString = tmpBeforeString
            MenuString = tmpMenu
            AfterString = tmpAfterString
        Else
            'new file - easy
            FileSave SaveString(ChangeExt(FileOnly(sfile)), .FilterIndex), sfile
        End If
    End With
woops:

End Sub

Private Sub mnuFileSave_Click()
    'Just save the menu data
    Dim sfile As String, mExt As String, tmpMenu As String
    On Error GoTo woops
    With cmnDlg
        .DialogTitle = "Save Menu As Template"
        .Filter = "Menu Template (*.bmu)|*.bmu"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        sfile = .FileName
        If InStr(sfile, ".") = 0 Then
            sfile = sfile + "." + "bmu"
        Else
            sfile = ChangeExt(sfile, "bmu")
        End If
        If FileExists(sfile) Then Kill sfile
        FileSave SaveString(ChangeExt(FileOnly(sfile)), .FilterIndex), sfile
    End With
woops:
End Sub
Private Function SaveString(mForm As String, mType As SaveType) As String
    'To create a new VB module with correct header
    Dim TT As String
    Select Case mType
        Case 1
            TT = "Begin VB.Form "
        Case 2
            TT = "Begin VB.MDIForm "
        Case 3
            TT = "Begin VB.UserControl "
    End Select
    SaveString = "VERSION 5.00" + vbCrLf + TT + mForm + vbCrLf + _
    MenuString + vbCrLf + "End" + vbCrLf + _
    "Attribute VB_Name = " + mForm + vbCrLf + _
    "Attribute VB_GlobalNameSpace = False" + vbCrLf + _
    "Attribute VB_Creatable = " + IIf(mType = 4, "True", "False") + vbCrLf + _
    "Attribute VB_PredeclaredId = " + IIf(mType = 4, "False", "True") + vbCrLf + _
    "Attribute VB_Exposed = False"
End Function


Private Sub ExtractoFlange()
    Dim z As Long, z2 As Long, st As Long, St2 As Long, BB As bVBmenu
    Dim a As Long, temp As String, temp2 As String, Prop As Collection, PropValue As Collection, PropIndent As Long
    Dim SubMenustring As String, lItem As ListItem
'    Sort the menu properties and display in a Listview
    st = 1
    LV.ListItems.Clear
    Do
        temp = ""
        With BB 'Empty current properties
            .Caption = ""
            .Checked = ""
            .Enabled = ""
            .HelpContextID = 0
            .Index = -1
            .Key = ""
            .Negotiate = 0
            .Shortcut = ""
            .Visible = ""
            .WindowList = ""
        End With
        z = InStr(st, MenuString, "Begin VB.Menu")
        If z = 0 Then Exit Do
        z2 = InStrRev(MenuString, vbCrLf, z + 1) 'Determine indentation
        If LV.ListItems.Count = 0 Then
            PropIndent = 0
        Else
            PropIndent = z - z2 - 5
        End If
        Set lItem = LV.ListItems.Add(, , String(PropIndent * 3, ".")) 'Add indentation dots - caption comes later
        lItem.SubItems(2) = Mid(MenuString, z + 14, InStr(z + 14, MenuString, vbCrLf) - z - 14)
        z2 = InStr(z + 1, MenuString, "Begin VB.Menu") 'Locate where the next menuitem starts(if any)
        If z2 = 0 Then 'Put all relavent data for this menuitem into a variable we can parse
            SubMenustring = Right(MenuString, Len(MenuString) - z)
        Else
            SubMenustring = Mid(MenuString, z, z2 - z)
        End If
        Set Prop = New Collection
        Set PropValue = New Collection
        For a = 1 To UBound(Split(SubMenustring, vbCrLf)) 'Go through menu properties line by line
            temp = Trim(Split(SubMenustring, vbCrLf)(a))
            If InStr(temp, "=") Then 'its got an "=" - must be a property
                temp2 = Trim(Split(temp, "=")(1))
                If InStr(temp2, "'") Then
                    temp2 = Left(temp2, InStr(temp2, "'") - 1) 'VB sometimes puts comments on properties
                End If                                         'We dont need this
                PropValue.Add temp2 'Remember property value
                Prop.Add Trim(Split(temp, "=")(0)) 'Remember property name
            End If
        Next
        If Prop.Count > 0 Then 'Go through what was found
            For a = 1 To Prop.Count 'adding to our Type as needed
                Select Case LCase(Prop(a))
                    Case "caption"
                        temp2 = Right(PropValue(a), Len(PropValue(a)) - 1)
                        temp2 = Left(temp2, Len(temp2) - 1)
                        BB.Caption = temp2
                    Case "index"
                        BB.Index = PropValue(a)
                    Case "windowlist"
                        BB.WindowList = PropValue(a)
                   Case "checked"
                        BB.Checked = PropValue(a)
                    Case "enabled"
                        BB.Enabled = PropValue(a)
                    Case "visible"
                        BB.Visible = PropValue(a)
                    Case "helpcontextid"
                        BB.HelpContextID = Val(PropValue(a))
                    Case "shortcut"
                        BB.Shortcut = GetShortCutKey(PropValue(a))
                    Case "negotiateposition"
                        BB.Negotiate = Val(PropValue(a))
                End Select
            Next
            'We cleared the type info above because VB only records
            'properties that are set to a value
            With BB 'Load the type info into the Listview
                lItem.Text = lItem.Text + .Caption
                lItem.SubItems(3) = IIf(.Index > -1, .Index, "")
                lItem.SubItems(8) = IIf(.WindowList = "", "False", "True")
                lItem.SubItems(4) = IIf(.Checked = "", "False", "True")
                lItem.SubItems(5) = IIf(.Enabled = "", "True", "False")
                lItem.SubItems(6) = IIf(.Visible = "", "True", "False")
                lItem.SubItems(7) = .HelpContextID
                lItem.SubItems(1) = .Shortcut
                lItem.SubItems(9) = .Negotiate
            End With
            
        End If
        st = z + 1
    Loop
End Sub
Private Function GetShortCutKey(mKey As String) As String
    'convert to readable form
    mKey = Replace(mKey, "^", "Ctrl")
    mKey = Replace(mKey, "+", "Shift")
    mKey = Replace(mKey, "%", "Alt")
    mKey = Replace(mKey, "INSERT", "Ins")
    mKey = Replace(mKey, "BKSP", "Bksp")
    mKey = Replace(mKey, "DEL", "Del")
    mKey = Replace(mKey, "{", "")
    mKey = Replace(mKey, "}", "")
    GetShortCutKey = mKey
End Function



Private Sub mnuHelpAbout_Click()
    MsgBox "This is a demo of string parsing and file association" + vbCrLf + "in the form of a handy tool for VB6 programmers." + vbCrLf + vbCrLf + "Submitted to Planet Source Code on 21/3/2002 by MrBobo." + vbCrLf + "Hope you find it useful !", vbInformation, "PSST Software 2002"
End Sub

Private Sub mnuHelpHelp_Click()
    MsgBox "Functions" + vbCrLf + _
    "1.Extract Menu :" + vbCrLf + _
    "     Reads a VB form/usercontrol file into four variables" + vbCrLf + _
    "           FullString - entire file" + vbCrLf + _
    "           BeforeString - header before Menu data" + vbCrLf + _
    "           MenuString - Menu data" + vbCrLf + _
    "           AfterString - remainder of file" + vbCrLf + _
    "2.Insert Menu in... :" + vbCrLf + _
    "     Opens target file(or creates new file) and replaces Menu data already extracted" + vbCrLf + _
    "3.Save Menu As... :" + vbCrLf + _
    "     Creates a .bmu file - a menu template ready to insert into a VB file" + vbCrLf + _
    "4.Always make backup when inserting :" + vbCrLf + _
    "     If checked will make a backup of the target file with the" + vbCrLf + _
    "     extension .bak in the same directory as the target file" + vbCrLf + _
    "     If a backup already exists, a number will be appended to the filename." + vbCrLf + _
    "The program additionally associates itself to .bmu files and adds to" + vbCrLf + _
    "Explorers context menus for VB Forms/Usercontrols a menu item" + vbCrLf + _
    "Extract Menu. It does this with only 4 registry entries found in the" + vbCrLf + _
    "sub AssocBMU. If you wish to remove these settings simply delete the" + vbCrLf + _
    "4 registry settings and comment out the source.", vbInformation, "PSST Software 2002"


End Sub
