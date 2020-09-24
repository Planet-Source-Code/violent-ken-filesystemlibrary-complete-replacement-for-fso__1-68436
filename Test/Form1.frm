VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test form"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Options de recherche"
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   7200
      Width           =   9255
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   5640
         TabIndex        =   30
         Top             =   360
         Width           =   1935
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   1695
            TabIndex        =   31
            Top             =   240
            Width           =   1695
            Begin VB.CheckBox chkRefresh 
               Caption         =   "Refresh infos"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   840
               Width           =   1335
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Object"
               Height          =   255
               Left            =   0
               TabIndex        =   37
               Top             =   480
               Width           =   1455
            End
            Begin VB.OptionButton optStr 
               Caption         =   "String"
               Height          =   255
               Left            =   0
               TabIndex        =   36
               Top             =   120
               Value           =   -1  'True
               Width           =   1455
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   9015
         TabIndex        =   18
         Top             =   240
         Width           =   9015
         Begin MSComCtl2.DTPicker dtCrea 
            Height          =   255
            Left            =   3000
            TabIndex        =   39
            Top             =   840
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   16515075
            UpDown          =   -1  'True
            CurrentDate     =   39200
         End
         Begin VB.Frame Frame4 
            Height          =   1695
            Left            =   7560
            TabIndex        =   32
            Top             =   120
            Width           =   1455
            Begin VB.CheckBox chkSubFolder 
               Caption         =   "Subfolder"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   120
               ScaleHeight     =   1335
               ScaleWidth      =   1215
               TabIndex        =   33
               Top             =   240
               Width           =   1215
               Begin VB.CheckBox chkCase 
                  Caption         =   "Case"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   43
                  Top             =   1080
                  Value           =   1  'Checked
                  Width           =   1215
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "File"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  Top             =   480
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton optFolder 
                  Caption         =   "Folder"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   120
                  Width           =   855
               End
            End
         End
         Begin VB.ComboBox opMod 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form1.frx":0000
            Left            =   2040
            List            =   "Form1.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox opAcc 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form1.frx":0028
            Left            =   2040
            List            =   "Form1.frx":003B
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox opCrea 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form1.frx":0050
            Left            =   2040
            List            =   "Form1.frx":0063
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox opSize 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form1.frx":0078
            Left            =   2040
            List            =   "Form1.frx":008B
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSize 
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txtString 
            Height          =   255
            Left            =   3000
            TabIndex        =   24
            Top             =   120
            Width           =   2295
         End
         Begin VB.CheckBox chkMod 
            Caption         =   "DateLastModified"
            Height          =   195
            Left            =   0
            TabIndex        =   23
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CheckBox chkAcc 
            Caption         =   "DateLastAccess"
            Height          =   195
            Left            =   0
            TabIndex        =   22
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox chkCrea 
            Caption         =   "DateCreated"
            Height          =   195
            Left            =   0
            TabIndex        =   21
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox chkSize 
            Caption         =   "Size"
            Height          =   195
            Left            =   0
            TabIndex        =   20
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkString 
            Caption         =   "String"
            Height          =   195
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtAcc 
            Height          =   255
            Left            =   3000
            TabIndex        =   40
            Top             =   1200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   16515075
            UpDown          =   -1  'True
            CurrentDate     =   39200
         End
         Begin MSComCtl2.DTPicker dtMod 
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   1560
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   16515075
            UpDown          =   -1  'True
            CurrentDate     =   39200
         End
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search NOW !"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   9480
      Width           =   9255
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   120
         ScaleHeight     =   6735
         ScaleWidth      =   4215
         TabIndex        =   9
         Top             =   120
         Width           =   4215
         Begin VB.ListBox lst 
            Height          =   5325
            Left            =   0
            TabIndex        =   15
            Top             =   1320
            Width           =   4215
         End
         Begin VB.CheckBox chkSub 
            Caption         =   "Include sub folders"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   960
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CommandButton cmdEnumerateFiles 
            Caption         =   "Enumerate files"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   13
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton cmdEnumeratesFolder 
            Caption         =   "Enumerate folders"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton cmdBrowse2 
            Caption         =   "..."
            Height          =   255
            Left            =   3600
            TabIndex        =   11
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtFolder 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   3375
         End
      End
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4080
      Width           =   4695
   End
   Begin VB.PictureBox pctIcon 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3960
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   6
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Drive informations"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Change folder..."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdInfoFolder 
      Caption         =   "Folder informations"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2565
      Hidden          =   -1  'True
      Left            =   2880
      System          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISearchEvent

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private FS As FileSystemLibrary.FileSystem

Private Sub chkAcc_Click()
    opAcc.Enabled = CBool(chkAcc.Value)
End Sub

Private Sub chkCrea_Click()
    opCrea.Enabled = CBool(chkCrea.Value)
End Sub

Private Sub chkMod_Click()
    opMod.Enabled = CBool(chkMod.Value)
End Sub

Private Sub chkSize_Click()
    opSize.Enabled = CBool(chkSize.Value)
End Sub

Private Sub cmdSearch_Click()
Dim MySearch As SearchDefinition
Dim Folders As Folders
Dim Files As Files
Dim s() As String
Dim o As String

    lst.Clear
    
    o = FS.BrowseForFolder("Dossier source", Me.hWnd)
    
    If FS.FolderExists(o) = False Then Exit Sub
    
    With MySearch
        .FolderName = o
        .DateCreatedOperator = GetOp(Me.opCrea)
        .DateCreatedValue = dtCrea.Value
        .DateLastAccessedOperator = GetOp(Me.opAcc)
        .DateLastAccessedValue = dtAcc.Value
        .DateLastModifiedOperator = GetOp(Me.opMod)
        .DateLastModifiedValue = dtMod.Value
        .RespectCase = CBool(chkCase.Value)
        .SearchName = Me.txtString
        .SizeOperator = GetOp(Me.opSize)
        .SizeValue = Val(txtSize.Text)
        .SubFolder = CBool(chkSubFolder.Value)
        
        If Me.chkAcc.Value Then .Criteria = SearchDateLastAccessed + .Criteria
        If Me.chkCrea.Value Then .Criteria = SearchDateCreated + .Criteria
        If Me.chkMod.Value Then .Criteria = SearchDateLastModified + .Criteria
        If Me.chkSize.Value Then .Criteria = SearchSize + .Criteria
        If Me.chkString.Value Then .Criteria = SearchName + .Criteria
        
    End With
    
    If optStr.Value Then
        If optFolder.Value Then
            s() = FS.SearchForFoldersStr(MySearch, Me)
        Else
            s() = FS.SearchForFilesStr(MySearch, Me)
        End If
    Else
        If optFolder.Value Then
            Set Folders = FS.SearchForFolders(MySearch, CBool(Me.chkRefresh.Value), Me)
        Else
            Set Files = FS.SearchForFiles(MySearch, CBool(Me.chkRefresh.Value), Me)
        End If
    End If
End Sub

Private Sub ISearchEvent_ItemFound(Item As String)
   lst.AddItem Item
End Sub

Private Sub ISearchEvent_SearchIsFinished()
    MsgBox "Search is finished"
End Sub

Private Sub cmdBrowse_Click()
Dim s As String

    s = FS.BrowseForFolder("Choose a folder", Me.hWnd)
        
    On Error Resume Next
    Drive1.Drive = FS.GetDriveName(s)
    Dir1.Path = s
End Sub

Private Sub cmdBrowse2_Click()
    txtFolder.Text = FS.BrowseForFolder("Choose a folder", Me.hWnd)
End Sub

Private Sub cmdEnumerateFiles_Click()
Dim cFiles As Files
Dim cfile As file
Dim l1 As Long, l2 As Long, l3 As Long
Dim s() As String
Dim x As Long

    '//plusieurs méthodes pour l'énumération
    
    '//METHODE 1 ==> VITESSE ELEVEE
    'on ne récupère que les noms des fichiers
    'stockés de 1 à Ubound(s()) dans la liste
    l1 = GetTickCount
    s() = FS.EnumFilesStr(txtFolder.Text, CBool(chkSub.Value))
    l1 = GetTickCount - l1
    
    '//METHODE 2 ==> VITESSE PLUS LENTE
    'on récupère les dossiers dans une collection de fichiers, mais
    'sans aller récupérer les infos complètes sur chaque fichiers
    l2 = GetTickCount
    Set cFiles = FS.EnumFiles(txtFolder.Text, CBool(chkSub.Value), False)
    l2 = GetTickCount - l2
    
    '//METHODE 3 ==> VITESSE ENCORE PLUS LENTE
    'on récupère les dossiers dans une collection de fichiers, mais
    'on va récupérer les infos complètes sur chaque fichiers
    'comme çà, on fait cFiles.Item(x).PROPERTY et la PROPERTY est à jour
    l3 = GetTickCount
    Set cFiles = FS.EnumFiles(txtFolder.Text, CBool(chkSub.Value), True)
    l3 = GetTickCount - l3
    
    MsgBox "Temps : " & vbNewLine & "Only paths (strings) : " & CStr(l1) & _
        vbNewLine & "Item collection without informations : " & CStr(l2) & _
        vbNewLine & "Item collection with informations : " & CStr(l3), _
        vbInformation, "Results"
    
    
    '//plusieurs méthodes pour l'affichage à partir de la collection
    lst.Visible = False
    lst.Clear
    
    '//METHODE 1 ==> CLASSIQUE
    'en fonction de l'Index et du Count
    l1 = GetTickCount
    For x = 1 To cFiles.Count
        lst.AddItem cFiles.Item(x).Path
    Next x
    l1 = GetTickCount - l1
    
    '//METHODE 2 ==> FOR EACH
    lst.Clear
    l2 = GetTickCount
    For Each cfile In cFiles
        lst.AddItem cfile.Path
    Next cfile
    l2 = GetTickCount - l2
    
    lst.Visible = True
    
    MsgBox "Times : " & vbNewLine & "Normal : " & CStr(l1) & vbNewLine & _
        "For Each : " & CStr(l2), vbInformation, "Résultats"

End Sub

Private Sub cmdEnumeratesFolder_Click()
Dim cFols As Folders
Dim cFol As Folder
Dim l1 As Long, l2 As Long, l3 As Long
Dim s() As String
Dim x As Long

    '//plusieurs méthodes pour l'énumération
    
    '//METHODE 1 ==> VITESSE ELEVEE
    'on ne récupère que les noms des dossiers
    'stockés de 1 à Ubound(s()) dans la liste
    l1 = GetTickCount
    s() = FS.EnumFoldersStr(txtFolder.Text, CBool(chkSub.Value))
    l1 = GetTickCount - l1
    
    '//METHODE 2 ==> VITESSE PLUS LENTE
    'on récupère les dossiers dans une collection de dossier, mais
    'sans aller récupérer les infos complètes sur chaque dossier
    l2 = GetTickCount
    Set cFols = FS.EnumFolders(txtFolder.Text, CBool(chkSub.Value), False)
    l2 = GetTickCount - l2
    
    '//METHODE 3 ==> VITESSE ENCORE PLUS LENTE
    'on récupère les dossiers dans une collection de dossier, mais
    'on va récupérer les infos complètes sur chaque dossier
    'comme çà, on fait cFols.Item(x).PROPERTY et la PROPERTY est à jour
    l3 = GetTickCount
    Set cFols = FS.EnumFolders(txtFolder.Text, CBool(chkSub.Value), True)
    l3 = GetTickCount - l3
    
    MsgBox "Temps : " & vbNewLine & "Only paths (strings) : " & CStr(l1) & _
        vbNewLine & "Item collection without informations : " & CStr(l2) & _
        vbNewLine & "Item collection with informations : " & CStr(l3), _
        vbInformation, "Results"
    
    
    '//plusieurs méthodes pour l'affichage à partir de la collection
    lst.Visible = False
    lst.Clear
    
    '//METHODE 1 ==> CLASSIQUE
    'en fonction de l'Index et du Count
    l1 = GetTickCount
    For x = 1 To cFols.Count
        lst.AddItem cFols.Item(x).Path
    Next x
    l1 = GetTickCount - l1
    
    '//METHODE 2 ==> FOR EACH
    lst.Clear
    l2 = GetTickCount
    For Each cFol In cFols
        lst.AddItem cFol.Path
    Next cFol
    l2 = GetTickCount - l2
    
    lst.Visible = True
    
    MsgBox "Times : " & vbNewLine & "Normal : " & CStr(l1) & vbNewLine & _
        "For Each : " & CStr(l2), vbInformation, "Résultats"

End Sub

Private Sub cmdInfoFolder_Click()
Dim s As String
Dim cFol As Folder
Dim l As Long

    
    l = GetTickCount
    Set cFol = FS.GetFolder(Dir1.Path)
    l = GetTickCount - l

    Me.Caption = "Informations retrieved in : " & CStr(l) & " ms"

    With cFol
        s = "Attributes : " & .Attributes
        s = s & vbNewLine & "DateCreated : " & .DateCreated
        s = s & vbNewLine & "DateLastAccessed : " & .DateLastAccessed
        s = s & vbNewLine & "DateLastModified : " & .DateLastModified
        s = s & vbNewLine & "Short Path : " & .ShortPath
        s = s & vbNewLine & "Path : " & .Path
        s = s & vbNewLine & "ParentFolderName : " & .ParentFolderName
        s = s & vbNewLine & "IsHidden : " & .IsHidden
        s = s & vbNewLine & "IsNormal : " & .IsNormal
        s = s & vbNewLine & "IsReadOnly : " & .IsReadOnly
        s = s & vbNewLine & "IsSystem : " & .IsSystem
        s = s & vbNewLine & "IsCompressed : " & .IsCompressed
        s = s & vbNewLine & "IsArchive : " & .IsArchive
        s = s & vbNewLine & "IsFolderAvailable : " & .IsFolderAvailable
    End With

    txt.Text = s

    Set cFol = Nothing
    
    pctIcon.Picture = FS.GetFolder(Dir1.Path).GetIcon(Size32)
    
End Sub


Private Sub Command1_Click()
Dim s As String
Dim s2 As String
Dim cDrive As Drive
Dim l As Long
Dim i As Long

    s2 = Drive1.Drive
    
    l = GetTickCount
    Set cDrive = FS.GetDrive(Left$(Drive1.Drive, 1))
    l = GetTickCount - l
    
    Me.Caption = "Informations retrieved in : " & CStr(l) & " ms"
    
    With cDrive
        s = "BytesPerCluster : " & CStr(.BytesPerCluster)
        s = s & vbNewLine & "BytesPerSector : " & CStr(.BytesPerSector)
        s = s & vbNewLine & "Cylinders : " & CStr(.Cylinders)
        s = s & vbNewLine & "DriveType : " & CStr(.DriveType)
        s = s & vbNewLine & "FileSystemName : " & CStr(.FileSystemName)
        s = s & vbNewLine & "FreeClusters : " & CStr(.FreeClusters)
        s = s & vbNewLine & "FreeSpace : " & CStr(.FreeSpace)
        s = s & vbNewLine & "HiddenSectors : " & CStr(.HiddenSectors)
        s = s & vbNewLine & "PartitionLength : " & CStr(.PartitionLength)
        s = s & vbNewLine & "PercentageFree : " & CStr(.PercentageFree)
        s = s & vbNewLine & "SectorPerCluster : " & CStr(.SectorPerCluster)
        s = s & vbNewLine & "SectorsPerTrack : " & CStr(.SectorsPerTrack)
        s = s & vbNewLine & "StartingOffset : " & CStr(.StartingOffset)
        s = s & vbNewLine & "strDriveType : " & (.strDriveType)
        s = s & vbNewLine & "strMediaType : " & (.strMediaType)
        s = s & vbNewLine & "TotalClusters : " & CStr(.TotalClusters)
        s = s & vbNewLine & "TotalLogicalSectors : " & CStr(.TotalLogicalSectors)
        s = s & vbNewLine & "TotalPhysicalSectors : " & CStr(.TotalPhysicalSectors)
        s = s & vbNewLine & "TotalSpace : " & CStr(.TotalSpace)
        s = s & vbNewLine & "TracksPerCylinder : " & CStr(.TracksPerCylinder)
        s = s & vbNewLine & "UsedClusters : " & CStr(.UsedClusters)
        s = s & vbNewLine & "UsedSpace : " & CStr(.UsedSpace)
        s = s & vbNewLine & "VolumeLetter : " & CStr(.VolumeLetter)
        s = s & vbNewLine & "VolumeName : " & CStr(.VolumeName)
        s = s & vbNewLine & "VolumeSerialNumber : " & CStr(.VolumeSerialNumber) & "  (" & Hex$(.VolumeSerialNumber) & ")"
        s = s & vbNewLine & "First sector : " & .ReadDriveString(0, .BytesPerSector)
        pctIcon.Picture = .GetIcon(Size32)
    End With
    
    txt.Text = s
    
    Set cDrive = Nothing
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    txtFolder.Text = Dir1.Path
    Call cmdInfoFolder_Click
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim s As String
Dim s2 As String
Dim cFil As file
Dim l As Long
    
    s2 = Dir1.Path & "\" & File1.List(File1.ListIndex)
    s2 = Replace$(s2, "\\", "\")
    
    l = GetTickCount
    Set cFil = FS.GetFile(s2)
    l = GetTickCount - l
    
    Me.Caption = "Informations retrieved in : " & CStr(l) & " ms"
    
    With cFil
        s = "AssociatedExecutableProgram : " & .AssociatedExecutableProgram
        s = s & vbNewLine & "Attributes : " & .Attributes
        s = s & vbNewLine & "DateCreated : " & .DateCreated
        s = s & vbNewLine & "DateLastAccessed : " & .DateLastAccessed
        s = s & vbNewLine & "DateLastModified : " & .DateLastModified
        s = s & vbNewLine & "FileCompressedSize : " & .FileCompressedSize
        s = s & vbNewLine & "FileExtension : " & .FileExtension
        s = s & vbNewLine & "FileSize : " & .FileSize
        s = s & vbNewLine & "ShortName : " & .ShortName
        s = s & vbNewLine & "ShortPath : " & .ShortPath
        s = s & vbNewLine & "FileType : " & .FileType
        s = s & vbNewLine & "(exe) CompanyName :" & .FileVersionInfos.CompanyName
        s = s & vbNewLine & "(exe) Copyright :" & .FileVersionInfos.Copyright
        s = s & vbNewLine & "(exe) FileDescription :" & .FileVersionInfos.FileDescription
        s = s & vbNewLine & "(exe) FileVersion :" & .FileVersionInfos.FileVersion
        s = s & vbNewLine & "(exe) InternalName :" & .FileVersionInfos.InternalName
        s = s & vbNewLine & "(exe) OriginalFileName :" & .FileVersionInfos.OriginalFileName
        s = s & vbNewLine & "(exe) ProductName :" & .FileVersionInfos.ProductName
        s = s & vbNewLine & "(exe) ProductVersion :" & .FileVersionInfos.ProductVersion
        s = s & vbNewLine & "Folder.Path : " & .Folder.Path
        s = s & vbNewLine & "Folder.Attributes : " & .Folder.Attributes
        s = s & vbNewLine & "IsHidden : " & .IsHidden
        s = s & vbNewLine & "IsNormal : " & .IsNormal
        s = s & vbNewLine & "IsReadOnly : " & .IsReadOnly
        s = s & vbNewLine & "IsSystem : " & .IsSystem
        s = s & vbNewLine & "IsCompressed : " & .IsCompressed
        s = s & vbNewLine & "IsArchive : " & .IsArchive
        s = s & vbNewLine & "IsFileAvailable : " & .IsFileAvailable
        s = s & vbNewLine & "50 first chars of file : " & vbNewLine & vbNewLine & .ReadFileString(0, 50)
    End With
    
    txt.Text = s
    
    Set cFil = Nothing
    
    pctIcon.Picture = FS.GetFile(s2).GetIcon(Size32)
    
End Sub

Private Sub Form_Load()
    Set FS = New FileSystemLibrary.FileSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FS = Nothing
End Sub

Private Sub Option3_Click()
    chkRefresh.Enabled = (Option3.Value)
End Sub

Private Sub optStr_Click()
    chkRefresh.Enabled = (Option3.Value)
End Sub

Private Sub txtFolder_Change()
    Me.cmdEnumerateFiles.Enabled = FS.FolderExists(txtFolder.Text)
     Me.cmdEnumeratesFolder.Enabled = Me.cmdEnumerateFiles.Enabled
End Sub

Private Function GetOp(cb As ComboBox) As FileSystemLibrary.Operator
Dim s As String
    
    s = cb.List(cb.ListIndex)
    
    Select Case s
        Case "="
            GetOp = [ = ]
        Case ">="
            GetOp = [ >= ]
        Case "<="
            GetOp = [ <= ]
        Case ">"
            GetOp = [ > ]
        Case "<"
            GetOp = [ < ]
    End Select
End Function
