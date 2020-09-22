VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Folder Monitor Demo"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      ExtentX         =   873
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   420
      Left            =   4440
      TabIndex        =   10
      Top             =   5880
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Help"
      Height          =   420
      Left            =   3000
      TabIndex        =   9
      Top             =   5880
      Width           =   1350
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   0
   End
   Begin VB.CheckBox chkIncludeSubs 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Include subdirectories"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2760
      TabIndex        =   8
      Top             =   390
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2100
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2340
      Left            =   120
      TabIndex        =   4
      Top             =   750
      Width           =   7845
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   420
      Left            =   6600
      TabIndex        =   3
      Top             =   5880
      Width           =   1350
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop Monitoring"
      Height          =   420
      Left            =   1560
      TabIndex        =   2
      Top             =   5880
      Width           =   1350
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start &Monitoring"
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   1350
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   3465
      Width           =   7845
   End
   Begin VB.Label Label3 
      Caption         =   "some_clever_name@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Files changed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   3210
      Width           =   4230
   End
   Begin VB.Label Label1 
      Caption         =   "Folder to monitor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFolder2Monitor               As String
Dim IncludeSubs                     As Boolean

Dim WithEvents objFolderMonitor     As FolderMonitor
Attribute objFolderMonitor.VB_VarHelpID = -1
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Sub cmdAbout_Click()

    ShellAbout Me.hWnd, "Folder Monitor Demonstration", "Author: Darrell Koller," & vbCrLf & "E-mail: some_clever_name@hotmail.com", Me.Icon

End Sub

Private Sub cmdExit_Click()

    cmdStop_Click
    End

End Sub

Private Sub cmdStart_Click()

    Dim objCurrentControl As Object
    For Each objCurrentControl In Me.Controls
        If TypeName(objCurrentControl) <> "WebBrowser" Then
            objCurrentControl.Enabled = False
        End If
    Next
    cmdStop.Enabled = True
    Label2.Enabled = True
    Me.Refresh

    Set objFolderMonitor = New FolderMonitor
    objFolderMonitor.WaitTime = 100
    objFolderMonitor.IncludeSubFolders = chkIncludeSubs.Value = vbChecked
    objFolderMonitor.Attributes = True
    objFolderMonitor.AddFolder strFolder2Monitor
    
    Timer1.Enabled = True
    objFolderMonitor.StartMonitoring

End Sub

Private Sub cmdStop_Click()

    Dim objCurrentControl As Object

    If Not objFolderMonitor Is Nothing Then objFolderMonitor.StopMonitoring
    Timer1.Enabled = False
    
    For Each objCurrentControl In Me.Controls
        If TypeName(objCurrentControl) <> "WebBrowser" Then
            objCurrentControl.Enabled = True
        End If
    Next
    Me.Refresh

End Sub

Private Sub Command1_Click()

    Dim strCommand As String
    strCommand = GetAssociatedExecutable(App.Path & "\FOLDER_MONITOR_DEMO.rtf") & " """ & App.Path & "\Folder_Monitor_Demo.rtf"""
    
    Shell strCommand, vbNormalFocus
    
End Sub

Private Sub Dir1_click()

    strFolder2Monitor = Dir1.List(Dir1.ListIndex)
    
End Sub

Private Sub Drive1_Change()

    Dir1.Path = Left(Drive1.Drive, 1) & ":\"

End Sub

Private Sub Form_Load()

    Dir1.Path = Left(Drive1.Drive, 1) & ":\"
    strFolder2Monitor = Dir1.List(Dir1.ListIndex)

End Sub

Private Sub Label3_Click()

    WebBrowser1.Navigate "mailto:some_clever_name@hotmail.com?subject=About Folder Monitor Demo"

End Sub

Sub objFolderMonitor_ChangeOccurred()

'    UpdateList
    
End Sub

Private Sub Timer1_Timer()

    UpdateList

End Sub

Private Sub UpdateList()

    Dim aTemp As Variant
    Dim i As Integer
    Dim intFileCount As Integer
    
    aTemp = objFolderMonitor.ChangedList
    
    intFileCount = aLen(aTemp)
    
    For i = 0 To intFileCount - 1
         List1.AddItem aTemp(i)
    Next i

End Sub

