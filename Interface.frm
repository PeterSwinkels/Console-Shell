VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form InterfaceWindow 
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "Interface.frx":0000
   ScaleHeight     =   13.5
   ScaleMode       =   4  'Character
   ScaleWidth      =   39
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FontDialog 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton EnterButton 
      Caption         =   "&Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Send the input to a console application."
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox ConsoleInputBox 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Specify the input for a console application here."
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox ConsoleOutputBox 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Display's a console application's output."
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu ListProcessesMenu 
         Caption         =   "&List Processes"
         Shortcut        =   ^L
      End
      Begin VB.Menu StartProcessMenu 
         Caption         =   "&Start Process"
         Shortcut        =   ^S
      End
      Begin VB.Menu ProgramMainMenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu OptionsMainMenu 
      Caption         =   "&Options"
      Begin VB.Menu ClearOutputMenu 
         Caption         =   "&Clear Output"
         Shortcut        =   {F1}
      End
      Begin VB.Menu OutputBufferSizeMenu 
         Caption         =   "&Output Buffer Size"
         Shortcut        =   {F2}
      End
      Begin VB.Menu RepeatInputMenu 
         Caption         =   "&Repeat Input"
         Shortcut        =   {F3}
      End
      Begin VB.Menu SelectFontMenu 
         Caption         =   "&Select Font"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit



'This procedure stores the last input by the user and returns it.
Private Function LastInput(Optional NewLastInput As String = vbNullString) As String
On Error GoTo ErrorTrap
Static CurrentLastInput As String

   If Not NewLastInput = vbNullString Then CurrentLastInput = NewLastInput
   
EndRoutine:
   LastInput = CurrentLastInput
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the last started process and gives the command to start any new process specified.
Private Function LastStartedProcess(Optional NewStartedProcess As String = vbNullString) As String
On Error GoTo ErrorTrap
Static CurrentStartedProcess As String

   If Not NewStartedProcess = vbNullString Then
      CurrentStartedProcess = NewStartedProcess
      StartProcess NewStartedProcess
   End If
   
EndRoutine:
   LastStartedProcess = CurrentStartedProcess
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure clears the console output box.
Private Sub ClearOutputMenu_Click()
On Error GoTo ErrorTrap
   ConsoleOutputBox.Text = vbNullString
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to write the user's input to the console.
Private Sub EnterButton_Click()
On Error GoTo ErrorTrap
Dim ErrorAt As Long

   If Not ConsoleInputBox.Text = vbNullString Then
      LastInput NewLastInput:=Unescape(ConsoleInputBox.Text, , ErrorAt)
      If ErrorAt > 0 Then
         MsgBox "Bad escape sequence at character #" & CStr(ErrorAt) & ".", vbExclamation
      Else
         If UBound(GetConsoleProcessIDs()) > 1 Then
            ConsoleInputBox.Text = vbNullString
            Display LastInput() & vbCrLf, ConsoleOutputBox
            WriteToConsole LastInput()
         Else
            MsgBox NO_PROCESSES_MESSAGE, vbExclamation
         End If
      End If
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap

   With App
      Me.Caption = .Title & " - by: " & .CompanyName & ", v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision)
   End With
   
   Me.Width = Screen.Width / 1.1
   Me.Height = Screen.Height / 1.1
   
   LastInput NewLastInput:=vbNullString
   LastStartedProcess NewStartedProcess:=vbNullString
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure adjusts the size and position of the interface object's to the window's new size.
Private Sub Form_Resize()
On Error Resume Next
   ConsoleOutputBox.Height = Me.ScaleHeight - ConsoleInputBox.Height
   ConsoleOutputBox.Left = 0
   ConsoleOutputBox.Top = 0
   ConsoleOutputBox.Width = Me.ScaleWidth
   
   ConsoleInputBox.Left = 0
   ConsoleInputBox.Width = Me.ScaleWidth - EnterButton.Width - 2
   ConsoleInputBox.Top = Me.ScaleHeight - ConsoleInputBox.Height
   
   EnterButton.Left = ConsoleInputBox.Width + 1
   EnterButton.Top = (Me.ScaleHeight - (ConsoleInputBox.Height / 2)) - (EnterButton.Height / 2)
End Sub



'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   MsgBox App.Comments, vbInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to list the processes attached to the console.
Private Sub ListProcessesMenu_Click()
On Error GoTo ErrorTrap
   ListProcesses ConsoleOutputBox
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure requests the user to specify a new output buffer size.
Private Sub OutputBufferSizeMenu_Click()
On Error GoTo ErrorTrap
Dim NewOutputBufferSize As Long

   NewOutputBufferSize = CLng(Val(InputBox$("The maximum number of bytes to read at once:", , OutputBufferSize())))
   If Not NewOutputBufferSize = Empty Then OutputBufferSize NewOutputBufferSize:=NewOutputBufferSize
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to close this program.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure repeats the previous user console input.
Private Sub RepeatInputMenu_Click()
On Error GoTo ErrorTrap
   ConsoleInputBox.Text = LastInput()
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure requests the user to select a font.
Private Sub SelectFontMenu_Click()
On Error GoTo ErrorTrap
   With FontDialog
      .FontName = ConsoleInputBox.Font
      .FontSize = ConsoleInputBox.Font.Size
      .ShowFont
      
      ConsoleInputBox.Font = .FontName
      ConsoleInputBox.Font.Size = .FontSize
      ConsoleOutputBox.Font = ConsoleInputBox.Font
      ConsoleOutputBox.Font.Size = ConsoleInputBox.Font.Size
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to start the process specified by the user.
Private Sub StartProcessMenu_Click()
On Error GoTo ErrorTrap
Dim Path As String

   Path = InputBox$("Process path:", , LastStartedProcess())
   LastStartedProcess NewStartedProcess:=Path
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub




