Attribute VB_Name = "ConsoleShellModule"
'This module contains this program's main procedures.
Option Explicit

'The Microsoft Windows API structures used by this program:
Private Type INPUT_RECORD
   EventType As Long
   bKeyDown As Long
   wRepeatCount As Integer
   wVirtualKeyCode As Integer
   wVirtualScanCode As Integer
   uChar As Integer
   dwControlKeyState As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * 260
   szTypeName As String * 80
End Type

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Byte
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

'The Microsoft Windows API functions used by this program:
Private Declare Function AllocConsole Lib "Kernel32.dll" () As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateFileA Lib "Kernel32.dll" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateProcessA Lib "Kernel32.dll" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function FreeConsole Lib "Kernel32.dll" () As Long
Private Declare Function GetConsoleProcessList Lib "Kernel32.dll" (lpdwProcessList As Long, ByVal dwProcessCount As Long) As Long
Private Declare Function GetConsoleWindow Lib "Kernel32.dll" () As Long
Private Declare Function GetCurrentProcessId Lib "Kernel32.dll" () As Long
Private Declare Function GetFileSize Lib "Kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetStdHandle Lib "Kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function QueryFullProcessImageNameA Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As String, lpdwSize As Long) As Long
Private Declare Function ReadFile Lib "Kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetStdHandle Lib "Kernel32.dll" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Private Declare Function SHGetFileInfo Lib "Shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ShowWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function WaitMessage Lib "User32.dll" () As Long
Private Declare Function WriteConsoleInputA Lib "Kernel32.dll" (ByVal hConsoleInput As Long, ByRef lpBuffer As INPUT_RECORD, ByVal nLength As Long, ByRef lpNumberOfEventsWritten As Long) As Long

'The Microsoft Windows API constants used by this program:
Private Const CREATE_ALWAYS As Long = &H2&
Private Const ERROR_ALREADY_EXISTS As Long = 183
Private Const ERROR_INVALID_HANDLE As Long = 6
Private Const ERROR_INVALID_WINDOW_HANDLE As Long = 1400
Private Const ERROR_NO_MORE_FILES As Long = 18
Private Const ERROR_SEM_NOT_FOUND As Long = 187
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_TOO_MANY_POSTS As Long = 298
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Private Const FILE_FLAG_DELETE_ON_CLOSE As Long = &H4000000
Private Const FILE_SHARE_DELETE As Long = &H4&
Private Const FILE_SHARE_READ As Long = &H1&
Private Const FILE_SHARE_WRITE As Long = &H2&
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const IMAGE_NT_SIGNATURE As Long = &H4550&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const KEY_EVENT As Long = &H1
Private Const OPEN_ALWAYS As Long = &H4
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const SHGFI_EXETYPE As Long = &H2000&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const STD_ERROR_HANDLE As Long = -12
Private Const STD_INPUT_HANDLE As Long = -10
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const SW_HIDE As Long = 0

'The constants used by this program:
Private Const MAX_PATH As Long = 260       'The maximum length allowed for a directory/file path.
Private Const MAX_STRING As Long = 65535   'The maximum length allowed for a string buffer.
Private Const NO_HANDLE As Long = 0        'Indicates that there is no handle.
Public Const NO_PROCESSES_MESSAGE As String = "There are no processes attached to the console."   'Error message text.

'This procedure checks the console for output and gives the command to display it when present.
Private Sub CheckForConsoleOutput(TextBox As Object)
On Error GoTo ErrorTrap
Dim BytesRead As Long
Dim ConsoleOutput As String
Dim OutputFileHandle As Long
Dim OutputSize As Long
Static PreviousOutputSize As Long

   OutputFileHandle = OutputFile()
   If Not (OutputFileHandle = NO_HANDLE Or OutputFileHandle = INVALID_HANDLE_VALUE) Then
      OutputSize = CheckForError(GetFileSize(OutputFile(), CLng(0)), , ERROR_INVALID_WINDOW_HANDLE)
      If OutputSize > PreviousOutputSize Then
      
         Do While DoEvents() > 0
            ConsoleOutput = String$(OutputBufferSize(), vbNullChar)
            CheckForError ReadFile(OutputFile(), ConsoleOutput, Len(ConsoleOutput), BytesRead, CLng(0))
            If BytesRead <= 0 Then Exit Do
            ConsoleOutput = Left$(ConsoleOutput, BytesRead)
            Display ConsoleOutput, TextBox
         Loop
      
         PreviousOutputSize = OutputSize
      End If
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long, Optional ExtraInformation As String = vbNullString, Optional Ignored1 As Long = ERROR_SUCCESS, Optional Ignored2 As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   On Error GoTo ErrorTrap
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored1 Or ErrorCode = Ignored2) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_ARGUMENT_ARRAY Or FORMAT_MESSAGE_FROM_SYSTEM, CLng(0), ErrorCode, CLng(0), Description, Len(Description), StrPtr(StrConv(ExtraInformation, vbFromUnicode)))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCrLf
      If MsgBox(Message, vbExclamation Or vbOKCancel) = vbCancel Then End
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure closes the processes attached to the console.
Private Sub CloseConsoleProcesses()
On Error GoTo ErrorTrap
Dim Index As Long
Dim ProcessHandle As Long
Dim ProcessIDList() As Long

   ProcessIDList() = GetConsoleProcessIDs()
   If UBound(ProcessIDList()) > 0 Then
      For Index = LBound(ProcessIDList()) To UBound(ProcessIDList())
         If Not ProcessIDList(Index) = GetCurrentProcessId() Then
            ProcessHandle = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(True), ProcessIDList(Index)))
            CheckForError TerminateProcess(ProcessHandle, CLng(0))
            CheckForError CloseHandle(ProcessHandle)
         End If
      Next Index
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure creates a console and sets its output handles.
Private Sub CreateConsole(OutputHandle As Long)
On Error GoTo ErrorTrap
   If CheckForError(GetConsoleWindow(), , ERROR_INVALID_HANDLE) = NO_HANDLE Then
      CheckForError AllocConsole(), , ERROR_SEM_NOT_FOUND
      CheckForError ShowWindow(GetConsoleWindow(), SW_HIDE), , ERROR_SEM_NOT_FOUND
      CheckForError SetStdHandle(STD_ERROR_HANDLE, OutputHandle), , ERROR_SEM_NOT_FOUND
      CheckForError SetStdHandle(STD_OUTPUT_HANDLE, OutputHandle), , ERROR_SEM_NOT_FOUND
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure adds text to the specified text box.
Public Sub Display(NewText As String, TextBox As Object)
On Error GoTo ErrorTrap
Dim Position As Long
Dim Text As String

   If InStr(NewText, vbCrLf) = 0 Then
      If InStr(NewText, vbCr) Then
         NewText = Replace(NewText, vbCr, vbCrLf)
      ElseIf InStr(NewText, vbLf) Then
         NewText = Replace(NewText, vbLf, vbCrLf)
      End If
   End If
   
   NewText = Escape(NewText)
   
   With TextBox
      For Position = 1 To Len(NewText) Step MAX_STRING
         If Len(Mid$(NewText, Position)) < MAX_STRING Then
            Text = Mid$(NewText, Position)
         Else
            Text = Mid$(NewText, Position, MAX_STRING)
         End If
   
         If Len(.Text & NewText) > MAX_STRING Then .Text = Mid$(.Text, Len(NewText))
                     
         .SelLength = 0
         .SelStart = Len(.Text)
         .SelText = .SelText & NewText
      Next Position
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure converts non-displayable characters in the specified text to escape sequences.
Private Function Escape(Text As String, Optional EscapeCharacter As String = "/", Optional EscapeLineBreaks As Boolean = False) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Escaped As String
Dim Index As Long
Dim NextCharacter As String

   Escaped = vbNullString
   Index = 1
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = EscapeCharacter Then
         Escaped = Escaped & String$(2, EscapeCharacter)
      ElseIf Character = vbTab Or Character >= " " Then
         Escaped = Escaped & Character
      ElseIf Character & NextCharacter = vbCrLf And Not EscapeLineBreaks Then
         Escaped = Escaped & vbCrLf
         Index = Index + 1
      Else
         Escaped = Escaped & EscapeCharacter & String$(2 - Len(Hex$(Asc(Character))), "0") & Hex$(Asc(Character))
      End If
      Index = Index + 1
   Loop
   
EndRoutine:
   Escape = Escaped
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure returns a list of process ID's of processes attached to the console.
Public Function GetConsoleProcessIDs() As Long()
On Error GoTo ErrorTrap
Dim ProcessCount As Long
Dim ProcessIDList() As Long

   ReDim ProcessIDList(0 To 0) As Long
   
   If Not CheckForError(GetConsoleWindow()) = NO_HANDLE Then
      ReDim ProcessIDList(1 To 1) As Long
      ProcessCount = CheckForError(GetConsoleProcessList(ProcessIDList(LBound(ProcessIDList())), UBound(ProcessIDList())))
      
      ReDim ProcessIDList(LBound(ProcessIDList()) To ProcessCount) As Long
      CheckForError GetConsoleProcessList(ProcessIDList(1), UBound(ProcessIDList()))
   End If

EndRoutine:
   GetConsoleProcessIDs = ProcessIDList()
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the path of the specified process.
Private Function GetProcessPath(ProcessH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim Path As String
Dim ReturnValue As Long

   Path = String$(MAX_PATH, vbNullChar)
   Length = Len(Path)
   ReturnValue = CheckForError(QueryFullProcessImageNameA(ProcessH, CLng(0), Path, Length))
   If Not ReturnValue = 0 Then Path = Left$(Path, Length)
EndRoutine:
   GetProcessPath = Path
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   
   On Error Resume Next
   MsgBox Description & vbCr & "Error code: " & CStr(ErrorCode), vbExclamation
End Sub


'This procedure manages the console's input handle.
Private Function InputHandle() As Long
On Error GoTo ErrorTrap
Static CurrentInputHandle As Long

   If CurrentInputHandle = NO_HANDLE Then CurrentInputHandle = CheckForError(GetStdHandle(STD_INPUT_HANDLE))
   
EndRoutine:
   InputHandle = CurrentInputHandle
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure checks whether the specified executable is a console process and returns the result.
Private Function IsConsoleProcess(Path As String) As Boolean
On Error GoTo ErrorTrap
Dim FileInformation As SHFILEINFO
Dim ReturnValue As Long

   ReturnValue = CheckForError(SHGetFileInfo(Path, CLng(0), FileInformation, Len(FileInformation), SHGFI_EXETYPE))
   
EndRoutine:
   IsConsoleProcess = (((ReturnValue And &HFFFF) = IMAGE_NT_SIGNATURE) And (((ReturnValue / &H10000) And &HFFFF) = &H0&))
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure gives lists the processes attached to the console.
Public Sub ListProcesses(TextBox As Object)
On Error GoTo ErrorTrap
Dim ProcessHandle As Long
Dim ProcessId As Variant
Dim ProcessIds() As Long

   ProcessIds() = GetConsoleProcessIDs()
   If UBound(GetConsoleProcessIDs()) = 0 Then
      MsgBox NO_PROCESSES_MESSAGE, vbExclamation
   Else
      Display vbCrLf, TextBox
      Display "Processes attached to this program:" & vbCrLf, TextBox
      For Each ProcessId In ProcessIds()
         ProcessHandle = CheckForError(OpenProcess(PROCESS_QUERY_INFORMATION, CLng(False), ProcessId))
         If Not ProcessHandle = NO_HANDLE Then
            Display CStr(ProcessId) & vbTab & GetProcessPath(CLng(ProcessHandle)) & vbCrLf, TextBox
            CheckForError CloseHandle(ProcessHandle)
         End If
      Next ProcessId
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure initializes this program.
Private Sub Main()
On Error GoTo ErrorTrap
   CreateConsole OutputFile()
   OutputBufferSize NewOutputBufferSize:=10000
   StartProcess Command$()
   InterfaceWindow.Show
   
   Do While DoEvents() > 0
      CheckForConsoleOutput InterfaceWindow.ConsoleOutputBox
      CheckForError WaitMessage(), , ERROR_INVALID_WINDOW_HANDLE
   Loop
   
EndRoutine:
   Quit
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure manages the output buffer size.
Public Function OutputBufferSize(Optional NewOutputBufferSize As Long = Empty) As Long
On Error GoTo ErrorTrap
Static CurrentOutputBufferSize As Long

   If Not NewOutputBufferSize = Empty Then CurrentOutputBufferSize = NewOutputBufferSize
   
EndRoutine:
   OutputBufferSize = CurrentOutputBufferSize
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the console output file and returns its read and write handles.
Private Function OutputFile(Optional ByRef OutputWriteHandle As Long = NO_HANDLE) As Long
On Error GoTo ErrorTrap
Dim Path As String
Dim Security As SECURITY_ATTRIBUTES
Static CurrentOutputReadHandle As Long
Static CurrentOutputWriteHandle As Long
   
   If CurrentOutputReadHandle = NO_HANDLE Then
      Path = App.Path
      If Not Right$(Path, 1) = "\" Then Path = Path & "\"
      Path = Path & "Console Output"
      
      Security.bInheritHandle = True
      Security.nLength = Len(Security)
      
      CurrentOutputWriteHandle = CheckForError(CreateFileA(Path, GENERIC_WRITE, FILE_SHARE_DELETE Or FILE_SHARE_READ Or FILE_SHARE_WRITE, Security, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_DELETE_ON_CLOSE, CLng(0)), , ERROR_ALREADY_EXISTS)
      
      If CurrentOutputWriteHandle = INVALID_HANDLE_VALUE Then
         CurrentOutputReadHandle = INVALID_HANDLE_VALUE
      Else
         CurrentOutputReadHandle = CheckForError(CreateFileA(Path, GENERIC_READ, FILE_SHARE_DELETE Or FILE_SHARE_READ Or FILE_SHARE_WRITE, Security, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, CLng(0)), , ERROR_ALREADY_EXISTS)
      End If
      
      If CurrentOutputReadHandle = INVALID_HANDLE_VALUE Or CurrentOutputWriteHandle = INVALID_HANDLE_VALUE Then MsgBox "Could not create a console output capture file.", vbExclamation
   End If
   
EndRoutine:
   OutputWriteHandle = CurrentOutputWriteHandle
   OutputFile = CurrentOutputReadHandle
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function




'This procedure closes this program.
Private Sub Quit()
On Error GoTo ErrorTrap
Dim OutputReadHandle As Long
Dim OutputWriteHandle As Long

   CloseConsoleProcesses
   WaitForConsoleProcessesToClose
   
   OutputReadHandle = OutputFile(OutputWriteHandle)
   
   If Not CheckForError(GetConsoleWindow()) = NO_HANDLE Then CheckForError FreeConsole(), , ERROR_INVALID_HANDLE
   If Not OutputFile() = NO_HANDLE Then CheckForError CloseHandle(OutputReadHandle), , ERROR_INVALID_HANDLE
   If Not OutputWriteHandle = NO_HANDLE Then CheckForError CloseHandle(OutputWriteHandle), , ERROR_INVALID_HANDLE
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure starts the specified process.
Public Sub StartProcess(Path As String)
On Error GoTo ErrorTrap
Dim OutputWriteHandle As Long
Dim ProcessInformation As PROCESS_INFORMATION
Dim ReturnValue As Long
Dim Security As SECURITY_ATTRIBUTES
Dim StartUpInformation As STARTUPINFO

   If Not Path = vbNullString Then
      If Left$(Path, 1) = """" Then Path = Mid$(Path, 2)
      If Right$(Path, 1) = """" Then Path = Left$(Path, Len(Path) - 1)

      OutputFile OutputWriteHandle
      
      Security.bInheritHandle = True
      Security.nLength = Len(Security)
      StartUpInformation.cb = Len(StartUpInformation)
      StartUpInformation.dwFlags = STARTF_USESTDHANDLES
      StartUpInformation.hStdError = OutputWriteHandle
      StartUpInformation.hStdInput = InputHandle()
      StartUpInformation.hStdOutput = OutputWriteHandle
      
      ReturnValue = CheckForError(CreateProcessA(vbNullString, Path, Security, Security, CLng(True), CLng(0), CLng(0), vbNullString, StartUpInformation, ProcessInformation), Path, ERROR_NO_MORE_FILES, ERROR_TOO_MANY_POSTS)
      If Not ReturnValue = 0 Then
         Path = GetProcessPath(ProcessInformation.hProcess)
         If Not IsConsoleProcess(Path) Then MsgBox """" & Path & """ is not a console application and cannot interact with " & App.Title & ".", vbExclamation
      End If
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure converts any escape sequences in the specified text to characters.
Public Function Unescape(Text As String, Optional EscapeCharacter As String = "/", Optional ErrorAt As Long = 0) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Hexadecimals As String
Dim Index As Long
Dim NextCharacter As String
Dim Unescaped As String

   ErrorAt = 0
   Index = 1
   Unescaped = vbNullString
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = EscapeCharacter Then
         If NextCharacter = EscapeCharacter Then
            Unescaped = Unescaped & Character
            Index = Index + 1
         Else
            Hexadecimals = UCase$(Mid$(Text, Index + 1, 2))
            If Len(Hexadecimals) = 2 Then
               If Left$(Hexadecimals, 1) = "0" Then Hexadecimals = Right$(Hexadecimals, 1)
      
               If UCase$(Hex$(CLng(Val("&H" & Hexadecimals & "&")))) = Hexadecimals Then
                  Unescaped = Unescaped & Chr$(CLng(Val("&H" & Hexadecimals & "&")))
                  Index = Index + 2
               Else
                  ErrorAt = Index
                  Exit Do
               End If
            Else
               ErrorAt = Index
               Exit Do
            End If
         End If
      Else
         Unescaped = Unescaped & Character
      End If
      Index = Index + 1
   Loop
   
EndRoutine:
   Unescape = Unescaped
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure waits for the processes attached to the console to close.
Private Sub WaitForConsoleProcessesToClose()
On Error GoTo ErrorTrap

   Do While UBound(GetConsoleProcessIDs()) > 1
      DoEvents
   Loop
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure writes the specified text to the console.
Public Sub WriteToConsole(Text As String)
On Error GoTo ErrorTrap
Dim BytesWritten As Long
Dim Index As Long
Dim InputRecord() As INPUT_RECORD

   ReDim InputRecord(1 To Len(Text & vbCrLf)) As INPUT_RECORD
   For Index = LBound(InputRecord()) To UBound(InputRecord())
      With InputRecord(Index)
         .bKeyDown = True
         .dwControlKeyState = 0
         .EventType = KEY_EVENT
         .uChar = CInt(Asc(Mid$(Text & vbCrLf, Index, 1)))
         .wRepeatCount = 0
         .wVirtualKeyCode = 0
         .wVirtualScanCode = 0
      End With
   Next Index
   
   CheckForError WriteConsoleInputA(InputHandle(), InputRecord(1), UBound(InputRecord()), BytesWritten)
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


