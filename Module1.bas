Attribute VB_Name = "Module1"
Option Explicit

'Constants
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_MINIMIZE = 6
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const IncludeStartComments As String = "'+++ START OF INCLUDED LINES ------------------------------------------------"
Private Const IncludeStopComments As String = "'+++ END OF INCLUDED LINES --------------------------------------------------"

'Types
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
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

'Declarations
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle&, ByVal dwMilliseconds&) As Long
Private Declare Function CreateProcessA Lib "Kernel32" (ByVal lpApplicationName&, ByVal lpCommandLine$, ByVal _
   lpProcessAttributes&, ByVal lpThreadAttributes&, ByVal bInheritHandles&, ByVal dwCreationFlags&, _
   ByVal lpEnvironment&, ByVal lpCurrentDirectory&, lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "Kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Variables
Private PathToTemp As String
Private PathToVB As String
Private PathToTempVBP As String
Private PathToErrors As String

'Methods
'------------------------------------------------------------------------------------------
'Startup and Shutdown Routines
'------------------------------------------------------------------------------------------

Public Sub Main()

    'Initialize the global variables
    PathToVB = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe"
    PathToTemp = App.Path & "\Temp"
    PathToErrors = PathToTemp & "\Compiler Output.txt"
    PathToTempVBP = PathToTemp & "\Include.vbp"
    
    'If the pathname to VB6 is wrong, try just executing it.
    If Dir(PathToVB, vbNormal Or vbArchive) = "" Then
        PathToVB = "VB6.exe"
    End If
    
    'Create the temporary folder if it doesn't exist.
    If Dir(PathToTemp, vbDirectory) = "" Then
        MkDir PathToTemp
    End If
    
    'Show the main form
    Form1.Show
End Sub

'------------------------------------------------------------------------------------------
'Compile Routines
'------------------------------------------------------------------------------------------
Public Sub CompileProject(ByVal VBP As String)
    Dim s As String, RetVal As Long, OpenVB As Long, CompileErrors As String
    Screen.MousePointer = vbHourglass
    
    'Empty out the folder
    Status "Removing temporary files"
    KillFolder PathToTemp
    
    'Copy over all the files to the temporary folder
    Status "Copying over files in project to temporary folder"
    CopyFilesInVBP VBP
    
    'Compile the VBP
    Status "Compiling the project, please wait..."
    s = PathToVB & " /make """ & PathToTempVBP & """" & " /out " & """" & PathToErrors & """" & " /outdir " & """" & PathToTemp & """"
    RetVal = ExecCmd(s, "")
    Screen.MousePointer = vbNormal
    
    'Check return code
    LoadFile PathToErrors, CompileErrors
    If RetVal = 0 Then
        'Program successfully compiled
        'Remove the include lines, and notify the user.
        Status "EXE was placed in " & PathToTemp
        MsgBox CompileErrors
    Else
        'Program wouldn't compile!
        'Open the VBP with the #included lines if desired
        Status "UNSUCCESSFUL COMPILE."
        OpenVB = MsgBox("Couldn't compile the program for the following reason(s):" & vbCrLf & _
        CompileErrors, vbCritical)
    End If

End Sub

Private Function ExecCmd(ByVal CmdLine As String, ByVal TitleOfWindow As String) As Long
    'Shell and wait for an external program
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim RetVal As Long

    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    start.dwFlags = STARTF_USESHOWWINDOW
    start.wShowWindow = SW_MINIMIZE
    start.lpTitle = TitleOfWindow

    ' Start the shelled application:
    RetVal = CreateProcessA(0&, CmdLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    ' Wait for the shelled application to finish
    RetVal = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, RetVal)
    Call CloseHandle(proc.hProcess)
    ExecCmd = RetVal
End Function

Private Sub CopyFilesInVBP(ByVal VBP As String)
    'Copies all the files referenced by the VB project to the temporary directory.
    Dim s() As String, OrigText As String, i As Long
    Dim Pathname As String, Suffix As String
    Dim OldShortName As String, OldLongName As String
    Dim NewShortName As String, NewLongName As String
    
    'Load the VBP file into an array
    If LoadFile(VBP, OrigText) = False Then
        MsgBox "Fatal error! Couldn't find file " & VBP
        End
    End If
    s = Split(OrigText, vbCrLf)
    
    'Save the path of the VBP
    Pathname = PathPart(VBP)
    
    'Search for files referenced in the VBP
    For i = 0 To UBound(s, 1)
        If IsIn(FirstPart(s(i), "="), "Module,Class,Form,RelatedDoc,ResFile32") Then
            'Copy the file
            If InStr(1, s(i), ";") > 0 Then
                OldShortName = Trim(LastPart(s(i), ";"))
            Else
                OldShortName = Trim(LastPart(s(i), "="))
            End If
            OldShortName = Replace(OldShortName, """", "")
            OldLongName = GetFullPathname(OldShortName, Pathname)
            NewLongName = PathToTemp & "\" & LastPart(OldLongName)
            NewShortName = LastPart(NewLongName)
            s(i) = Replace(s(i), OldShortName, NewShortName)
            FileCopy OldLongName, NewLongName
            
            'Copy related files
            Suffix = Right$(OldLongName, 4)
            If IsIn(Suffix, ".frm") = True Then
                OldLongName = Replace(OldLongName, ".frm", ".frx")
                If Dir(OldLongName, vbNormal Or vbArchive) > "" Then
                    NewLongName = Replace(NewLongName, ".frm", ".frx")
                    FileCopy OldLongName, NewLongName
                End If
            End If
            
            'Add the include lines
            If IsIn(Suffix, ".frm,.bas,.cls") Then
                AddIncludeLines OldShortName, OldLongName, NewShortName, NewLongName
            End If
            
        End If
    Next i
    
    'Save the modified VBP file to the temporary directory as well.
    SaveFile PathToTempVBP, Join(s, vbCrLf)
    
End Sub

Private Sub AddIncludeLines(ByVal OldShortName As String, ByVal OldLongName As String, _
                            ByVal NewShortName As String, ByVal NewLongName As String)
    Dim s As String, s2() As String, i As Long, c As Long
    Dim IncludeFilename As String, IncludeText As String, Suffix As String
    
    'Load the lines
    If LoadFile(NewLongName, s) = False Then
        MsgBox "File " & NewLongName & " not found!", vbCritical
        End
    End If
    s2 = Split(s, vbCrLf)
    
    'Check one by one for #include lines
    For i = LBound(s2, 1) To UBound(s2, 1)
        c = InStr(1, s2(i), "#INCLUDE ", vbTextCompare)
        If c > 0 Then
            IncludeFilename = Right$(s2(i), Len(s2(i)) - c)
            IncludeFilename = Right$(IncludeFilename, Len(IncludeFilename) - 8)
            If InStr(1, IncludeFilename, "\", vbTextCompare) = 0 Then
                IncludeFilename = PathPart(OldLongName) & "\" & IncludeFilename
                If InStr(1, IncludeFilename, ".") > 0 Then
                    'The person has already included a file extension
                Else
                    'Assume a file extension of .txt
                    IncludeFilename = IncludeFilename & ".txt"
                End If
            End If
            'Load up the include text
            If LoadFile(IncludeFilename, IncludeText) = False Then
                MsgBox "Unable to find #INCLUDE file " & IncludeFilename & "!", vbCritical
                End
            End If
            s2(i) = IncludeStartComments & vbCrLf & IncludeText & vbCrLf & IncludeStopComments
        End If
    Next i
    SaveFile NewLongName, Join(s2, vbCrLf)
End Sub

'------------------------------------------------------------------------------------------
'File Routines
'------------------------------------------------------------------------------------------

Public Function LoadFile(ByVal Filename As String, ByRef ReturnedLines As String) As Boolean
    Dim s As String, FileNo As Integer, b() As Byte
    
    'Make sure the file exists
    If Dir(Filename, vbNormal Or vbArchive) = "" Then
        LoadFile = False
        Exit Function
    End If
    
    'Open the file
    FileNo = FreeFile
    Open Filename For Binary Access Read As #FileNo
    ReDim b(1 To LOF(FileNo))
    Get #FileNo, , b
    Close #FileNo
    
    'Convert to an array of strings
    s = StrConv(b, vbUnicode)
    ReturnedLines = s

    'Return success
    LoadFile = True
End Function

Private Function SaveFile(ByVal Filename As String, ByRef AllLines As String) As Boolean
    Dim s As String, FileNo As Integer
    
    'Open the file
    FileNo = FreeFile
    Open Filename For Output As #FileNo
    Print #FileNo, AllLines
    Close #FileNo
    
    'Return success
    SaveFile = True
End Function

Private Sub KillFolder(ByVal Pathname As String)
    'Deletes all files in a folder.
    Dim Filename As String
    If Right$(Pathname, 1) = "\" Then
        Pathname = Left$(Pathname, Len(Pathname) - 1)
    End If
    Filename = Dir(Pathname & "\*.*", vbNormal Or vbArchive)
    Do
        If Filename = "" Then
            Exit Do
        End If
        Kill Pathname & "\" & Filename
        Filename = Dir()
    Loop
End Sub

'------------------------------------------------------------------------------------------
'String Manipulation Routines
'------------------------------------------------------------------------------------------

Private Function GetFullPathname(ByVal Filename As String, ByVal Pathname As String) As String
    'Converts pathnames like "..\Shared Code\clsCryptAPI.cls" to their full equivalent
    
    If InStr(1, Filename, "..\", vbTextCompare) = 1 Then
        'Replace dots with higher level path
        Do While InStr(1, Filename, "..\", vbTextCompare) = 1
            Pathname = PathPart(Pathname)
            Filename = Right$(Filename, Len(Filename) - 3)
        Loop
        Filename = Pathname & "\" & Filename
    ElseIf InStr(1, Filename, "\", vbTextCompare) Then
        'It is a pathname, don't touch it
    Else
        'Prefix the pathname to the filename
        Filename = Pathname & "\" & Filename
    End If
    
    GetFullPathname = Filename
End Function

Private Function PathPart(ByVal s As String, Optional ByVal Delim As String = "\")
    'Returns all but the last part of a pathname
    Dim c As Long
    c = InStrRev(s, Delim, , vbTextCompare)
    If c = 0 Then
        PathPart = s
    Else
        PathPart = Left$(s, c - 1)
    End If
End Function

Private Function FirstPart(ByVal s As String, Optional ByVal Delim As String = "\") As String
    'Returns the first word of a string
    Dim c As Long
    c = InStr(1, s, Delim, vbTextCompare)
    If c > 0 Then
        FirstPart = Left$(s, c - 1)
    Else
        FirstPart = s
    End If
End Function

Private Function LastPart(ByVal s As String, Optional ByVal Delim As String = "\") As String
    'Returns the last word of a string
    Dim c As Long
    c = InStrRev(s, Delim, , vbTextCompare)
    If c > 0 Then
        LastPart = Right$(s, Len(s) - c - Len(Delim) + 1)
    Else
        LastPart = s
    End If
End Function

Private Function IsIn(ByVal CheckWord As String, ByVal ListOfValidWords As String, Optional CaseSensitive As Boolean) As Boolean
    'Checks to see if a word is in contained in a series of other words.
    
    'Ensure non-blank parameters
    If Len(CheckWord) = 0 Or Len(ListOfValidWords) = 0 Then
        Exit Function
    End If
    
    'Check for case sensitivity
    If CaseSensitive = False Then
        CheckWord = LCase$(CheckWord)
        ListOfValidWords = LCase$(ListOfValidWords)
    End If
    
    'Ensure list of valid words starts and ends with comma
    If Right$(ListOfValidWords, 1) <> "," Then
        ListOfValidWords = ListOfValidWords & ","
    End If
    If Left$(ListOfValidWords, 1) <> "," Then
        ListOfValidWords = "," & ListOfValidWords
    End If
    
    'Finally, perform check WITHOUT vbTextCompare to accommodate case!
    If InStr(1, ListOfValidWords, "," & CheckWord & ",") > 0 Then
        IsIn = True
    End If
    
End Function

'------------------------------------------------------------------------------------------
'Misc Routines
'------------------------------------------------------------------------------------------
Private Sub Status(ByVal s As String)
    Form1.StatusBar1.SimpleText = s
    Form1.StatusBar1.Refresh
End Sub

Public Sub OpenDocument(ByVal Filename As String)
    'Open the specified document with the default application.
    ShellExecute 0, "Open", Filename, "", "", 1
End Sub

Public Sub OpenTempProject()
    OpenDocument PathToTempVBP
End Sub
