' Creates a complete path, e.g. "mkdir -p /a/path/to/heaven" unix like command.
' It uses recursion for moving "directories" from left to right as follows:
'
' Function should be called: CreatePath("C:\this\path\does\not\necessarily\exist\completely")
' and goes from rigth to left checking which folder exist in order to create the path. 
' Left path and right path in 2nd run in the example above will be:
'
' LEFT PATH:                                    RIGHT PATH:
' "C:\this\path\does\not\necessarily\exist\"    "completely"
'
' and so on...
' LEFT PATH:                                    RIGHT PATH:
' "C:\this\path\does\not\necessarily\"    "exist\completely"
' "C:\this\path\does\not\"    "necessarily\exist\completely"
' "C:\this\path\does\"    "not\necessarily\exist\completely"
'
' Supposing "C:\this\path\does" exists, the function will start creating and building the
' path and it will go reversed up to the complete path:
'
' LEFT PATH (already exists):                    RIGHT PATH:
' "C:\this\path\does\not\"    "necessarily\exist\completely"
' "C:\this\path\does\not\necessarily\"    "exist\completely"
' "C:\this\path\does\not\necessarily\exist\"    "completely"
'
' It lacks for some error checking while calling CreateFolder and it does not validate
' strings like "C:\this\is\a\malicious\\\path\being\\built\" and it does not check for
' invalid characters for folders like ':' and so on.
Private Function CreatePath(ByVal lPath As String, Optional ByVal rPath As String = "") As Boolean
    Dim fso As Scripting.FileSystemObject
    fso = New Scripting.FileSystemObject
    If fso.FolderExists(lPath & rPath) Then
        CreatePath = True
    Else
        If fso.FolderExists(lPath) Then
            Dim splitPath() As String = Split(rPath, "\", 2)
            Dim folder As String = fso.BuildPath(lPath, splitPath(0))
            fso.CreateFolder(folder)
            If splitPath.Length = 1 Then
                rPath = ""
            ElseIf splitPath.Length = 2 Then
                rPath = splitPath(1)
            End If
            CreatePath = CreatePath(folder, rPath)
        Else
            lPath = StrReverse(lPath)
            If Strings.Left(lPath, 1) = "\" Then
                lPath = Strings.Right(lPath, Len(lPath) - 1)
            End If
            Dim index As Integer = InStr(lPath, "\")
            lPath = StrReverse(lPath)
            CreatePath = CreatePath(Strings.Mid(lPath, 1, Len(lPath) - index + 1), fso.BuildPath(Strings.Right(lPath, index - 1), rPath))
        End If
    End If
End Function

' This is a straightforward solution. It builds the solution up, from the first directory until the last one.
' It creates an array of "directories" and starts concatenating each directory building it if necessary.
' It does not instantiates a FileSystemObject each time. Only 1 FileSystemObject is created instead.
' The only difference remains that all the directories are treated no matter if they exist or not.
' The above solution (it leaks) stops when a directory path exists, and starts building the solution.
' Maybe improving some strings functions and freeing FileSystemObject memory the solutions is _better_ in terms
' of complexity, but I stick with this one in real life.
Sub CreatePath2(ByVal path As String)
        Dim fso As Scripting.FileSystemObject
        Dim splitPath() As String = Split(path, "\")
        Dim folder As String = splitPath(0)
        Dim i As Integer
        fso = New Scripting.FileSystemObject
        For i = 1 To splitPath.Length - 1
            folder = folder & "\" & splitPath(i)
            If Not fso.FolderExists(folder) Then
                fso.CreateFolder(folder)
            End If
        Next i
        fso = Nothing
End Sub