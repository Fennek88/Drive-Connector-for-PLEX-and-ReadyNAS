Module Module1
    Public Sub renamenetworkdrive(driveletter As String, newname As String)
        Dim oShell
        Dim done As Boolean = False
        oShell = CreateObject("Shell.Application")
        Do Until done = True
            If fctVerzeichnisExists(driveletter & ":\") Then
                Try
                    oShell.NameSpace(driveletter).Self.Name = newname
                    done = True
                Catch
                End Try
            End If
        Loop
    End Sub

    Public Sub getfreedrives()
        Dim drives
        For i = 1 To 26
            If FileIO.FileSystem.DirectoryExists(Chr(i + 64) & ":\") Then
                drives = System.IO.DriveInfo.GetDrives
            Else
                Main.ComboBox1.Items.Add(Chr(i + 64))
            End If
        Next
    End Sub

    Public Function fctVerzeichnisExists(ByVal strPfad As String) As Boolean
        On Error GoTo fehler
        ChDir(strPfad)
        fctVerzeichnisExists = True
        Exit Function
fehler:
        fctVerzeichnisExists = False
    End Function

End Module
