Public Class LoginForm1

    ' TODO: Code zum Durchführen der benutzerdefinierten Authentifizierung mithilfe des angegebenen Benutzernamens und des Kennworts hinzufügen 
    ' (Siehe http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' Der benutzerdefinierte Prinzipal kann anschließend wie folgt an den Prinzipal des aktuellen Threads angefügt werden: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' wobei CustomPrincipal die IPrincipal-Implementierung ist, die für die Durchführung der Authentifizierung verwendet wird. 
    ' Anschließend gibt My.User Identitätsinformationen zurück, die in das CustomPrincipal-Objekt gekapselt sind,
    ' z.B. den Benutzernamen, den Anzeigenamen usw.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        
        Shell("net use " & Main.ComboBox1.Text & ": \\192.168.2.17\c /user:" & UsernameTextBox.Text & " " & PasswordTextBox.Text, AppWinStyle.NormalFocus)
        Main.ComboBox1.Items.Clear()
        MsgBox("ReadyNAS Root Drive connected")
        Call getfreedrives()

        Me.Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub LoginForm1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        UsernameTextBox.Text = "admin"
        PasswordTextBox.Select()
    End Sub
End Class
