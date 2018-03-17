Public Class Prénom

    Private Sub Prénom3(sender As Object, e As RoutedEventArgs) Handles Prénom2.Loaded
        prenom.Focus()
        prenom.TextAlignment = TextAlignment.Center
    End Sub
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        If prenom.Text <> vbNullString Then
            Dim mainwindow As New MainWindow
            mainwindow.Bienvenue_Copy.Content = "Bonne chance " & prenom.Text & " !"
            mainwindow.copie.Content = prenom.Text
            mainwindow.Show()
            Me.Close()
        Else
            prenom.Focus()
            prenom.TextAlignment = TextAlignment.Center
        End If

    End Sub

    Private Sub prenom_TextChanged(sender As Object, e As TextChangedEventArgs) Handles prenom.TextChanged
        prenom.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
        prenom.VerticalContentAlignment = Windows.VerticalAlignment.Center
    End Sub
End Class
