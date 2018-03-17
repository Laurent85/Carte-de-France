Imports System.IO

Class MainWindow

    Public temp As Integer
    Public score As Integer
    Public coups1 As Integer
    Public p As Integer = 0
    Public triche As Integer = 0

    Private Sub Fenêtre_Principale(sender As Object, e As RoutedEventArgs) Handles Principal.Loaded

        'Activation du "déposer" et tous les "?" en bleu
        For i As Integer = 0 To 95
            Dim cible As Label = Me.Cibles.Children.Item(i)
            If cible IsNot Nothing Then
                cible.AllowDrop = True
                cible.Foreground = Brushes.Blue
            End If
        Next

        'Dimensions et marges des noms de départements
        For i As Integer = 0 To 95
            Dim etiquette As Label = Me.Noms.Children.Item(i)
            etiquette.HorizontalAlignment = Windows.HorizontalAlignment.Left
            etiquette.VerticalAlignment = Windows.VerticalAlignment.Top
        Next

        Choix_dep.IsChecked = True
        Bilan_des_placements()
        p = 0
        En_cours.Foreground = Brushes.Gray
        En_cours.Content = "Départements non vérifiés : " & p.ToString

        For i As Int32 = 0 To 95
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            AddHandler source.MouseEnter, AddressOf Souris
            AddHandler source.MouseLeave, AddressOf plus_souris
        Next

        temp = 0
        score = 0
        coups1 = 0
        Coups.Content = "Coups : " & coups1
        Points.Content = score

    End Sub

    Private Sub Window1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Dim result As MessageBoxResult = MessageBox.Show("Voulez-vous quitter le jeu ?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Question)

        If result = MessageBoxResult.Yes Then
            fichier_Scores()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub label1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Carte01.MouseDown
        'Activation du "glisser"
        Dim lbl As Label = CType(sender, Label)
        DragDrop.DoDragDrop(lbl, lbl.Content, DragDropEffects.Copy)
    End Sub

    Private Sub label2_DragEnter(ByVal sender As Object, ByVal e As System.Windows.DragEventArgs) Handles Carte01.DragEnter
        'Activation de la copie de l'élément glissé
        If e.Data.GetDataPresent(DataFormats.Text) Then
            e.Effects = DragDropEffects.Copy
        Else
            e.Effects = DragDropEffects.None
        End If

    End Sub

    Private Sub label2_Drop(ByVal sender As Object, ByVal e As System.Windows.DragEventArgs) Handles Carte01.Drop

        For j = 0 To 95
            Dim cible As Label = CType(Cibles.Children.Item(j), Label)
            Dim nom As Label = CType(Noms.Children.Item(j), Label)
            'Cherche si existe déjà
            'Si cible = source et source différent de "?" alors..
            If cible.Content = e.Data.GetData(DataFormats.Text) And e.Data.GetData(DataFormats.Text) <> "?" Then
                cible.Content = "?"
                cible.ToolTip = vbNullString
                cible.Foreground = Brushes.Blue
                CType(sender, Label).Content = e.Data.GetData(DataFormats.Text)
                CType(sender, Label).Foreground = Brushes.White
            Else
                CType(sender, Label).Content = e.Data.GetData(DataFormats.Text)
                CType(sender, Label).Foreground = Brushes.White
            End If
        Next j

        'Numéros gris et infobulle pour les départements placés
        For i As Int32 = 0 To 95
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            For j As Int32 = 0 To 95
                Dim cible As Label = CType(Cibles.Children.Item(j), Label)
                If cible.Content = source.Content Then
                    source.Foreground = Brushes.LightGray
                    cible.ToolTip = CType(Noms.Children.Item(i), Label).Content & " (" & CType(Villes.Children.Item(i), Label).Content & ")"
                    Exit For
                Else
                    source.Foreground = Brushes.Red
                End If
                If cible.Content = "?" Then
                    cible.ToolTip = vbNullString
                    cible.Foreground = Brushes.Blue
                End If
            Next j
        Next i

        p = 0
        For i As Int32 = 0 To 95
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If cible.Content <> "?" And (Equals(cible.Foreground, Brushes.White)) Then
                p = p + 1
                En_cours.Foreground = Brushes.Gray
                En_cours.Content = "Départements non vérifiés : " & p.ToString
            End If
        Next


        Test_réponse()

        'Vérifier_Click(sender, e)
        My.Computer.Audio.Play(My.Resources.Bruit1, AudioPlayMode.Background)

    End Sub

    Private Sub clear(ByVal sender As Object, ByVal e As MouseButtonEventArgs) Handles Carte01.MouseRightButtonDown, Carte03.MouseRightButtonDown, Carte10.MouseRightButtonDown, Carte02.MouseRightButtonDown, Carte08.MouseRightButtonDown, Carte05.MouseRightButtonDown, Carte07.MouseRightButtonDown, Carte04.MouseRightButtonDown, Carte06.MouseRightButtonDown, Carte11.MouseRightButtonDown, Carte09.MouseRightButtonDown

        'effacement par clic-droit
        CType(sender, Label).AllowDrop = True

        For i As Int32 = 0 To Sources.Children.Count - 1
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            If CType(sender, Label).Content = source.Content Then
                source.Foreground = Brushes.Red
                AddHandler source.MouseDown, AddressOf label1_MouseDown 'activation du glissé
            End If
        Next i

        CType(sender, Label).Content = "?"
        CType(sender, Label).Foreground = Brushes.Blue
        CType(sender, Label).ToolTip = vbNullString

    End Sub

    Private Sub Vérifier_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles Vérifier.Click

        Dim b As Integer = 0
        Dim m As Integer = 0
        Dim n As Integer = 0
        Dim j As String
        Dim total As Integer

        Bilan_des_placements()

        For i As Int32 = 0 To 20
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If i < 9 Then
                j = "0" & i + 1.ToString
            Else
                j = i + 1.ToString
            End If
            If i = 19 Then
                j = "2A"
            End If
            If i = 20 Then
                j = "2B"
            End If
            If cible.Content = j.ToString Then
                cible.Foreground = Brushes.Green
                b = b + 1
                Bien.Foreground = Brushes.Green
                Bien.Content = "Départements bien placés : " & b.ToString
            End If
            If cible.Content = "?" Then
                cible.Foreground = Brushes.Blue
                n = n + 1
                Absent.Foreground = Brushes.Blue
                Absent.Content = "Départements non placés : " & n.ToString
            End If
            If cible.Content <> "?" And cible.Content <> j.ToString Then
                cible.Foreground = Brushes.Red
                m = m + 1
                Mal.Foreground = Brushes.Red
                Mal.Content = "Départements mal placés : " & m.ToString
            End If
        Next i

        For i As Int32 = 21 To 95
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            j = (i).ToString

            If cible.Content = (j).ToString Then
                cible.Foreground = Brushes.Green
                b = b + 1
                Bien.Foreground = Brushes.Green
                Bien.Content = "Départements bien placés : " & b.ToString
            End If
            If cible.Content = "?" Then
                cible.Foreground = Brushes.Blue
                n = n + 1
                Absent.Foreground = Brushes.Blue
                Absent.Content = "Départements non placés : " & n.ToString
            End If
            If cible.Content <> "?" And cible.Content <> (j).ToString Then
                cible.Foreground = Brushes.Red
                m = m + 1
                Mal.Foreground = Brushes.Red
                Mal.Content = "Départements mal placés : " & m.ToString
            End If
        Next i
        If temp <> b Then
            total = (b - temp) * 2
        Else
            total = 0
        End If
        score = score + Val(total) - m
        temp = b
        coups1 = coups1 + 1
        Coups.Content = "Coups : " & coups1

        Points.Foreground = Brushes.SaddleBrown
        Points.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
        Points.VerticalContentAlignment = Windows.VerticalAlignment.Top
        Points.Content = score
        My.Computer.Audio.Play(My.Resources.Bruit, AudioPlayMode.Background)

        Test_réponse()

        p = 0
        En_cours.Content = "Départements non vérifiés : " & p.ToString
        Cacher_Checked(sender, e)
        Voir_Checked(sender, e)

        If b = 96 And triche = 0 Then
            Dim result As MessageBoxResult = MessageBox.Show("Bravo " & copie.Content & " !" & Chr(10) & "Tu as placé les 96 départements en " & coups1 & " coups." & Chr(10) & "Tu as totalisé " & score & " points." & Chr(10) & "Veux-tu refaire une partie ?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.Yes Then
                fichier_Scores()
                Effacer_Click(sender, e)
            Else
                fichier_Scores()
                Me.Close()
            End If
        End If

    End Sub

    Private Sub Solutions_Click(sender As Object, e As RoutedEventArgs) Handles Solutions.Click

        Dim b As Integer = 0
        Dim m As Integer = 0
        Dim n As Integer = 0
        Dim j As String
        triche = 1

        Bilan_des_placements()

        For i As Int32 = 0 To 20
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If i < 9 Then
                j = "0" & i + 1.ToString
                cible.Content = j.ToString
                cible.ToolTip = CType(Noms.Children.Item(i), Label).Content & " (" & CType(Villes.Children.Item(i), Label).Content & ")"
            Else
                j = i + 1.ToString
                cible.Content = j.ToString
                cible.ToolTip = CType(Noms.Children.Item(i), Label).Content & " (" & CType(Villes.Children.Item(i), Label).Content & ")"
            End If
            If i = 19 Then
                j = "2A"
                cible.Content = j.ToString
                cible.ToolTip = CType(Noms.Children.Item(i), Label).Content & " (" & CType(Villes.Children.Item(i), Label).Content & ")"
            End If
            If i = 20 Then
                j = "2B"
                cible.Content = j.ToString
                cible.ToolTip = CType(Noms.Children.Item(i), Label).Content & " (" & CType(Villes.Children.Item(i), Label).Content & ")"
            End If
            If cible.Content = j.ToString Then
                cible.Foreground = Brushes.Green
                b = b + 1
                Bien.Foreground = Brushes.Green
                Bien.Content = "Départements bien placés : " & b.ToString
            End If
            If cible.Content = "?" Then
                cible.Foreground = Brushes.Blue
                n = n + 1
                Absent.Foreground = Brushes.Blue
                Absent.Content = "Départements non placés : " & n.ToString
            End If
            If cible.Content <> "?" And cible.Content <> j.ToString Then
                cible.Foreground = Brushes.Red
                m = m + 1
                Mal.Foreground = Brushes.Red
                Mal.Content = "Départements mal placés : " & m.ToString
            End If
        Next i

        For i As Int32 = 21 To 95
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            j = (i).ToString
            cible.Content = (j).ToString
            cible.ToolTip = CType(Noms.Children.Item(i), Label).Content & " (" & CType(Villes.Children.Item(i), Label).Content & ")"
            If cible.Content = (j).ToString Then
                cible.Foreground = Brushes.Green
                b = b + 1
                Bien.Foreground = Brushes.Green
                Bien.Content = "Départements bien placés : " & b.ToString
            End If
            If cible.Content = "?" Then
                cible.Foreground = Brushes.Blue
                n = n + 1
                Absent.Foreground = Brushes.Blue
                Absent.Content = "Départements non placés : " & n.ToString
            End If
            If cible.Content <> "?" And cible.Content <> (j).ToString Then
                cible.Foreground = Brushes.Red
                m = m + 1
                Mal.Foreground = Brushes.Red
                Mal.Content = "Départements mal placés : " & m.ToString
            End If
        Next i

        For i As Integer = 0 To 95
            Dim cible As Label = Me.Cibles.Children.Item(i)
            If cible IsNot Nothing Then
                cible.AllowDrop = False
            End If
        Next

        Test_réponse()

        fichier_Scores()
        Cacher_Checked(sender, e)
        Voir_Checked(sender, e)

    End Sub

    Private Sub Effacer_Click(sender As Object, e As RoutedEventArgs) Handles Effacer.Click

        For i As Int32 = 0 To 95
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            Dim nom As Label = CType(Noms.Children.Item(i), Label)
            Dim ville As Label = CType(Villes.Children.Item(i), Label)
            cible.Content = "?"
            cible.AllowDrop = True
            cible.Foreground = Brushes.Blue
            cible.ToolTip = vbNullString
            source.Foreground = Brushes.Red
            AddHandler source.MouseDown, AddressOf label1_MouseDown
            AddHandler cible.MouseDown, AddressOf label1_MouseDown
            AddHandler cible.MouseRightButtonDown, AddressOf clear
            Bilan_des_placements()
            source.Visibility = Windows.Visibility.Visible
            nom.Visibility = Windows.Visibility.Visible
            ville.Visibility = Windows.Visibility.Hidden
        Next i

        fichier_Scores()

        temp = 0
        score = 0
        coups1 = 0
        Coups.Content = "Coups : " & coups1
        Points.Content = score

    End Sub

    Private Sub Choix_dep_Checked(sender As Object, e As RoutedEventArgs) Handles Choix_dep.Checked

        For i As Integer = 0 To 95
            Dim nom As Label = Me.Noms.Children.Item(i)
            nom.Visibility = Windows.Visibility.Visible
            nom.ToolTip = CType(Villes.Children.Item(i), Label).Content
        Next

        For i As Integer = 0 To 95
            Dim nom As Label = Me.Villes.Children.Item(i)
            nom.Visibility = Windows.Visibility.Hidden
        Next

        Cacher_Checked(sender, e)
        Voir_Checked(sender, e)

    End Sub

    Private Sub Choix_villes_Checked(sender As Object, e As RoutedEventArgs) Handles Choix_villes.Checked

        For i As Integer = 0 To 95
            Dim nom As Label = Me.Noms.Children.Item(i)
            nom.Visibility = Windows.Visibility.Hidden
        Next

        For i As Integer = 0 To 95
            Dim nom As Label = Me.Villes.Children.Item(i)
            nom.Visibility = Windows.Visibility.Visible
            nom.ToolTip = CType(Noms.Children.Item(i), Label).Content
        Next

        Cacher_Checked(sender, e)
        Voir_Checked(sender, e)

    End Sub

    Sub fichier_Scores()
        If triche = 0 Then
            Try
                Dim FichierScores As StreamWriter = New StreamWriter("\\laurent\d\Départements.txt", True)
                FichierScores.WriteLine(Now & Chr(9) & copie.Content & Chr(9) & "Points : " & Points.Content & Chr(9) & Coups.Content & Chr(9) & Bien.Content & Chr(9) & Mal.Content & Chr(9) & Absent.Content)
                FichierScores.Close()

            Catch ex As Exception

            End Try

            Try
                Dim FichierScores1 As StreamWriter = New StreamWriter("c:\Départements.txt", True)
                FichierScores1.WriteLine(Now & Chr(9) & copie.Content & Chr(9) & "Points : " & Points.Content & Chr(9) & Coups.Content & Chr(9) & Bien.Content & Chr(9) & Mal.Content & Chr(9) & Absent.Content)
                FichierScores1.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Sub Test_réponse()
        'Si bonne réponse, on met la source en vert et on bloque le déplacement du département
        For i As Int32 = 0 To 95
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If (Equals(cible.Foreground, Brushes.Green)) Then
                source.Foreground = Brushes.Green
                RemoveHandler source.MouseDown, AddressOf label1_MouseDown
                RemoveHandler cible.MouseDown, AddressOf label1_MouseDown
                RemoveHandler cible.MouseRightButtonDown, AddressOf clear
                cible.AllowDrop = False
                cible.ToolTip = CType(Noms.Children.Item(i), Label).Content & " (" & CType(Villes.Children.Item(i), Label).Content & ")"
            Else
                'cible.ToolTip = vbNullString
            End If
        Next i
    End Sub

    Sub Bilan_des_placements()
        Bien.Foreground = Brushes.Green
        Mal.Foreground = Brushes.Red
        Absent.Foreground = Brushes.Blue
        Bien.Content = "Départements bien placés : 0"
        Mal.Content = "Départements mal placés : 0"
        Absent.Content = "Départements non placés : 0"
    End Sub

    Private Sub Cacher_Checked(sender As Object, e As RoutedEventArgs) Handles Cacher.Checked

        If Cacher.IsChecked = True Then
            For i As Int32 = 0 To 95
                Dim cible As Label = CType(Cibles.Children.Item(i), Label)
                Dim source As Label = CType(Sources.Children.Item(i), Label)
                Dim nom As Label = CType(Noms.Children.Item(i), Label)
                Dim ville As Label = CType(Villes.Children.Item(i), Label)
                If (Equals(cible.Foreground, Brushes.Green)) And Choix_dep.IsChecked Then
                    source.Visibility = Windows.Visibility.Hidden
                    nom.Visibility = Windows.Visibility.Hidden
                End If
                If (Equals(cible.Foreground, Brushes.Green)) And Choix_villes.IsChecked Then
                    source.Visibility = Windows.Visibility.Hidden
                    ville.Visibility = Windows.Visibility.Hidden
                End If
            Next i
        End If

    End Sub

    Private Sub Voir_Checked(sender As Object, e As RoutedEventArgs) Handles Cacher.Unchecked

        If Cacher.IsChecked = False Then
            For i As Int32 = 0 To 95
                Dim cible As Label = CType(Cibles.Children.Item(i), Label)
                Dim source As Label = CType(Sources.Children.Item(i), Label)
                Dim nom As Label = CType(Noms.Children.Item(i), Label)
                Dim ville As Label = CType(Villes.Children.Item(i), Label)
                If (Equals(cible.Foreground, Brushes.Green)) And Choix_dep.IsChecked Then
                    source.Visibility = Windows.Visibility.Visible
                    nom.Visibility = Windows.Visibility.Visible
                End If
                If (Equals(cible.Foreground, Brushes.Green)) And Choix_villes.IsChecked Then
                    source.Visibility = Windows.Visibility.Visible
                    ville.Visibility = Windows.Visibility.Visible
                End If
            Next i
        End If

    End Sub

    Private Sub Souris(sender As Object, e As MouseEventArgs) Handles lbl61.MouseEnter

        For i As Int32 = 0 To 95
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            Dim nom As Label = CType(Noms.Children.Item(i), Label)
            If source.IsMouseOver = True And source.Foreground.ToString = "#FFF90B0B" Then
                nom.Foreground = Brushes.Red
                nom.FontSize = 14
            End If
            If source.IsMouseOver = True And (Equals(source.Foreground, Brushes.Red)) Then
                nom.Foreground = Brushes.Red
                nom.FontSize = 14
            End If
        Next

    End Sub

    Private Sub plus_souris(sender As Object, e As MouseEventArgs) Handles lbl61.MouseLeave

        For i As Int32 = 0 To 95
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            Dim nom As Label = CType(Noms.Children.Item(i), Label)
            If source.IsMouseOver = False Then
                nom.Foreground = Brushes.Gray
                nom.FontSize = 12
            End If
        Next

    End Sub
End Class
