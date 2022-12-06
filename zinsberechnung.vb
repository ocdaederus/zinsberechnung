Public Class zinsberechnung

    Dim kapital, zinssatz, gesamtkapital As Single
    Dim zeitraum As Long

    Private Sub cmd_berechnen_Click(sender As Object, e As System.EventArgs) Handles cmd_berechnen.Click

        'Eingabe
        kapital = txt_kapital.Text
        zinssatz = txt_zinssatz.Text
        zeitraum = txt_zeitraum.Text

        'Verarbeitung
        For i = Now.Year To Now.Year + zeitraum - 1 Step 1

            gesamtkapital = ((kapital * zinssatz) / 100) + kapital

            kapital = gesamtkapital

            'Ausgabe
            'Kommentar #1: Die Formatierung der Variable "gesamtkapital" wird zu einem späteren Zeitpunkt hinzugefügt.
            lst_ergebnisse.Items.Add(i & "        " & zinssatz & "        " & gesamtkapital)

        Next i

    End Sub

    Private Sub cmd_löschen_Click(sender As Object, e As System.EventArgs) Handles cmd_löschen.Click

        txt_zinssatz.Clear()
        txt_zeitraum.Clear()
        txt_kapital.Clear()
        lst_ergebnisse.Items.Clear()

    End Sub

    Private Sub cmd_beenden_Click(sender As Object, e As System.EventArgs) Handles cmd_beenden.Click

        'Application.Exit()
        Me.Close()
        End

    End Sub
End Class