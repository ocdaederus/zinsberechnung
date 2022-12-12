Public Class zinsberechnung

    Dim kapital, zinssatz, gesamtkapital As Single
    Dim zeitraum As Long

    Private Sub cmd_berechnen_Click(sender As Object, e As System.EventArgs) Handles cmd_berechnen.Click

        'Eingabe
        kapital = txt_kapital.Text
        zinssatz = txt_zinssatz.Text
        zeitraum = txt_zeitraum.Text

        'Verarbeitung
        lst_ergebnisse.Items.Clear()
        
        For i = Now.Year To Now.Year + zeitraum - 1 Step 1

            gesamtkapital = ((kapital * zinssatz) / 100) + kapital.ToString("#,##0.00")
        
            gesamtkapital = Math.Round(gesamtkapital, 2)
        
            zinsen = kapital * zinssatz / 100
        
            zinsen = Math.Round(zinsen, 2)

            kapital = gesamtkapital

            'Ausgabe
            lst_ergebnisse.Items.Add(i & "                  " & Format("#,##0.00", [zinsen]) & "€" & "                  " & gesamtkapital & "€")

        Next i

    End Sub

        Private Sub cmd_kontrolle_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmd_kontrolle.Click

        'Eingabe

        kapital = txt_kapital.Text
        zinssatz = txt_zinssatz.Text
        zeitraum = txt_zeitraum.Text

        'Verarbeitung

        kumultzinsen = kapital * ((1 + zinssatz / 100) ^ zeitraum) - kapital

        kumultzinsen = Math.Round(kumultzinsen, 2)

        'Ausgabe

        lbl_kumultzinsen.Text = kumultzinsen.ToString("#,##0.00€")

        lbl_gesamtkapital.Text = gesamtkapital.ToString("#,##0.00€")

        If kapital <= 0 And zeitraum <= 0 And (zinssatz = 0 Or zinssatz < 0) Then

            MsgBox("Bitte geben Sie Ihr Kapital, den Zinssatz sowie den Zeitraum an und lassen Sie das am Ende zu erhaltene Geld berechnen, damit die Kontrolle erfolgen kann")

        End If

    End Sub

    Private Sub cmd_löschen_Click(sender As Object, e As System.EventArgs) Handles cmd_löschen.Click

        txt_zinssatz.Clear()
        txt_zeitraum.Clear()
        txt_kapital.Clear()
        lst_ergebnisse.Items.Clear()
        lbl_kumultzinsen.Text = ""
        lbl_gesamtkapital.Text = ""

    End Sub

    Private Sub cmd_beenden_Click(sender As Object, e As System.EventArgs) Handles cmd_beenden.Click

        Application.Exit()
        'Me.Close()
        End

    End Sub

    Private Sub cmd_zusatz_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmd_zusatz.Click

        zusatz.Show()

    End Sub

End Class
