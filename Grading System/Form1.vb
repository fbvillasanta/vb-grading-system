Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (IsNumeric(mq1s.Text) And IsNumeric(mq2s.Text) And IsNumeric(msw1s.Text) And IsNumeric(msw2s.Text) And IsNumeric(mhw1s.Text) And IsNumeric(mhw2s.Text) And IsNumeric(mprojs.Text) And IsNumeric(mq1n.Text) And IsNumeric(mq2n.Text) And IsNumeric(msw1n.Text) And IsNumeric(msw2n.Text) And IsNumeric(mhw1n.Text) And IsNumeric(mhw2n.Text) And IsNumeric(mprojn.Text)) Then
            mq1r.Text = (mq1s.Text / mq1n.Text * 50) + 50
            mq2r.Text = (mq2s.Text / mq2n.Text * 50) + 50
            msw1r.Text = (msw1s.Text / msw1n.Text * 50) + 50
            msw2r.Text = (msw2s.Text / msw2n.Text * 50) + 50
            mhw1r.Text = (mhw1s.Text / mhw1n.Text * 50) + 50
            mhw2r.Text = (mhw2s.Text / mhw2n.Text * 50) + 50
            mprojr.Text = (mprojs.Text / mprojn.Text * 50) + 50
            mer.Text = (mes.Text / men.Text * 50) + 50
            mg.Text = ((mq1r.Text + +mq2r.Text + +msw1r.Text + +msw2r.Text + +mhw1r.Text + +mhw2r.Text + +mprojr.Text) / 700 * 70) + (mer.Text / 100 * 30)
        Else
            MsgBox("Invalid Input")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If (IsNumeric(fq1s.Text) And IsNumeric(fq2s.Text) And IsNumeric(fsw1s.Text) And IsNumeric(fsw2s.Text) And IsNumeric(fhw1s.Text) And IsNumeric(fhw2s.Text) And IsNumeric(fprojs.Text) And IsNumeric(fq1n.Text) And IsNumeric(fq2n.Text) And IsNumeric(fsw1n.Text) And IsNumeric(fsw2n.Text) And IsNumeric(fhw1n.Text) And IsNumeric(fhw2n.Text) And IsNumeric(fprojn.Text)) Then
            fq1r.Text = (fq1s.Text / fq1n.Text * 50) + 50
            fq2r.Text = (fq2s.Text / fq2n.Text * 50) + 50
            fsw1r.Text = (fsw1s.Text / fsw1n.Text * 50) + 50
            fsw2r.Text = (fsw2s.Text / fsw2n.Text * 50) + 50
            fhw1r.Text = (fhw1s.Text / fhw1n.Text * 50) + 50
            fhw2r.Text = (fhw2s.Text / fhw2n.Text * 50) + 50
            fprojr.Text = (fprojs.Text / fprojn.Text * 50) + 50
            fer.Text = (fes.Text / fen.Text * 50) + 50
            fg.Text = ((fq1r.Text + +fq2r.Text + +fsw1r.Text + +fsw2r.Text + +fhw1r.Text + +fhw2r.Text + +fprojr.Text) / 700 * 70) + (fer.Text / 100 * 30)
        Else
            MsgBox("Invalid Input")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If (IsNumeric(mg.Text) And IsNumeric(fg.Text)) Then
            sg.Text = (mg.Text + +fg.Text) / 2
        Else
            MsgBox("Invalid Input")
        End If
    End Sub

    Private Sub fer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fer.TextChanged

    End Sub

    Private Sub mes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mes.TextChanged

    End Sub

    Private Sub fes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fes.TextChanged

    End Sub

    Private Sub men_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles men.TextChanged

    End Sub

    Private Sub fen_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fen.TextChanged

    End Sub

    Private Sub mer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mer.TextChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        n.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

        mq1s.Text = ""
        mq2s.Text = ""
        msw1s.Text = ""
        msw2s.Text = ""
        mhw1s.Text = ""
        mhw2s.Text = ""
        mprojs.Text = ""
        mq1n.Text = ""
        mq2n.Text = ""
        msw1n.Text = ""
        msw2n.Text = ""
        mhw1n.Text = ""
        mhw2n.Text = ""
        mprojn.Text = ""
        mq1r.Text = ""
        mq2r.Text = ""
        msw1r.Text = ""
        msw2r.Text = ""
        mhw1r.Text = ""
        mhw2r.Text = ""
        mprojr.Text = ""
        mes.Text = ""
        men.Text = ""
        mer.Text = ""

        fq1s.Text = ""
        fq2s.Text = ""
        fsw1s.Text = ""
        fsw2s.Text = ""
        fhw1s.Text = ""
        fhw2s.Text = ""
        fprojs.Text = ""
        fq1n.Text = ""
        fq2n.Text = ""
        fsw1n.Text = ""
        fsw2n.Text = ""
        fhw1n.Text = ""
        fhw2n.Text = ""
        fprojn.Text = ""
        fq1r.Text = ""
        fq2r.Text = ""
        fsw1r.Text = ""
        fsw2r.Text = ""
        fhw1r.Text = ""
        fhw2r.Text = ""
        fprojr.Text = ""
        fes.Text = ""
        fen.Text = ""
        fer.Text = ""

        mg.Text = ""
        fg.Text = ""
        sg.Text = ""
    End Sub
End Class