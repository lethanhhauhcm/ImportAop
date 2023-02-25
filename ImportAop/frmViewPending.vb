Public Class frmViewPending
    Private Sub dgrPendings_SelectionChanged(sender As Object, e As EventArgs) Handles dgrPendings.SelectionChanged
        If dgrPendings.CurrentRow Is Nothing Then
            txtQuerry.Text = ""
        Else
            txtQuerry.Text = dgrPendings.CurrentRow.Cells("Querry").Value
        End If
    End Sub

    Private Sub lbkSearch_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSearch.LinkClicked

    End Sub
    Private Function Search() As Boolean
        Dim strQuerry As String = "Select * from AopQueue where Status='OK' order by RecId"

        pobjSqlRas.LoadDataGridView(dgrPendings, strQuerry)
        dgrPendings.Columns("Querry").Visible = False

        Return True
    End Function
End Class