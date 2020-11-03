Imports System.Data

Public Class UserControl1

    Private Sub Button_Click(sender As Object, e As System.Windows.RoutedEventArgs)

        'Build a data table
        Dim dt As New DataTable
        dt.Columns.Add("ID")
        dt.Columns.Add("NAME")
        dt.Columns.Add("DATE")

        dt.Rows.Add({1, "Jonh", New Date(1981, 10, 12)})
        dt.Rows.Add({1, "Ford", New Date(1971, 1, 22)})

    End Sub

End Class
