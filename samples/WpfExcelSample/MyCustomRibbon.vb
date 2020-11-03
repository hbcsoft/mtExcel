Imports System
Imports ExcelDna.Integration.CustomUI
Imports mtExcel.Core.TaskPane

Public Class MyCustomRibbon : Inherits ExcelRibbon

    Public Overrides Sub OnStartupComplete(ByRef custom As Array)


        Dim uc As New UserControl1
        uc.ShowInTaskPane("real cool")


        MyBase.OnStartupComplete(custom)
    End Sub

End Class
