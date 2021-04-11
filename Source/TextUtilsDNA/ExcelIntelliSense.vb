Option Explicit On
Option Strict On

Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense

Namespace TextUtilsDna
    Public Class ExcelIntelliSense
        Implements IExcelAddIn

        Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
            IntelliSenseServer.Install()
        End Sub

        Public Sub AutoClose() Implements IExcelAddIn.AutoClose
            IntelliSenseServer.Uninstall()
        End Sub

    End Class
End Namespace

