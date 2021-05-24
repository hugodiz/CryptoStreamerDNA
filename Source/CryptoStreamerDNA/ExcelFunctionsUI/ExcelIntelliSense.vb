Option Explicit On
Option Strict On

Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense
Imports System.Net.Http


Public Class ExcelIntelliSense
    Implements IExcelAddIn

    ' Initialization of the ExcelDna IntelliSense Server
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        SessionBinanceAPIClient = New HttpClient()
        IntelliSenseServer.Install()
    End Sub

    ' Disposing of resources when closing Excel
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        SessionBinanceAPIClient.Dispose()
        IntelliSenseServer.Uninstall()
    End Sub

End Class


