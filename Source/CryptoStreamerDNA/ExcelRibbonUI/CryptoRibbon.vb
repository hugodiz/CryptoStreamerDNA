Option Explicit On
Option Strict On

Imports System.IO
Imports System.Resources
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports Sys = System.Drawing

' *****************************************************************************************************************************
' UI Ribbon Class - CryptoRibbon
' -----------------------------------------------------------------------------------------------------------------------------
' This class inherits / relies on the ExcelDNA ExcelRibbon class
'     Basically, it loads into Excel the Ribbon we have defined in CryptoRibbon.xml
'     Then, it binds our relevant reactive functionalities (the Doge button images and the telemetry feeds)
'         using the getXXX / OnGetXXX methods assigned in the Ribbon xml and implemented here
'     Lastly, here the InvalidateControl(ControlId) calls will essentially trigger a refresh / recalculation
'         of all quantities bound to each control (which will in turn call the OnGet methods, recalculating quantities)
' 
' Hence, the InvalidateControl() calls are made when handling ServerStateChanged and TelemetryChanged events
'     (raised by the CryptoRTDServer)
'
' Conversely, when the Doge button is clicked, a notification event is raised here (DogeClicked)
'     which is handled by the CryptoRTDServer, which toggles RTDStatus.IsSwitchedOn accordingly
'         which of course then raises ServerStateChanged, causing the Doge image and the Doge button label to change
<ComVisible(True)>
Public Class CryptoRibbon
    Inherits ExcelRibbon

    Private ThisApp As Excel.Application
    Private ThisRibbon As IRibbonUI
    ' Telemetry values holder (will be refreshed whenever we handle a TelemetryChanged event here)
    '     The latest telemetry values are passed as event arguments by the TelemetryChanged event (raised by CryptoRTDServer)
    Private LatestTelemetry As _
        (WeightInterval As Double, WeightLimit As Double, UsedWeight As Integer, TimeBetweenRequests As String) = (0, 0, 0, "")

    ' All 4 images are stored as embedded resources of the project - GetCustomImage is their Getter
    Private DogeOpenEyes As Sys.Bitmap = GetCustomImage("DogeOpenEyes.bmp")
    Private DogeClosedEyes As Sys.Bitmap = GetCustomImage("DogeClosedEyes.bmp")
    Private DogeOpenEyesHourglass As Sys.Bitmap = GetCustomImage("DogeOpenEyesHourglass.bmp")
    Private DogeClosedEyesHourglass As Sys.Bitmap = GetCustomImage("DogeClosedEyesHourglass.bmp")

    Public Shared Event DogeClicked()

    ' The Ribbon xml is stored as an embedded resource of the project - GetCustomRibbonXml is its Getter
    Public Overrides Function GetCustomUI(RibbonId As String) As String
        ThisApp = DirectCast(ExcelDnaUtil.Application, Excel.Application)
        Dim RibbonXml = GetCustomRibbonXml()
        Return RibbonXml
    End Function

    ' Getter of the embedded Doge button images
    Private Function GetCustomImage(ImageFileName As String) As Sys.Bitmap

        Dim ImageBmp As Sys.Bitmap
        Dim ThisAssembly = GetType(CryptoRibbon).Assembly
        Dim ResourceName = GetType(CryptoRibbon).Namespace & "." & ImageFileName

        Using Stream = ThisAssembly.GetManifestResourceStream(ResourceName)
            ImageBmp = New Sys.Bitmap(Stream)
        End Using
        If ImageBmp Is Nothing Then Throw New MissingManifestResourceException(ResourceName)
        Return ImageBmp

    End Function

    ' Getter of the embedded Ribbon xml resource
    Private Function GetCustomRibbonXml() As String

        Dim RibbonXml As String

        Dim ThisAssembly = GetType(CryptoRibbon).Assembly
        Dim ResourceName = GetType(CryptoRibbon).Namespace & "." & "CryptoRibbon.xml"

        Using Stream = ThisAssembly.GetManifestResourceStream(ResourceName)
            Using Reader = New StreamReader(Stream)
                RibbonXml = Reader.ReadToEnd()
            End Using
        End Using

        If RibbonXml Is Nothing Then Throw New MissingManifestResourceException(ResourceName)
        Return RibbonXml

    End Function

    ' Ribbon start-up
    Public Sub OnLoad(Ribbon As IRibbonUI)

        If Ribbon Is Nothing Then Throw New ArgumentNullException(NameOf(Ribbon))

        ThisRibbon = Ribbon
        AddHandler ThisApp.WorkbookActivate, AddressOf OnInvalidateRibbon
        AddHandler ThisApp.WorkbookDeactivate, AddressOf OnInvalidateRibbon
        AddHandler ThisApp.SheetActivate, AddressOf OnInvalidateRibbon
        AddHandler ThisApp.SheetDeactivate, AddressOf OnInvalidateRibbon
        AddHandler CryptoRTDServer.RTDStatus.ServerStateChanged, AddressOf OnRTDServerStatusChanged
        AddHandler CryptoRTDServer.TelemetryChanged, AddressOf OnRTDServerTelemetryChanged

        If ThisApp.ActiveWorkbook Is Nothing Then ThisApp.Workbooks.Add()

    End Sub

    Private Sub OnInvalidateRibbon(Obj As Object)
        ThisRibbon.Invalidate()
    End Sub

    Public Sub OnPressMe(Control As IRibbonControl)
        If Control.Id = "Toggler" Then
            RaiseEvent DogeClicked()
        End If
    End Sub

    ' Triggers a Doge button refresh whenever ServerStatusChanged
    Public Sub OnRTDServerStatusChanged()
        ThisRibbon.InvalidateControl("Toggler")
        InvalidateTelemetry()
    End Sub

    ' Triggers Ribbon telemetry display recalculation whenever ServerTelemetryChanged
    Public Sub OnRTDServerTelemetryChanged(
        tel As (WeightInterval As Double, WeightLimit As Double, UsedWeight As Integer, TimeBetweenRequests As String))
        LatestTelemetry = tel
        InvalidateTelemetry()
    End Sub

    Public Sub InvalidateTelemetry()
        ThisRibbon.InvalidateControl("WeightLimitLabel")
        ThisRibbon.InvalidateControl("WeightLimitValue")
        ThisRibbon.InvalidateControl("UsedWeightValue")
        ThisRibbon.InvalidateControl("TimeBetweenRequestsValue")
    End Sub

    ' -------------------------------------------------------------------------------------------------------------------------
    ' Functions to actually recalculate the Ribbon when triggered by InvalidateControl() calls
    ' -------------------------------------------------------------------------------------------------------------------------

    ' Update Doge button image (depends on whether there are feeds, cooling down is happening 
    '     And whether it is supposed to be SwitchedOn/Off)
    Public Function OnGetImage(Control As IRibbonControl) As Sys.Bitmap
        If Not CryptoRTDServer.RTDStatus.HasFeeds Then
            Return DogeClosedEyes
        ElseIf CryptoRTDServer.RTDStatus.IsCooling Then
            Return If(CryptoRTDServer.RTDStatus.IsSwitchedOn, DogeOpenEyesHourglass, DogeClosedEyesHourglass)
        Else
            Return If(CryptoRTDServer.RTDStatus.IsSwitchedOn, DogeOpenEyes, DogeClosedEyes)
        End If
    End Function

    ' Update the label below the Doge button (should essentially echo a short version of the ServerStatus)
    Public Function OnGetButtonLabel(Control As IRibbonControl) As String
        Return CryptoRTDServer.RTDStatus.Echo(WithConfigErrorDesc:=False)
    End Function

    ' The request weight limit label, which is just a constant text but with a small parametrized descriptive bit:
    '     the WeightInterval (ie. the time interval over which the defined weight limit applies - eg. a minute)
    Public Function OnGetWeightLimitLabel(Control As IRibbonControl) As String
        Return If(CryptoRTDServer.RTDStatus.IsRunning AndAlso CryptoRTDServer.RTDStatus.IsConnected, String.Format("Request Weight Limit / {0} s :", LatestTelemetry.WeightInterval), "Request Weight Limit :")
    End Function

    ' The request weight limit value
    '     as defined by the API docs And enforced by CryptoRTDServer (should Not change during a Excel session))
    Public Function OnGetWeightLimitValue(Control As IRibbonControl) As String
        Return If(CryptoRTDServer.RTDStatus.IsRunning AndAlso CryptoRTDServer.RTDStatus.IsConnected, LatestTelemetry.WeightLimit.ToString(), "---")
    End Function

    ' The currently used weight (which must not exceed the weight limit during weight interval, 
    '     and which is reset after weight interval)
    Public Function OnGetUsedWeightValue(Control As IRibbonControl) As String
        Return If(CryptoRTDServer.RTDStatus.IsRunning AndAlso CryptoRTDServer.RTDStatus.IsConnected, LatestTelemetry.UsedWeight.ToString(), "---")
    End Function

    ' The moving-average calculated elapsed time between requests
    '     This aims to be as regular as possible whilst ensuring the weight limit isn't exceeded, so it depends on the
    '     current weight demands being placed on CryptoRTDServer via the Excel feeds
    Public Function OnGetTimeBetweenRequestsValue(Control As IRibbonControl) As String
        Return If(CryptoRTDServer.RTDStatus.IsRunning AndAlso CryptoRTDServer.RTDStatus.IsConnected, LatestTelemetry.TimeBetweenRequests, "---")
    End Function

End Class