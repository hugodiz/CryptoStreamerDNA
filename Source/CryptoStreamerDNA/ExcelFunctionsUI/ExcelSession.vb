Option Explicit On
Option Strict On

Imports ExcelDna.Integration
Imports System.Net.Http
Imports Newtonsoft.Json

' This module exposes its defined functions to Excel as Worksheet user-defined functions, callable from Excel cells
Public Module ExcelSession

    ' Work-horse HTTP client for API requests originated in this module: CryptoSymbols() function calls
    Public SessionBinanceAPIClient As HttpClient

    ' Holder of CryptoSymbols() function call results: provides a list of currently featured Binance API crypto symbols
    Public ValidSymbolsDict As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)

    ' Constant words: strealined basic feedback messages for the CryptoStream function
    Public Const EmptySymbolMsg = "Empty Symbol"
    Public Const UnrecognizedSymbolMsg = "Unrecognized Symbol"
    Public Const EmptyMetricMsg = "Empty Metric"
    Public Const UnrecognizedMetricMsg = "Unrecognized Metric"

    ' *************************************************************************************************************************
    ' CryptoStream(symbol, metric) Excel worksheet function
    ' Takes a crypto symbol (eg. "BNBBTC") and a metric (eg. "last_price") and creates a API feed for its near-time data value
    ' -------------------------------------------------------------------------------------------------------------------------
    <ExcelFunction(Description:=
"Streams price And other daily stats on crypto symbols from the Binance API As an automated near-time-data feed 
(inputs are case-insensitive)")>
    Function CryptoStream(
        <ExcelArgument(Description:="Crypto symbol to feed - for a full list, call CryptoSymbols() on a free column)")>
        symbol As String,
        <ExcelArgument(Description:=
"Crypto metric to feed (last price, ask, bid, etc.) - for a full list, call CryptoMetrics() on an free column")>
        metric As String) As Object

        ' Internally, all symbols and metrics are trimmed and case-cleansed by uppercasing
        Dim SymbolSanitized = symbol.Trim().ToUpper()
        If SymbolSanitized = "" Then Return EmptySymbolMsg
        Dim MetricSanitized = metric.Trim().ToUpper()
        If MetricSanitized = "" Then Return EmptyMetricMsg

        ' The available metrics aren't likely to change (if they do, the code must be reviewed against new API docs)
        ' Hence, the metric input is checked against a constant white-list
        If Not CryptoRTDServer.CryptoMetricsDict.ContainsKey(MetricSanitized) Then Return UnrecognizedMetricMsg

        ' Binding a CryptoRTDServer feed to the cell which called this function
        Return XlCall.RTD(CryptoRTDServer.ServerProgId, Nothing, SymbolSanitized, MetricSanitized)

    End Function

    ' *************************************************************************************************************************
    ' CryptoSymbols() Excel worksheet function
    ' Produces a Excel dynamic column array listing all crypto symbols currently featured in the Binance API
    ' -------------------------------------------------------------------------------------------------------------------------
    <ExcelFunction(Description:="Lists all available crypto symbols in the API Binance API to choose from")>
    Function CryptoSymbols() As Object(,)

        ' If this hasn't been populated yet in this Excel session, then
        '     request the last prices for all symbols from the API, to obtain a JSON featuring all symbols as keys
        '     then, deserialize it into a degenerate dictionary of the type (String, Boolean) -> (Symbol, True)
        If ValidSymbolsDict.Count = 0 Then
            Try
                ValidSymbolsDict =
                    JsonConvert.
                        DeserializeObject(Of IEnumerable(Of Dictionary(Of String, String)))(SessionBinanceAPIClient.
                            GetAsync(CryptoRTDServer.LastPriceBaseUrl).Result.Content.ReadAsStringAsync().Result).
                                ToDictionary(Function(x As Dictionary(Of String, String)) x("symbol"), Function() True)

                If ValidSymbolsDict.Count = 0 Then Return {{"symbol list appears to be empty - check code against API docs"}}
            Catch
                ' If an exception is thrown, the most likely cause is bad internet connection
                Return {{"failed to get list of all symbols - check Internet connection"}}
            End Try
        End If

        ' Convert ValidSymbolsDict contents into a degenerate 1-column 2D-array (so Excel will see it as a column) 
        Dim SymbolArray(0 To ValidSymbolsDict.Count - 1, 0 To 0) As Object
        Dim i = 0
        For Each Symbol In ValidSymbolsDict.Keys
            SymbolArray(i, 0) = Symbol
            i += 1
        Next

        ' Return to Excel
        Return SymbolArray

    End Function

    ' *************************************************************************************************************************
    ' CryptoMetrics() Excel worksheet function
    ' Produces a Excel dynamic column array listing all crypto metrics (for a given symbol) currently featured in the API
    ' -------------------------------------------------------------------------------------------------------------------------
    <ExcelFunction(Description:=
"Lists all available crypto metrics in the API Binance API to choose from (for a given symbol)")>
    Function CryptoMetrics() As Object(,)

        ' Convert dictionary contents into a degenerate 1-column 2D-array (so Excel will see it as a column)
        Dim Output(0 To CryptoRTDServer.CryptoMetricsDict.Count - 1, 0 To 0) As Object
        Dim i = 0
        For Each k In CryptoRTDServer.CryptoMetricsDict.Keys
            Output(i, 0) = k
            i += 1
        Next

        Return Output

    End Function

End Module
