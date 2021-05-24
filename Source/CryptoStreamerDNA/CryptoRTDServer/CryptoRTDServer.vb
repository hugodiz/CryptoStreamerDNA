Option Explicit On
Option Strict On

Imports Newtonsoft.Json
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.Rtd
Imports System.Threading
Imports System.Net.Http


' *****************************************************************************************************************************
' CryptoRTDServer class - an application of ExcelDNA ExcelRtdServer functionality
' -----------------------------------------------------------------------------------------------------------------------------
' This RTD Server works by being periodically prompted to action whenever the timer completes a Base Clock Period cycle) 
'     During 1 out of N of those base cycles (where N is modulated based on request weight rates), work is actually done:
'       - 'Works' means fulfilling a batch of API requests according to the totality of Excel feeds currently live
'     N is effectivly a frequency divider which (if > 1) imples that a deliberate lag of (N-1)*BaseClockPeriod is imposed
'         in order to reduce the speed of requests when needed, avoiding exceeding the weight limit established by the API docs
'
' The necessary API docs' constraints are read at runtime by this server, during the initial ServerConfigure routine
'
' The total time between request executions is N * (BaseClockPeriod + small residual lag) + TimeToDoWork
'     TimeToDoWork is a calculated average of how long the actual request executions have been taking lately
' This information is then used to continuously correct N to seek the goal of approximately producing a certain request weight
'     per each weight interval. EG. the target weight percentage could be for instance 12/16 of the allowed maximum,
'         and the allowed maximum might be 1200 request weight / minute.
'     N would continuously be updated to values which would steer the total request weight towards that number, every minute
'
' The elasticity of the request rates (ie. the N) is controlled via a negative feedback loop, which:
'     - measures the currently used weight (which resets every WeightInterval And must not exceed WeightLimit)
'     - adjusts the FreqDivTicker Threshold (ie. the N), by calling GetSuggestedFreqDivBaseThres()
<ComVisible(True)>
<ProgId(CryptoRTDServer.ServerProgId)>
Class CryptoRTDServer
    Inherits ExcelRtdServer

    ' *************************************************************************************************************************
    ' Server Fields
    ' -------------------------------------------------------------------------------------------------------------------------

    ' Work-horse for API requests made by CryptoRTDServer
    Private Shared BinanceAPIClient As HttpClient

    ' These will queried from API policy in ServerConfigure()
    Public Shared ReadOnly Property WeightInterval As Long
    Public Shared ReadOnly Property WeightLimit As Integer
    Public Shared ReadOnly Property CurrentlyUsedWeight As Integer = 0

    ' Current and previous used weight holders (between consecutive request cycles) - used by GetSuggestedFreqDivBaseThres
    Private Shared CurrW As Integer = 0
    Private Shared PrevW As Integer = 0

    ' Timers / Stopwatches to sample relevant time periods:
    Private Shared Tm As Timer ' Base Server Clock (main heartbeat), responsible for every callback
    Private Shared Tw As Stopwatch ' Measurer of the effective time a batch of requests is taking (TimeToDoWork)

    ' Bijection Dictionary keeping track of what topic Id is feeding what: [Integer] Id <-> [String] (Symbol, Metric)
    Private Shared CryptoTopicTracker As BijectiveDictionary(Of Integer, (String, String))

    ' Implementation of the estimator for TimeToDoWork as a MovingAverageCalculator
    Private Shared TimeToDoWorkEstimator As MovingAverageCalculator
    Public Shared ReadOnly Property AvgTimeBtwRequestsFeed As String
        Get
            If TimeToDoWorkEstimator.IsMinimallyReady Then
                Try
                    Return Math.Round((TimeToDoWorkEstimator.SmoothedValue + BaseClockPeriod * FreqDivTicker.Threshold) / 1000, 1).ToString() + " s"
                Catch ex As Exception
                    Return "resolving"
                End Try
            Else
                Return "waiting for samples"
            End If
        End Get
    End Property

    ' 2-step trigger to deliberately throw away the first would-be sample of TimeToDoWork 
    '     (immediately after turning the CryptoStream On) - likely to be a non-representative outlier
    Private Shared SamplingTrigger As DiscreteStepper = New DiscreteStepper(Threshold:=3, IsCyclical:=False)

    ' Percentage of used weight from which damping action (FreqDivDamp) starts
    Private Shared DampWeightPercTrigger As Double

    ' Base candidate for the frequency divider value (N), based on GetSuggestedFreqDivBaseThres() outputs
    Private Shared FreqDivBase As Integer = 4

    ' Secondary candidate for the frequency divider value (N), applying a damping protection mechanism if (and only if) 
    '     the currently used request weight becomes close to the allowed limit, even after the weight modulation'
    ' This overrides FreqDivBase (and of course FreqDivDamp > FreqDivBase)
    Private Shared FreqDivDamp As Integer = 4

    ' Structure holding the actual frequency divider, after taking FreqDivBase and FreqDivDamp into account
    ' This structure is a DiscreteStepper which effectively marks the compass for which Tm callbacks actually do work
    ' The Threshold of this DiscreteSteper is the N discussed above in the context of weight modulation
    Private Shared FreqDivTicker As DiscreteStepper = New DiscreteStepper(Threshold:=FreqDivBase, IsCyclical:=True)

    ' Instance of tracker class for the global state of the server, responsible for raising events handled by the UI Ribbon
    Public Shared RTDStatus As ServerStatus = New ServerStatus()

    ' Data structure used by the server to actually implements the feeds 
    '   - the CryptoTopicTracker will keep track of more detailed information regarding what's held in each TopicId,
    '       which is necessary for the CryptoRTDServer to manage its workload and time the requests properly
    Public Shared Topics As List(Of Topic)

    ' *************************************************************************************************************************
    ' Server Constants
    ' -------------------------------------------------------------------------------------------------------------------------

    ' Server Id for internal calling by Excel feed streamer functions
    Public Const ServerProgId = "CryptoStreamerDNA.CryptoRTDServer"

    ' Base period between Tm callbacks (each 'idle' call (ie. N-1 out of N) will take approximately BaseClockPeriod + overhead)
    '     This is not perfectly rigorous and the choice of 250 ms is basically because it's small enough to allow fine tuning
    '         (otherwise a small change in N would correspond to a big step in the time between requests)
    '         but also and large enough that the true duration of an 'idle' call is percentually close to the nominal 250 ms
    '             so as to make the N * BaseClockPeriod + TimeToDoWork formula sufficiently reliable
    Private Const BaseClockPeriod = 250 ' in milliseconds

    ' Known a priori from Binance API documentation
    ' These quantities are critical for many-simultaneous symbol requests, because there is a fixed request weight
    '     for requests asking for all symbols, which quickly becomes preferrable to many individual requests
    ' E.g. if we have 20 symbol feeds, then it is better to ask for all symbol info (weight 40) every 3 seconds, 
    '     than to try and request each individual symbol (weight 1 * 20) every 1.5 seconds, because the TimeToDoWork of the
    '         individual requests will be significantly greater than the 3 seconds for the average case
    ' The rule of thumb is that we assume it's better to 'ask in bulk' when the individual requests' weight is more than
    ' 1/3 of the bulk request. However, this can be tweaked below (look for the IndividualVsBulk tag below)
    Public Const OtherStatBulkRequestWeight = 40 ' How much it costs to request all symbols's Other Stats at once
    Public Const LastPriceBulkRequestWeight = 2 ' How much it costs to request all symbols's Last Prices at once

    ' Target used weight percentage per WeightInterval that the server aims to achieve (12 / 16 is quite conservative) 
    Private Const TargetWeightPerc = 12 / 16

    ' Multiplier (between 1 and DumpLevels) to apply to FreqDivBase when damping (closer to WeightLimit = higher multiplier)
    '     Higher means stronger damping for the same currently used weights, once we go into the red zone (the damping zone)
    Private Const DampLevels = 8

    ' Full and Min window sizes for the TimeToDoWork estimator moding average (used by TimeToDoWorkEstimator)
    Private Const TimeToDoWorkMovAvgFullWindow = 8
    Private Const TimeToDoWorkMovAvgMinWindow = 3

    ' API end-points and query stems
    Public Const LastPriceBaseUrl = "https://api.binance.com/api/v3/ticker/price"
    Public Const OtherStatBaseUrl = "https://api.binance.com/api/v3/ticker/24hr"
    Public Const ExchangeInfoBaseUrl = "https://api.binance.com/api/v3/exchangeInfo"
    Public Const PingBaseUrl = "https://api.binance.com/api/v3/ping"
    Public Const SymbolQueryStem = "?symbol="

    ' Constant Words: Translation between Excel user syntax when calling CryptoStream, and the known API JSON keys for response reading
    Public Shared ReadOnly CryptoMetricsDict As Dictionary(Of String, String) = New Dictionary(Of String, String) From
    {
        {"PRICE_CHANGE", "priceChange"},
        {"PRICE_CHANGE_PERCENT", "priceChangePercent"},
        {"WEIGHTED_AVERAGE_PRICE", "weightedAvgPrice"},
        {"PREVIOUS_CLOSE_PRICE", "prevClosePrice"},
        {"LAST_PRICE", "price"},
        {"LAST_QUANTITY", "lastQty"},
        {"BID_PRICE", "bidPrice"},
        {"BID_QUANTITY", "bidQty"},
        {"ASK_PRICE", "askPrice"},
        {"ASK_QUANTITY", "askQty"},
        {"OPEN_PRICE", "openPrice"},
        {"HIGH_PRICE", "highPrice"},
        {"LOW_PRICE", "lowPrice"},
        {"VOLUME", "volume"},
        {"QUOTE_VOLUME", "quoteVolume"}
    }

    ' *************************************************************************************************************************
    ' Server Event / Handler-propagator for outside world
    ' -------------------------------------------------------------------------------------------------------------------------
    ' Raised every time actual work is done and the telemetry information is updated here in the server
    Public Shared Event TelemetryChanged(Telemetry As (WeightInterval As Double, WeightLimit As Double, UsedWeight As Integer, TimeBetweenRequests As String))

    ' Handled whenever the ServerStatus class notifies that something in the ServerState has changed
    Private Sub UpdateFeedsOnStateChanged()
        If RTDStatus.HasFeeds Then
            For Each tp In Topics
                tp.UpdateValue(RTDStatus.Echo(WithConfigErrorDesc:=True))
            Next
        End If
    End Sub

    ' Handled whenever the Doge button is clicked in the Ribbon
    '     Note that toggling RTDStatus.IsSwitchedOn then trigger a ServerStateChange, which the Ribbon will pick up
    '         and update the Ribbon elements accordingly (the Doge Button and label)
    Private Sub OnDogeClicked()
        RTDStatus.IsSwitchedOn = Not RTDStatus.IsSwitchedOn

        ' IsConnected is a state component which is always tentative by design
        '     it 's OK to keep saying it's 'False until proven True' whenever streaming stops - this makes the feed
        '         messages more reactive and intelligible (ie. no hanging frozen values when feed is paused)
        RTDStatus.IsConnected = False
    End Sub

    ' *************************************************************************************************************************
    ' Server Methods
    ' -------------------------------------------------------------------------------------------------------------------------

    ' Server Initialization function, occurs when the first Excel cell calls the CryptoStream() worksheet function
    '     This essentially acts as though Excel has just started up, making no assumptions about what is already initialized
    Protected Overrides Function ServerStart() As Boolean

        AddHandler RTDStatus.ServerStateChanged, AddressOf UpdateFeedsOnStateChanged
        AddHandler CryptoRibbon.DogeClicked, AddressOf OnDogeClicked

        BinanceAPIClient = New HttpClient() With {
            .Timeout = TimeSpan.FromSeconds(15)}

        ' This is a probe which not only establishes connectivity but gathers important information about weight limits
        '     which this server will use to modulate the request rates and comply with the API policy
        ServerConfigure()

        TimeToDoWorkEstimator = New MovingAverageCalculator(TimeToDoWorkMovAvgFullWindow, TimeToDoWorkMovAvgMinWindow)

        ' Trackers of Topics (each TopicId <-> a single Excel cell data feed of a (Symbol, Metric) combo)
        Topics = New List(Of Topic)
        CryptoTopicTracker = New BijectiveDictionary(Of Integer, (String, String))

        ' Timer controller of server callbacks and stopwatch measurer of TimeToDoWork
        Tm = New Timer(PeriodicAction, Nothing, BaseClockPeriod, BaseClockPeriod)
        Tw = New Stopwatch()

        ' Since ServerStart is only called when the first Excel feed is created, we know this must have feeds right now
        RTDStatus.HasFeeds = True

        Return True

    End Function

    ' Server Termination and resource cleanup, when the last Excel cell using the CryptoStream() UDF is deleted / erased
    Protected Overrides Sub ServerTerminate()
        Tm.Change(Timeout.Infinite, Timeout.Infinite)
        BinanceAPIClient.Dispose()
        Tm.Dispose()
        RemoveHandler CryptoRibbon.DogeClicked, AddressOf OnDogeClicked

        ' Since ServerTerminate is only called when the last Excel feed is deleted, we know this must NOT have feeds now
        RTDStatus.HasFeeds = False

        ' ServerStatus is not destroyed here, but essentially reset - the reason it's done amotically is because
        '     we want to preserve the cooling down status and counter, should the last Excel feed be terminated mid-cooling
        ' This way, if we then create a new first Excel feed, this server will still have been cooling in the background
        '     and will only start actually streaming once the cooling has ended
        RTDStatus.IsConfigured = False
        RTDStatus.IsSwitchedOn = False
        RTDStatus.IsConnected = False
    End Sub

    ' Hook-up of new Topic feed, representing a new quantum of (Symbol, Metric) information being fed to a new Excel cell
    '     This is called whenever an Excel cell calls CryptoStream with a new unique (Symbol, Metric) pair
    '         such that no other Excel cell was already calling the same Symbol and Metric
    ' In other words, if 2 Excel cells are bound to the same (redundant) feed, there is no duplication of topics;
    '     ie. both cells feed from the same place
    '     - this is what makes the Bijective Dictionary an appropriate structure to keep track of topic information
    Protected Overrides Function ConnectData(tp As Topic, tpInfo As IList(Of String), ByRef newValues As Boolean) As Object

        Try : Tm.Change(Timeout.Infinite, Timeout.Infinite) : Catch Ex As ObjectDisposedException : End Try

        Topics.Add(tp)
        CryptoTopicTracker.TryAdd(tp.TopicId, (tpInfo(0), tpInfo(1)))

        Try : Tm.Change(BaseClockPeriod, BaseClockPeriod) : Catch Ex As ObjectDisposedException : End Try

        Return RTDStatus.Echo(WithConfigErrorDesc:=True)

    End Function

    ' Break-up of Topic feed:
    '     This is called whenever an Excel cell calling CryptoStream is deleted / erased
    '         such that no other Excel cell exists which is calling that same Symbol and Metric
    Protected Overrides Sub DisconnectData(tp As Topic)

        Try : Tm.Change(Timeout.Infinite, Timeout.Infinite) : Catch Ex As ObjectDisposedException : End Try
        CryptoTopicTracker.TryRemoveLeft(tp.TopicId)
        Topics.Remove(tp)

        Try : Tm.Change(BaseClockPeriod, BaseClockPeriod) : Catch Ex As ObjectDisposedException : End Try

    End Sub

    ' Main server heartbeat routine, performed whenever the Timer finishes a base cycle (ie. every 250 ms)
    '     As discussed above, 1 out of N times, FreqDivTicker sill be 'activated' and do actual work
    '         where N = FreqDivTicker.Threshold
    Private PeriodicAction As TimerCallback =
        Sub()
            ' Each call, cycling from 1 to N -> back to 1, through each of the N DiscreteStepper states
            '     when count = N, FreqDivTicker 'is activated'
            FreqDivTicker.Increment()

            ' Stopping Timer to do work, then re-start by the end of this routine
            '     Heartbeat stops during PeriodAction, to ensure no overlapping calls
            '     - This is why the total time between requests is approximately N * BaseClockPeriod + TimeToDoWork
            ' Sometimes a residual action will take place here after ServerTerminate, depending on timing
            '     If so, trying to use the disposed Tm generates an inconsequential exception which is caught and ignored
            Try : Tm.Change(Timeout.Infinite, Timeout.Infinite) : Catch Ex As ObjectDisposedException : End Try

            ' Always check and configure server first if it's not configured already
            If Not RTDStatus.IsConfigured Then ServerConfigure()

            ' Tw measures the actual TimeToDoWork, when the heartbeat Tm is stopped 
            Tw.Start()

            ' Early-return if the streamer isn't running
            '     No material work done if CryptoStreamer is Off (in which case, early return)
            '     the CryptoStreamer state can be controlled via the ControlRibbon Doge button and 
            '     it can also undergo an emergency cut-off if the request weight becomes dangerously close to the Binance documented limits
            If Not RTDStatus.IsRunning Then
                For Each tp In Topics
                    tp.UpdateValue(RTDStatus.Echo(WithConfigErrorDesc:=True))
                Next
                SamplingTrigger.Reset()
                Tw.Stop()
                Try : Tm.Change(BaseClockPeriod, BaseClockPeriod) : Catch Ex As ObjectDisposedException : End Try
                Return
            End If

            ' Early return if the FreqDivTicker is not activated (ie. N-1 times out of N)
            '     No material work done if this is one of the Frequency Divider 'off-beats'
            If Not FreqDivTicker.IsActivated Then
                Tw.Stop()
                Try : Tm.Change(BaseClockPeriod, BaseClockPeriod) : Catch Ex As ObjectDisposedException : End Try
                Return
            End If

            ' From here downwards, 'work' is done -----------------------------------------------------------------------------
            Dim CurrentlyUsedWeightTemp = CurrentlyUsedWeight ' helper variable to refresh currently used weight from API

            ' The effect of this sampling trigger is discarding the first TimeToDoWork sample after turning the streamer On
            SamplingTrigger.Increment()

            ' Assessing the workload for this cycle:
            ' Counting the total number of LastPrice and OtherStat Requests
            Dim LastPriceRequests = New Dictionary(Of String, Boolean)
            Dim OtherStatRequests = New Dictionary(Of String, Boolean)
            For Each tp In Topics
                Dim tpCoords As (Symbol As String, Metric As String) = CryptoTopicTracker.GetLeftToRight(tp.TopicId)
                If tpCoords.Metric.ToUpper() = "LAST_PRICE" Then
                    LastPriceRequests(tpCoords.Symbol.ToUpper()) = True
                Else
                    OtherStatRequests(tpCoords.Symbol.ToUpper()) = True
                End If
            Next

            ' START of SubSection: Make API Request for Last Price ------------------------------------------------------------
            Dim LastPriceRequestList = LastPriceRequests.Keys.ToList()
            Dim LastPriceRaw As HttpResponseMessage
            Dim LastPriceDict As Dictionary(Of String, Dictionary(Of String, String))

            ' IndividualVsBulk:
            ' Decides whether to ask for Last Price data in bulk or do K individual requests for each symbol
            If LastPriceRequests.Count > Math.Max(1, LastPriceBulkRequestWeight \ 3) Then
                Try
                    LastPriceRaw = BinanceAPIClient.GetAsync(LastPriceBaseUrl).Result
                    CurrentlyUsedWeightTemp = Integer.Parse(LastPriceRaw.Headers.GetValues("x-mbx-used-weight")(0))
                    LastPriceDict =
                    JsonConvert.
                        DeserializeObject(Of IEnumerable(Of Dictionary(Of String, String)))(LastPriceRaw.Content.ReadAsStringAsync().Result).
                            ToDictionary(Function(x As Dictionary(Of String, String)) x("symbol"))
                    RTDStatus.IsConnected = True
                Catch
                    Tw.Stop()
                    Tw.Reset()
                    LastPriceDict = New Dictionary(Of String, Dictionary(Of String, String))
                    RTDStatus.IsConnected = False
                End Try

            Else
                LastPriceDict = New Dictionary(Of String, Dictionary(Of String, String))
                Try
                    For Each Symbol In LastPriceRequests.Keys
                        LastPriceRaw = BinanceAPIClient.GetAsync(LastPriceBaseUrl & SymbolQueryStem & Symbol).Result
                        CurrentlyUsedWeightTemp = Integer.Parse(LastPriceRaw.Headers.GetValues("x-mbx-used-weight")(0))
                        LastPriceDict(Symbol) = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(LastPriceRaw.Content.ReadAsStringAsync().Result)
                        RTDStatus.IsConnected = True
                    Next
                Catch
                    Tw.Stop()
                    Tw.Reset()
                    RTDStatus.IsConnected = False
                End Try

            End If

            ' Cooling emergency measure, kicks-in only if the next batch of requests (which will be of OtherStats) 
            '     may possibly exceed the weight limit
            _CurrentlyUsedWeight = CurrentlyUsedWeightTemp
            If CurrentlyUsedWeight >= WeightLimit - OtherStatBulkRequestWeight Then
                RTDStatus.InitiateCooling()
                _CurrentlyUsedWeight = 0
                RTDStatus.IsConnected = False
                Tw.Stop()
                Try : Tm.Change(BaseClockPeriod, BaseClockPeriod) : Catch Ex As ObjectDisposedException : End Try
                Return
            End If
            ' END Of SubSection: Make API Request for Last Price --------------------------------------------------------------

            ' START of SubSection: Make API Request for Other Stats -----------------------------------------------------------
            Dim OtherStatRequestList = OtherStatRequests.Keys.ToList()
            Dim OtherStatRaw As HttpResponseMessage
            Dim OtherStatDict As Dictionary(Of String, Dictionary(Of String, String))

            ' IndividualVsBulk:
            ' Decides whether to ask for Last Price data in bulk or do K individual requests for each symbol
            If OtherStatRequests.Count > Math.Max(1, OtherStatBulkRequestWeight \ 3) Then
                Try
                    OtherStatRaw = BinanceAPIClient.GetAsync(OtherStatBaseUrl).Result
                    CurrentlyUsedWeightTemp = Integer.Parse(OtherStatRaw.Headers.GetValues("x-mbx-used-weight")(0))
                    OtherStatDict =
                        JsonConvert.
                            DeserializeObject(Of IEnumerable(Of Dictionary(Of String, String)))(OtherStatRaw.Content.ReadAsStringAsync().Result).
                                ToDictionary(Function(x As Dictionary(Of String, String)) x("symbol"))
                    RTDStatus.IsConnected = True
                Catch
                    Tw.Stop()
                    Tw.Reset()
                    OtherStatDict = New Dictionary(Of String, Dictionary(Of String, String))
                    RTDStatus.IsConnected = False
                End Try

            Else
                OtherStatDict = New Dictionary(Of String, Dictionary(Of String, String))
                Try
                    For Each Symbol In OtherStatRequests.Keys
                        OtherStatRaw = BinanceAPIClient.GetAsync(OtherStatBaseUrl & SymbolQueryStem & Symbol).Result
                        CurrentlyUsedWeightTemp = Integer.Parse(OtherStatRaw.Headers.GetValues("x-mbx-used-weight")(0))
                        OtherStatDict(Symbol) = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(OtherStatRaw.Content.ReadAsStringAsync().Result)
                        RTDStatus.IsConnected = True
                    Next
                Catch
                    Tw.Stop()
                    Tw.Reset()
                    RTDStatus.IsConnected = False
                End Try

            End If

            ' Cooling emergency measure, kicks-in only if the next batch of requests (which will be of Last Price) 
            '     may possibly exceed the weight limit
            _CurrentlyUsedWeight = CurrentlyUsedWeightTemp
            If CurrentlyUsedWeight >= WeightLimit - LastPriceBulkRequestWeight Then
                RTDStatus.InitiateCooling()
                _CurrentlyUsedWeight = 0
                RTDStatus.IsConnected = False
                Tw.Stop()
                Try : Tm.Change(BaseClockPeriod, BaseClockPeriod) : Catch Ex As ObjectDisposedException : End Try
                Return
            End If
            ' END Of SubSection: Make API Request for Other Stats -------------------------------------------------------------

            ' START of SubSection: Request Weight Modulation ------------------------------------------------------------------
            If RTDStatus.IsConnected Then
                Dim CurrentlyUsedWeightPerc = CurrentlyUsedWeight / WeightLimit

                PrevW = CurrW
                CurrW = CurrentlyUsedWeight

                Tw.Stop()
                Dim WorkTimeElapsed = Tw.ElapsedMilliseconds()

                ' Acquisition of TimeToDoWork sample and restart of the Tw stopwatch, to start working towards next sample
                If WorkTimeElapsed > 0 AndAlso SamplingTrigger.IsActivated Then TimeToDoWorkEstimator.AddSample(WorkTimeElapsed)
                Tw.Restart()

                ' Calculation of a new N-value suggestion if the TimeToDoWorkEstimator is minimally ready
                If TimeToDoWorkEstimator.IsMinimallyReady Then
                    FreqDivBase = GetSuggestedFreqDivBaseThres(CurrW, PrevW, TimeToDoWorkEstimator.SmoothedValue)
                    FreqDivTicker.Threshold = FreqDivBase
                End If

                ' Application of (eventual) damping on top of the base weight modulation
                '     if the currently used weight percentage is deemed too high already
                '     - note this is independent (and comes before) the hard cut-off represented by the cooling down routine
                If CurrentlyUsedWeightPerc > DampWeightPercTrigger Then
                    FreqDivDamp = FreqDivBase * 2 * (CInt(Math.Floor((CurrentlyUsedWeightPerc - DampWeightPercTrigger) / ((1 - DampWeightPercTrigger) / DampLevels))) + 1)
                    FreqDivTicker.Threshold = FreqDivDamp
                End If

            Else ' (if not IsConnected, reset these control variables in preparation to start anew when we're back online)
                TimeToDoWorkEstimator = New MovingAverageCalculator(TimeToDoWorkMovAvgFullWindow, TimeToDoWorkMovAvgMinWindow)
                FreqDivBase = 4
                FreqDivTicker.Threshold = FreqDivBase
            End If
            ' End Of SubSection: Request Weight Modulation --------------------------------------------------------------------

            ' SubSection: Normal-operation periodic Topic value updates -------------------------------------------------------

            ' Telemetry update notification (this prompts the ribbon to re-read the telemetry from this server)
            RaiseEvent TelemetryChanged((WeightInterval:=WeightInterval / 1000, WeightLimit:=WeightLimit, UsedWeight:=CurrentlyUsedWeight, TimeBetweenRequests:=AvgTimeBtwRequestsFeed))

            ' Update of all Excel feeds
            If RTDStatus.IsRunning Then
                For Each tp In Topics

                    Dim tpCoords As (Symbol As String, Metric As String) = CryptoTopicTracker.GetLeftToRight(tp.TopicId)

                    Select Case tpCoords.Metric
                        Case "LAST_PRICE"
                            tp.UpdateValue(GetTopicInfoFromDeserializedJSON(tpCoords, LastPriceDict))
                        Case Else
                            tp.UpdateValue(GetTopicInfoFromDeserializedJSON(tpCoords, OtherStatDict))
                    End Select

                Next
            End If
            ' End Of SubSection: Normal-operation periodic Topic value updates ------------------------------------------------

            ' Periodic Action closing: Tw analysisc stowatch stops, Base Clock starts again
            Tw.Stop()
            Try : Tm.Change(BaseClockPeriod, BaseClockPeriod) : Catch Ex As ObjectDisposedException : End Try
            Return

        End Sub

    ' *************************************************************************************************************************
    ' Internal Helper Function - GetTopicInfoFromDeserializedJSON
    ' -------------------------------------------------------------------------------------------------------------------------
    ' Work-horse function for picking actual Topic concent from API Json responses 
    '   - assumes the Json has already been deserialized and takes the shape of a 1-level nested string-valued dictionary
    '   - provides a streamlined way of trying to pick the correct feed values from the JSON response
    '         and generates error information to give back to the Excel feeds when it can't
    Private Function GetTopicInfoFromDeserializedJSON(CoordinatesToSearch As (Symbol As String, Metric As String), ByRef DictionaryToSearchIn As Dictionary(Of String, Dictionary(Of String, String))) As String

        If Not RTDStatus.IsConnected Then Return "Trying to connect..."

        Dim DictOut As Dictionary(Of String, String)
        Dim DictOutSuccess = DictionaryToSearchIn.TryGetValue(CoordinatesToSearch.Symbol, DictOut)
        If DictOutSuccess Then
            Dim ValOut As String
            Dim ValOutSuccess = DictOut.TryGetValue(CryptoMetricsDict(CoordinatesToSearch.Metric), ValOut)
            If ValOutSuccess Then
                Return ValOut
            Else
                Dim ErrorOut As String
                Dim ErrorOutSuccess = DictOut.TryGetValue("code", ErrorOut)
                If ErrorOutSuccess Then
                    Return String.Format("API Error: code = {0}", ErrorOut)
                Else
                    Return String.Format("could not find metric {0} nor a proper error code in API response keys", CryptoMetricsDict(CoordinatesToSearch.Metric))
                End If
            End If
        Else
            Return String.Format("searching for {0}...", CoordinatesToSearch.Symbol)
        End If

    End Function

    ' *************************************************************************************************************************
    ' Internal Helper Function - GetSuggestedFreqDivBaseThres (essentially, this is the Request Weight Modulator)
    ' -------------------------------------------------------------------------------------------------------------------------
    ' This function is just a calculator which suggests an optimal frequency divider (the N) 
    '     based on data from the latest requests made by this server
    ' The input data consists of the current and previously used (accumulated) request weights and the latest estimate 
    ' of the average time to process a request (TimeToDoWork)
    ' It also relies on the CryptoRTDServer configuration for the base clock frequency and the target used weight
    Private Function GetSuggestedFreqDivBaseThres(CurrW As Integer, PrevW As Integer, TimeToDoWork As Long) As Integer
        If Not CurrW > PrevW Then Return FreqDivBase
        Return Math.Max(1, CInt(((CurrW - PrevW) * WeightInterval / (WeightLimit * TargetWeightPerc) - TimeToDoWork) / BaseClockPeriod))
    End Function


    ' *************************************************************************************************************************
    ' Internal Helper Sub - ServerConfigure
    ' -------------------------------------------------------------------------------------------------------------------------
    ' Part of the server initialization routine (called by ServerStart) 
    '   - obtains critical configuration information for the operation of the server, from the Binance API
    ' It serves the dual purpose of establishing connectivity and learning the weight limit policy of the API
    '   - Gives back error information to Excel feeds and ribbon, if the routine fails (most likely due to internet connection)
    Private Sub ServerConfigure()

        Try
            Dim Response = BinanceAPIClient.GetAsync(LastPriceBaseUrl).Result.StatusCode
            If Response = Net.HttpStatusCode.OK Then
                RTDStatus.IsConfigured = True
                RTDStatus.IsConnected = True
            Else
                RTDStatus.IsConfigured = False
                RTDStatus.IsConnected = True
                RTDStatus.ConfigErrorDesc = String.Format("failed to connect to api.binance.com - {0}", Response.ToString())
            End If
        Catch Ex As Exception
            RTDStatus.IsConfigured = False
            RTDStatus.IsConnected = False
            RTDStatus.ConfigErrorDesc = "failed to connect to api.binance.com - check Internet connection"
        End Try

        If RTDStatus.IsConfigured Then
            Try
                Dim ExchangeInfoDict =
                    JsonConvert.
                        DeserializeObject(Of Dictionary(Of String, Object))(BinanceAPIClient.GetAsync(ExchangeInfoBaseUrl).Result.Content.ReadAsStringAsync().Result)
                Dim RateLimitSpecsRaw As String = ExchangeInfoDict("rateLimits").ToString()
                Dim RateLimitSpecs =
                    JsonConvert.
                        DeserializeObject(Of IEnumerable(Of Dictionary(Of String, String)))(RateLimitSpecsRaw)

                Dim PolicyInterval As Long = Int64.MaxValue
                Dim PolicyLimit As Integer = 0

                For Each RateLimitDict In RateLimitSpecs
                    If RateLimitDict("rateLimitType").ToUpper() <> "REQUEST_WEIGHT" Then Continue For

                    Dim ThisIntervalType = RateLimitDict("interval")
                    Dim ThisIntervalNum As Integer = Int32.Parse(RateLimitDict("intervalNum"))
                    Dim ThisLimit As Integer = Int32.Parse(RateLimitDict("limit"))

                    If Not ThisLimit > 0 Then Throw New JsonReaderException("non-positive value in JSON for key 'limit'")

                    Dim ThisInterval As Long
                    Select Case ThisIntervalType.ToUpper()
                        Case "DAY"
                            ThisInterval = 24 * 60 * 60 * 1000 * ThisIntervalNum
                        Case "HOUR"
                            ThisInterval = 60 * 60 * 1000 * ThisIntervalNum
                        Case "MINUTE"
                            ThisInterval = 60 * 1000 * ThisIntervalNum
                        Case "SECOND"
                            ThisInterval = 1000 * ThisIntervalNum
                        Case Else
                            Throw New JsonReaderException("unexpected value in JSON for key 'interval'")
                    End Select

                    If Not ThisInterval > 0 Then Throw New JsonReaderException("non-positive value in JSON for key 'intervalNum'")

                    If ThisInterval < PolicyInterval Then
                        PolicyLimit *= CInt(ThisInterval / PolicyInterval)
                        PolicyInterval = ThisInterval
                    End If

                    If ThisLimit > PolicyLimit Then PolicyLimit = ThisLimit

                Next

                _WeightLimit = PolicyLimit
                _WeightInterval = PolicyInterval
                DampWeightPercTrigger = (WeightLimit - 4 * (OtherStatBulkRequestWeight + LastPriceBulkRequestWeight)) / WeightLimit
                RTDStatus.CoolingPeriod = CInt(WeightInterval \ 1000)

            Catch Ex As Exception
                RTDStatus.IsConfigured = False
                RTDStatus.IsConnected = False
                RTDStatus.ConfigErrorDesc = Ex.Message
            End Try
        End If

    End Sub

End Class
