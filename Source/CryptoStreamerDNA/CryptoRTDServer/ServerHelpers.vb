Option Explicit On
Option Strict On

Imports System.Runtime.CompilerServices
Imports System.Threading

' This module encapsulates a few core functionalities of the CryptoRTDServer into worker classes which are instantiated 
'     as fields of the CryptoRTDServer
Module ServerHelpers

    ' *************************************************************************************************************************
    ' Helper Class - DiscreteStepper
    ' -------------------------------------------------------------------------------------------------------------------------
    ' This class implements a Stepper counter which starts at 1 and goes up until (and including) a defined Threshold, then 
    ' If it's cyclical, it loops back To 1
    ' If it's non-cyclical, it then cannot advance further than Threshold
    ' 
    ' Both cyclical and acyclical steppers can have their counts reset to 1 (if acyclical, this is needed in order to restart)
    '
    ' The Threshold must be >= 1
    ' Calling Increment() counts up 1 unit (and eventually either stops increasing (acyclical) or loops around to 1 (cyclical))
    ' Re-setting the threshold to a different value after creation is possible, and also resets the count (ie. the 'step') to 1
    '
    ' The stepper can be thought of as a swtich, which is considered to be 'active' when its internal count 
    '     (ie. its current 'step') is in the maximum position (the Threshold) - otherwise, the stepper is 'idle' or 'inactive'
    '
    ' The CryptoRTDServer contains a frequency divider which is implemented through a cyclical instance of this class 
    '     With Threshold = Max(FreqDivBase, FreqDivDamp)
    ' It also contains a 2-step sampling trigger is implemented through a non-cyclical instance of this class 
    '     With Threshold = 3(ie.requiring 2 steps to hit the Threshold And be 'On')
    Public Class DiscreteStepper

        Private InternalCount As Integer

        Public ReadOnly Property IsCyclical As Boolean

        Private _Threshold As Integer
        Public Property Threshold As Integer
            Get
                Return _Threshold
            End Get
            Set(Threshold As Integer)
                If Threshold < 1 Then Throw New ArgumentOutOfRangeException("Threshold", "Must be >= 1")
                _Threshold = Threshold
                InternalCount = 1
            End Set
        End Property

        Public ReadOnly Property IsActivated As Boolean
            Get
                Return Threshold = InternalCount
            End Get
        End Property

        Public Sub New(Threshold As Integer, IsCyclical As Boolean)
            Me.Threshold = Threshold
            Me.IsCyclical = IsCyclical
        End Sub

        ' Counts one up and reset when it surpasses the Threshold (if cyclical)
        Public Sub Increment()
            If Not IsActivated Then
                InternalCount += 1
            ElseIf IsCyclical Then
                Reset()
            End If
        End Sub

        Public Sub Reset()
            InternalCount = 1
        End Sub

    End Class

    ' *************************************************************************************************************************
    ' Helper Class - BijectiveDictionary
    ' -------------------------------------------------------------------------------------------------------------------------
    ' A Bijective Dictionary is a bi-direccional dictionary made up of 2 unidirectional regular dictionaries
    ' This class ensures that the Key-Value pair relationship can be read in both directions
    '     by enforcing unicity And consistency Of all keys And values
    '
    ' This will be used to keep track of the coordinates (Symbol, Metric) that each TopicID tracks:
    ' The TopicID is an integer representing the ID of a Topic: for each such ID, a unique (Symbol, Metric) pair is assigned
    ' No 2 different TopcIds will ever be created by the RTD server which point to the same Symbol and Metric, which is why
    '     the Bijection Dictionary is appropriate to keep overall track of what is being requested by Excel at any given time
    Public Class BijectiveDictionary(Of T, U)

        Private LeftToRight As Dictionary(Of T, U)
        Private RightToLeft As Dictionary(Of U, T)

        Public ReadOnly Property RightKeys As Dictionary(Of U, T).KeyCollection
            Get
                Return RightToLeft.Keys
            End Get
        End Property

        Public ReadOnly Property LeftKeys As Dictionary(Of T, U).KeyCollection
            Get
                Return LeftToRight.Keys
            End Get
        End Property

        Public Sub New()
            LeftToRight = New Dictionary(Of T, U)
            RightToLeft = New Dictionary(Of U, T)
        End Sub

        Public Function ContainsLeftKey(LeftElement As T) As Boolean
            Return LeftToRight.ContainsKey(LeftElement)
        End Function

        Public Function ContainsRightKey(RightElement As U) As Boolean
            Return RightToLeft.ContainsKey(RightElement)
        End Function

        Public Function GetLeftToRight(LeftElement As T) As U
            Return LeftToRight(LeftElement)
        End Function

        Public Function GetRightToLeft(RightElement As U) As T
            Return RightToLeft(RightElement)
        End Function

        Public Function TryAdd(LeftElement As T, RightElement As U) As Boolean

            If LeftToRight.ContainsKey(LeftElement) OrElse RightToLeft.ContainsKey(RightElement) Then
                Return False
            Else
                LeftToRight(LeftElement) = RightElement
                RightToLeft(RightElement) = LeftElement
                Return True
            End If

        End Function

        Public Function TryRemoveLeft(LeftElement As T) As Boolean

            If Not LeftToRight.ContainsKey(LeftElement) Then
                Return False
            Else
                RightToLeft.Remove(LeftToRight(LeftElement))
                LeftToRight.Remove(LeftElement)
                Return True
            End If

        End Function

        Public Function TryRemoveRight(RightElement As U) As Boolean

            If Not RightToLeft.ContainsKey(RightElement) Then
                Return False
            Else
                LeftToRight.Remove(RightToLeft(RightElement))
                RightToLeft.Remove(RightElement)
                Return True
            End If

        End Function

    End Class

    ' *************************************************************************************************************************
    ' Helper Class - MovingAverageCalculator
    ' -------------------------------------------------------------------------------------------------------------------------
    ' Helper class which produces an estimate of TimeToDoWork (which is used by the RTD Server Request Weight Modulator)
    '   - Note that the Request Weight Modulator is a notional concept, in practice this is implemented via routine calls to
    '         GetSuggestedFreqDivBaseThres helper method
    ' 
    ' The CryptoRTDServer is makes N 'idle' calls to burn N * BaseClockPeriod milliseconds, then an active call which does work
    '
    ' The Request Weight Modulator calibrates N in near-time to steer the request weight towards an intended target
    '     eg. (12 / 16 of the maximum allowed per unitary interval (eg. per minute))
    ' In order to do this, we need a reliable average of how long the actual API requests are taking for current Excel feeds
    '
    ' The estimator is configured by 2 quantities: FullWindowSize and MinWindowSize
    '   - FullWindowSize is just the size of the moving average window (eg. 8)
    '   - MinWindowSize is the minimum allowed size of an (incomplete) window for the estimator to allow estimates to be made
    '     In other words, without at least MinWindowSize samples, the estimator refuses to make an estimate, ie. is not 'ready'
    '
    ' This is important because upon server startup, the first few requests may consume a non-representative amount of time
    '     hence it is useful to minimally smooth this out before outputting figures which steer the Request Weight Modulator
    Public Class MovingAverageCalculator

        ' The queue will hold (up to) the FullWindowSize most recent samples
        '     Based on its contents at any given time, the moving average is easily calculated from it
        ' An element of this queue is a time interval representing how long it took to fulfill a certain batch of requests
        Private Samples As Queue(Of Double)

        ' Tracks the sum of all queued sample values
        Public ReadOnly Property RunningTotal As Double

        ' Configuration variables
        Public ReadOnly Property MinWindowSize As Integer
        Public ReadOnly Property FullWindowSize As Integer

        ' Condition necessary for SmoothedValue to be available
        Public ReadOnly Property IsMinimallyReady As Boolean
            Get
                Return Samples.Count >= MinWindowSize
            End Get
        End Property

        Public ReadOnly Property IsFullyReady As Boolean
            Get
                Return Samples.Count = FullWindowSize
            End Get
        End Property

        ' The actual calcualted moving average (only valid if IsMinimallyReady)
        Public ReadOnly Property SmoothedValue As Long
            Get
                If Not IsMinimallyReady Then Throw New InvalidOperationException("Not enough samples added")
                Return CLng(RunningTotal / Samples.Count)
            End Get
        End Property

        ' Configuration varaibles FullWindowSize and MinWindowSize are defined on construction
        Public Sub New(FullWindowSize As Integer, MinWindowSize As Integer)

            If MinWindowSize < 1 Then Throw New ArgumentOutOfRangeException("MinWindowSize", "Must be >= 1")
            If FullWindowSize < MinWindowSize Then Throw New ArgumentOutOfRangeException("FullWindowSize", "Must be >= MinWindowSize")

            Reset()
            Me.FullWindowSize = FullWindowSize
            Me.MinWindowSize = MinWindowSize

        End Sub

        ' To try to spread computation time uniformly (to aid the time regularity of the RTDServer) we use a running total
        '     to constantly keep track of the most recent SMoothedValue
        ' This way, adding a sample costs the same as asking for the SmoothedValue, reducing the assymetry between 
        '     RTDServer cycles which compute the SmoothedValue and cycles which just add samples
        Public Sub AddSample(SampledValue As Double)

            If IsFullyReady Then _RunningTotal -= Samples.Dequeue()

            Samples.Enqueue(SampledValue)
            _RunningTotal += SampledValue

        End Sub

        Public Sub Reset()
            Samples = New Queue(Of Double)
            _RunningTotal = 0
        End Sub

    End Class

    ' *************************************************************************************************************************
    ' Helper Extension methods - Update[Boolean/Integer]State methods and [Boolean/Integer]StateChanged events
    ' -------------------------------------------------------------------------------------------------------------------------
    ' The CryptoRTDServer needs to know when it's 'state' has changed, in the sense of something in the ribbon Doge button
    '     needing to be refreshed
    ' 
    ' The Helper class ServerStatus below encapsulates the 'state' in the above terms, and is expressed via a combination
    '     of boolean flags and integer values
    '
    ' Hence, these extensions provide a general way to automatically generate notification event when a state variable
    '     has been updated (ie. a new value has been set which is actually different from the old value)
    Public Event BooleanStateChanged(NewValue As Boolean)
    Public Event IntegerStateChanged(NewValue As Integer)
    <Extension()>
    Public Sub UpdateBooleanState(ByRef StateToUpdate As Boolean, UpdatedState As Boolean)
        If UpdatedState <> StateToUpdate Then
            StateToUpdate = UpdatedState
            RaiseEvent BooleanStateChanged(UpdatedState)
        End If
    End Sub
    <Extension()>
    Public Sub UpdateIntegerState(ByRef StateToUpdate As Integer, UpdatedState As Integer)
        If UpdatedState <> StateToUpdate Then
            StateToUpdate = UpdatedState
            RaiseEvent IntegerStateChanged(UpdatedState)
        End If
    End Sub

    ' *************************************************************************************************************************
    ' Helper Class - ServerStatus
    ' -------------------------------------------------------------------------------------------------------------------------
    ' The CryptoRTDServer needs to know when it's 'state' has changed, in the sense of something in the ribbon Doge button
    '     needing to be refreshed
    ' 
    ' This Helper class encapsulates the 'state' in the above terms, and is expressed via a combination
    '     of boolean flags and integer values
    Public Class ServerStatus

        ' Is true if at least 1 Excel cell is bound to a CryptoRTDServer feed
        Private _HasFeeds As Boolean
        Public Property HasFeeds As Boolean
            Get
                Return _HasFeeds
            End Get
            Set(value As Boolean)
                _HasFeeds.UpdateBooleanState(value)
            End Set
        End Property

        ' Is true if the ConfigureServer routine (called within the ServerStart routine of the CryptoRTDServer) was successful
        Private _IsConfigured As Boolean
        Public Property IsConfigured As Boolean
            Get
                Return _IsConfigured
            End Get
            Set(value As Boolean)
                _IsConfigured.UpdateBooleanState(value)
                If IsConfigured Then ConfigErrorDesc = ""
            End Set
        End Property

        ' Is true if the Doge button is 'On' (ie. if the Doge has the eyes open)
        Private _IsSwitchedOn As Boolean
        Public Property IsSwitchedOn As Boolean
            Get
                Return _IsSwitchedOn
            End Get
            Set(value As Boolean)
                _IsSwitchedOn.UpdateBooleanState(value)
            End Set
        End Property

        ' Is true if the latest attempt at connecting to the API was successful
        Private _IsConnected As Boolean
        Public Property IsConnected As Boolean
            Get
                Return _IsConnected
            End Get
            Set(value As Boolean)
                _IsConnected.UpdateBooleanState(value)
            End Set
        End Property

        ' Pre-defined cooling time period which the RTDServer imposes in the unlikely event it surpasses teh weight limit
        '     This is the last resource in order to respect the API policy and avoid getting banned
        ' However, the other measures which kick-in before this (the Request Weight Modulation and the weight damping)
        '     are very, very likely to be enough, aiming for consistently-timed, but reactive, API calls
        ' Hence, although activation of the cooling period and its resolution are fully automated and part of normal operation,
        '     if you actually see the RTDServer needing to activate this measure, that means the code was not robust enough
        '         to be sufficiently elastic / reactive in its adjustment of the request weight rates
        Private _CoolingPeriod As Integer
        Public Property CoolingPeriod As Integer
            Get
                Return _CoolingPeriod
            End Get
            Set(value As Integer)
                If value < 0 Then Throw New ArgumentOutOfRangeException("CoolingPeriod", "Must be >= 0")
                _CoolingPeriod = value
                If CoolingTimeLeft > CoolingPeriod Then _CoolingTimeLeft = CoolingPeriod
            End Set
        End Property

        ' Cooling time left at any given moment: it's zero whenever not cooling down
        Private _CoolingTimeLeft As Integer
        Public Property CoolingTimeLeft As Integer
            Get
                Return _CoolingTimeLeft
            End Get
            Private Set(value As Integer)
                _CoolingTimeLeft.UpdateIntegerState(value)
            End Set
        End Property

        ' Timer responsible for fulfilling the cooling down routine
        Private CoolingTimer As Timer
        Public Property ConfigErrorDesc As String

        Public ReadOnly Property IsCooling As Boolean
            Get
                Return CoolingTimeLeft > 0
            End Get
        End Property

        Public ReadOnly Property IsRunning As Boolean
            Get
                Return IsConfigured AndAlso IsSwitchedOn AndAlso Not IsCooling
            End Get
        End Property

        Public Event ServerStateChanged()

        ' CoolingPeriod = 0 would mean no cooling measure at all
        '     in practice, the RTDServer will set CoolingPeriod = WeightInterval
        '     where WeightInterval is whatever time period the Weight Limit definition based on in the API docs
        ' Hence, cooling off during WeightInterval ensures the currently used Request Weight is zero afterwards
        Public Sub New(Optional CoolingPeriod As Integer = 0)

            AddHandler BooleanStateChanged, AddressOf OnInnerStateChanged
            AddHandler IntegerStateChanged, AddressOf OnInnerStateChanged

            CoolingTimer = New Timer(TimerTick, Nothing, Timeout.Infinite, Timeout.Infinite)

            HasFeeds = False
            IsConfigured = False
            ConfigErrorDesc = ""
            CoolingTimeLeft = 0
            IsSwitchedOn = False
            IsConnected = False

            Me.CoolingPeriod = CoolingPeriod

        End Sub

        ' Funneling of all individual state variable change notifications into a single 'ServerStateChanged' event
        '     for the outside world
        ' Note that technically, the CoolingTimeLeft is also an (integer) state variable
        '     hence when cooling down, every second the state changes - we want to flag this as such because we want
        '     the Doge button in the Ribbon to react to this and update the cooling down count accordingly
        Private Sub OnInnerStateChanged(ChangedInto As Object)
            RaiseEvent ServerStateChanged()
        End Sub

        ' Timer which handles the cooling down routine, if cooling down was activated
        '     Counts down the seconds until zero, then stops the timer
        Private TimerTick As TimerCallback =
            Sub()
                If CoolingTimeLeft = 0 Then
                    CoolingTimer.Change(Timeout.Infinite, Timeout.Infinite)
                Else
                    CoolingTimeLeft.UpdateIntegerState(CoolingTimeLeft - 1)
                End If
            End Sub

        ' Upon cooling initiation:
        '     set CoolingTimeLeft to be the CoolingPeriod
        '     Let the timer run (it will call TimerTick callback every second)
        Sub InitiateCooling()
            CoolingTimeLeft.UpdateIntegerState(CoolingPeriod)
            CoolingTimer.Change(1, 1000)
        End Sub

        ' The hierarchized / standardized description of the ServerState, to be shown both below the Doge button in the Ribbon
        '     and also as outputs for the Excel cell feeds when for any reason it is not streaming actual values
        Function Echo(WithConfigErrorDesc As Boolean) As String
            If Not HasFeeds Then
                Return "Sleeping (give me feeds)"
            ElseIf CoolingTimeLeft > 0 Then
                Return String.Format("Cooling ({0}s left)...", CoolingTimeLeft)
            ElseIf Not IsSwitchedOn Then
                Return "Switched Off (click Doge to stream)"
            ElseIf IsConfigured And IsConnected Then
                Return "Streaming..."
            Else
                Return String.Format("Trying to connect...{0}", If(WithConfigErrorDesc AndAlso ConfigErrorDesc <> "", String.Format(" ({0})", ConfigErrorDesc), ""))
            End If
        End Function

    End Class

End Module