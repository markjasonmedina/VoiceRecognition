Imports System
Imports Microsoft.Speech.Recognition
Imports Microsoft.Speech.Synthesis
Imports System.Globalization

Module Module1

    Public ss As SpeechSynthesizer = New SpeechSynthesizer()
    Public sre As SpeechRecognitionEngine
    Public done As Boolean = False
    Public speechOn As Boolean = True

    Sub Main()
        Try
            ss.SetOutputToDefaultAudioDevice()
            Console.WriteLine(vbLf & "(Speaking: I am awake)")
            ss.Speak("I am awake")
            Dim ci As CultureInfo = New CultureInfo("en-us")
            sre = New SpeechRecognitionEngine(ci)
            sre.SetInputToDefaultAudioDevice()
            AddHandler sre.SpeechRecognized, AddressOf sre_SpeechRecognized
            Dim ch_StartStopCommands As Choices = New Choices()
            ch_StartStopCommands.Add("on")
            ch_StartStopCommands.Add("speech off")
            ch_StartStopCommands.Add("klatu barada nikto")
            Dim gb_StartStop As GrammarBuilder = New GrammarBuilder()
            gb_StartStop.Append(ch_StartStopCommands)
            Dim g_StartStop As Grammar = New Grammar(gb_StartStop)
            Dim ch_Numbers As Choices = New Choices()
            ch_Numbers.Add("1")
            ch_Numbers.Add("2")
            ch_Numbers.Add("3")
            ch_Numbers.Add("4")
            Dim gb_WhatIsXplusY As GrammarBuilder = New GrammarBuilder()
            gb_WhatIsXplusY.Append("What is")
            gb_WhatIsXplusY.Append(ch_Numbers)
            gb_WhatIsXplusY.Append("plus")
            gb_WhatIsXplusY.Append(ch_Numbers)
            Dim g_WhatIsXplusY As Grammar = New Grammar(gb_WhatIsXplusY)
            sre.LoadGrammarAsync(g_StartStop)
            sre.LoadGrammarAsync(g_WhatIsXplusY)
            sre.RecognizeAsync(RecognizeMode.Multiple)

            While done = False
            End While

            Console.WriteLine(vbLf & "Hit <enter> to close shell" & vbLf)
            Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.ReadLine()
        End Try

    End Sub

    Private Sub sre_SpeechRecognized(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
        Dim txt As String = e.Result.Text
        Dim confidence As Single = e.Result.Confidence
        Console.WriteLine(vbLf & "Recognized: " & txt)
        If confidence < 0.1 Then Return

        If txt.IndexOf("on") >= 0 Then
            Console.WriteLine("Speech is now ON")
            speechOn = True
        End If

        If txt.IndexOf("speech off") >= 0 Then
            Console.WriteLine("Speech is now OFF")
            speechOn = False
        End If

        If speechOn = False Then Return

        If txt.IndexOf("klatu") >= 0 AndAlso txt.IndexOf("barada") >= 0 Then
            CType(sender, SpeechRecognitionEngine).RecognizeAsyncCancel()
            done = True
            Console.WriteLine("(Speaking: Farewell)")
            ss.Speak("Farewell")
        End If

        If txt.IndexOf("What") >= 0 AndAlso txt.IndexOf("plus") >= 0 Then
            Dim words As String() = txt.Split(" "c)
            Dim num1 As Integer = Integer.Parse(words(2))
            Dim num2 As Integer = Integer.Parse(words(4))
            Dim sum As Integer = num1 + num2
            Console.WriteLine("(Speaking: " & words(2) & " plus " & words(4) & " equals " & sum & ")")
            ss.SpeakAsync(words(2) & " plus " & words(4) & " equals " & sum)
        End If
    End Sub

End Module