Imports System
Imports System.Globalization
Imports System.IO.Ports  'This library is for connecting vb.net and Arduino to transmit and receive data through ports
'Below are libraries for voice recognition
Imports Microsoft.Speech.Recognition
Imports Microsoft.Speech.Synthesis


Module Module1
    Private myPort As SerialPort = New SerialPort()
    Public ss As SpeechSynthesizer = New SpeechSynthesizer()
    Public sre As SpeechRecognitionEngine
    Public done As Boolean = False
    Public speechOn As Boolean = True

    Sub Main()
        myPort.PortName = "COM7"
        myPort.BaudRate = 9600
        Try
            ss.SetOutputToDefaultAudioDevice()
            Console.WriteLine(vbLf & "(Speaking: I am awake)")
            ss.Speak("I am awake")

            Dim ci As CultureInfo = New CultureInfo("en-us")
            sre = New SpeechRecognitionEngine(ci)
            sre.SetInputToDefaultAudioDevice()

            AddHandler sre.SpeechRecognized, AddressOf sre_SpeechRecognized

            Dim ch_Numbers As Choices = New Choices()
            ch_Numbers.Add("one")
            ch_Numbers.Add("two")
            ch_Numbers.Add("six")
            ch_Numbers.Add("apple")

            Dim str_OnOff As Choices = New Choices()
            str_OnOff.Add("up")
            str_OnOff.Add("down")

            Dim gb_SwitchOnOff As GrammarBuilder = New GrammarBuilder()
            gb_SwitchOnOff.Append("Switch")
            gb_SwitchOnOff.Append(str_OnOff)
            gb_SwitchOnOff.Append(ch_Numbers)
            Dim g_SwitchOnOff As Grammar = New Grammar(gb_SwitchOnOff)

            sre.LoadGrammarAsync(g_SwitchOnOff)
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
        'If confidence < 0.2 Then Return

        'If speechOn = False Then Return

        If txt.IndexOf("Switch") >= 0 Then
            Dim words As String() = txt.Split(" "c)
            'Dim num1 As Integer = Integer.Parse(words(2))

            If (words(1) = "up" AndAlso words(2) = "one") Then
                sendDataToArduino("B"c)
            End If

            If (words(1) = "down" AndAlso words(2) = "one") Then
                sendDataToArduino("Z"c)
            End If

            If (words(1) = "up" AndAlso words(2) = "two") Then
                sendDataToArduino("R"c)
            End If

            If (words(1) = "down" AndAlso words(2) = "two") Then
                sendDataToArduino("X"c)
            End If

            If (words(1) = "up" AndAlso words(2) = "six") Then
                sendDataToArduino("G"c)
            End If

            If (words(1) = "down" AndAlso words(2) = "six") Then
                sendDataToArduino("C"c)
            End If

            If (words(1) = "up" AndAlso words(2) = "apple") Then
                sendDataToArduino("V"c)
            End If

            If (words(1) = "down" AndAlso words(2) = "apple") Then
                sendDataToArduino("M"c)
            End If

            Console.WriteLine("(Speaking: Device " & words(2) & " is " & words(1) & ")")
            ss.SpeakAsync("Device " & words(2) & " is " & words(1))
        End If


    End Sub

    Private Sub sendDataToArduino(ByVal character As Char)
        myPort.Open()
        myPort.Write(character.ToString())
        myPort.Close()
    End Sub

End Module