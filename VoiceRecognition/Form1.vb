Imports System.Speech

Public Class Form1

    'Create the first global variable MyVoice to recognise the new voice each time a person speaks
    Dim WithEvents MyVoice As New Recognition.SpeechRecognitionEngine

    'The first Private Sub - SetColor (it will allow you to action the command and set the background colour to the colour the user has said.
    Private Sub SetColor(ByVal color As System.Drawing.Color)
        Me.BackColor = color
    End Sub

    'Form load - if you have renamed your form, you will need to change the "Form1_Load" to the "NameOfYourForm_Loud
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'activate the default audio device "Mic"
        MyVoice.SetInputToDefaultAudioDevice()

        'Create a var MyGrammar
        Dim MyGrammar As New Recognition.SrgsGrammar.SrgsDocument

        'Create a var MyWordsRule
        Dim MyWordsRule As New Recognition.SrgsGrammar.SrgsRule("words")

        'Create a var MyWordsRule
        Dim MyWords As New Recognition.SrgsGrammar.SrgsOneOf("green", "red", "blue")

        'Add the words I speak onto the system
        MyWordsRule.Add(MyWords)

        'Add the MyWordRule onto the system
        MyGrammar.Rules.Add(MyWordsRule)

        'The location to MyWordRule
        MyGrammar.Root = MyWordsRule

        'When you hear my voice, LoadGrammar
        MyVoice.LoadGrammar(New Recognition.Grammar(MyGrammar))

        'recognise my voice on form load
        MyVoice.RecognizeAsync()

    End Sub


    'recognise my voice every time I speak
    Private Sub reco_RecognizeCompleted(ByVal sender As Object, ByVal e As System.Speech.Recognition.RecognizeCompletedEventArgs) Handles MyVoice.RecognizeCompleted

        MyVoice.RecognizeAsync()

    End Sub
    Dim SAPI


    'recognise my voice and if the case exists, execute the procedure
    Private Sub reco_SpeechRecognized(ByVal sender As Object, ByVal e As System.Speech.Recognition.RecognitionEventArgs) Handles MyVoice.SpeechRecognized
        SAPI = CreateObject("SAPI.spvoice")
        Console.WriteLine(e.Result.Text)


        Select Case e.Result.Text
            'if the user speaks "Blue"
            Case "red"
                'The Background of the Form will change to Blue
                SetColor(Color.Red)
                Label1.Visible = True
                Label1.Text = "Red 'to"

            'if the user speaks "Green"
            Case "green"
                'The Background of the Form will change to Green
                SetColor(Color.Green)
                Label1.Visible = True
                Label1.Text = "Green 'to"


            Case "blue"
                'The Background of the Form will change to blue
                SetColor(Color.Blue)
                Label1.Visible = True
                Label1.Text = "Blue 'to"


        End Select

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class