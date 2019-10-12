Imports System.Speech.Recognition
Imports System.Threading
Imports System.Globalization
Public Class Form1
    Dim SAPI As Object
    Dim Options As Integer = 0
    Dim recog As New SpeechRecognizer
    Dim gram As Grammar
    Dim Line As String
    Dim VoiceRec As Boolean = False

    Public Event SpeechRecognized As _
        EventHandler(Of SpeechRecognizedEventArgs)
    Public Event SpeechRecognitionRejected As _
        EventHandler(Of SpeechRecognitionRejectedEventArgs)
    Dim RecognizableWords As String() = New String() {"Yes", "No"}
    ' word recognised event
    Public Sub recevent(ByVal sender As System.Object,
            ByVal e As RecognitionEventArgs)
        If (e.Result.Text = "Yes") Then
            Select Case Options
                Case 0
                    Line = "Okay, the Online Wallet can be used to easily send money to and receive money from others almost instantly from any location! Would you like to see how it works?"
                    SAPI.speak(Line, 1)
                    Label1.Text = "Our Online Wallet makes money transactions easy."
                    Label2.Visible = False
                    Delay(4)
                    Label2.Visible = True
                    Label2.Text = "You don't have to go to the bank to make transactions"
                    Delay(4)
                    Label3.Visible = True
                    Label3.Text = "Do you want to know how it works?"
                    Options += 1

                Case 1
                    Line = "Great! we'll start with sending money. Let's pretend you have $50 in your wallet, and you want to send $20 to your friend John, but he's in another country. Using Paypal's online wallet, you can directly transfer $20 into his online wallet for him to spend at his leisure"
                    SAPI.speak(Line, 1)
                    Label1.Visible = False
                    Label2.Visible = False
                    Label3.Visible = False
                    Delay(7)
                    Label1.Text = ". . . Let's say you need to send $20 to your friend John."
                    Label1.Visible = True
                    Delay(3.5)
                    Label2.Text = ". . . But he's in another country."
                    Label2.Visible = True
                    Delay(3)
                    Label3.Text = "You can use Paypal's online wallet to do this instantly"
                    Label3.Visible = True
                    Delay(5)
                    Options += 1
                    Button1.Visible = True
                    Balance.Visible = True
                    Money.Visible = True
                    Line = "Do you see the big, navy blue button labeled send money? Try clicking it now."
                    SAPI.speak(Line, 1)
                    Delay(5)
                    Label2.Text = "Click the Send Money button to give John money."
                    Label1.Visible = False
                    Label3.Visible = False

            End Select

        ElseIf (e.Result.Text = "No") Then
            Select Case Options
                Case 0

                    SAPI.speak("That's okay. I understand. Please have an excellent rest of your day. Goodbye!")
                    Options += 1
                    Me.Close()

                Case 1

            End Select
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SAPI = CreateObject("SAPI.spvoice")
        Button1.Visible = False
        Button2.Visible = False
        Balance.Visible = False
        TextBox2.Visible = False
        Money.Visible = False
        Label2.Visible = False
        Label3.Visible = False

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Line = "Welcome to Paypal's Interactive Agent, would you like to learn the benefits of Paypal's online wallet?"
        SAPI.speak(Line, 1)

        Dim words As New Choices(RecognizableWords)
        gram = New Grammar(New GrammarBuilder(words))
        recog.LoadGrammar(gram)
        ' add handlers for the recognition events
        AddHandler recog.SpeechRecognized, AddressOf Me.recevent
        ' enable the recogniser
        recog.Enabled = True
        Label1.Text = "Welcome to PayPal's Interactive Agent!"
        Delay(3)
        Label2.Visible = True
        Label2.Text = "Would you like to learn about our online wallet?"
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Options = 2 Then
            recog.Enabled = False
            Options += 1
            SAPI.speak("Great job! Now that you've selected send money, how much money do you want to send?", 1)
            Button1.Visible = False
            Label1.Visible = True
            Label2.Visible = False
            Label3.Visible = False
            Label1.Text = ". . . Now that you've selected 'Send Money'"
            Delay(3)
            Label2.Visible = True
            Label2.Text = ". . . How much money do you want to send?"
            Delay(3)
            Label3.Visible = True
            Label3.Text = ". . . Click on the text box below and type in the number."
            TextBox2.Visible = True
            Button2.Visible = True

        End If



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Convert.ToInt32(TextBox2.Text) < 50 Then
            Dim Total As Integer = Convert.ToInt32((TextBox2.Text))
            Money.Text = "$" + (50 - Total).ToString
            SAPI.speak("With just a click of a button, you just sent $20. Congratulations, you now know how to send money.", 1)
            Button2.Visible = False
            TextBox2.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Label1.Text = "Congratulations! You did it!"
        Else
            SAPI.speak("The amount you entered is higher than your balance. Let's try again with a lower number.", 1)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Sub Delay(ByVal dblSecs As Double)

        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            Application.DoEvents()
        Loop
    End Sub
End Class
