Imports MyGames.Tetris
Imports MyGames.Tetris.TetrisBlock
Imports System.Reflection
Imports System.IO
Partial Class TetrisGame
    Private GameBoard As TetrisBoard
    Private FallingBlock As TetrisBlock
    Private PreviewBoard As TetrisBoard
    Private PreviewBlock As TetrisBlock
    Dim playername As String
    Dim scores(11, 3) As String
    Private Score As Double
    Private Level As Integer
    Private Speed As Integer
    Private Lines As Integer
    Private RandomNumbers As New Random
    Private Status As GameStatus = GameStatus.Stopped
    Private Enum GameStatus
        Running
        Paused
        Stopped
    End Enum
#Region "Event Handlers"
    'This is the event handlers code
    Private Sub TetrisGame_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PreviewBoard = New TetrisBoard(PreviewBox_1)
        With PreviewBoard
            .Rows = 4
            .Columns = 4
            .CellSize = New Size(20, 20)
            .Style = BorderStyle.FixedSingle
            .SetupBoard()
        End With
        PreviewBlock = New TetrisBlock(PreviewBoard)
        PreviewBlock.CenterCell = PreviewBoard.Cells(2, 2)
        PreviewBlock.Shape = GetRandomShape()
        GameBoard = New TetrisBoard(GameBox)
        With GameBoard
            .Rows = 20
            .Columns = 10
            .CellSize = New Size(20, 20)
            .Style = BorderStyle.FixedSingle
            .SetupBoard()
        End With
        'In the above code the preview board is loaded in as well as the size of the game board
        FallingBlock = New TetrisBlock(GameBoard)
        HelpLabel.Text = HelpLabel.Text.Replace("|", vbCrLf)
        ShowMessage(String.Format("{0}W E L C O M E{0}{0}T O{0}{0}T E T R I S{0}{0}{0}{0}R E V I S I O N{0}{0}{0}{0}Click here to start new game", vbCrLf))
        playername = InputBox("Please enter your name: ")
        scoreboard()
        'In the above code the falling block view is added in as well as the welcome to tetris revision message is in
    End Sub
    Sub scoreboard()
        ' This is the scoreboard code
        Dim len As Integer
        Dim pos As Integer
        Dim Textline As String
        Dim objrReader As New System.IO.StreamReader("n:\Scoreboard2.txt")
        Dim count = File.ReadAllLines("n:\Scoreboard2.txt").Length
        Leaderboard.Items.Clear()
        For i = 1 To count
            Textline = objrReader.ReadLine() & vbNewLine
            scores(i, 1) = Textline
        Next
        objrReader.Close()
        For i = 1 To count
            len = scores(i, 1).Length
            pos = scores(i, 1).IndexOf(",")
            If pos = -1 Then
                scores(i, 1) = "xxxx,000"
                scores(i, 2) = "xxxx"
                scores(i, 3) = "000"
            Else
                scores(i, 2) = Microsoft.VisualBasic.Left(scores(i, 1), pos)
                scores(i, 3) = Microsoft.VisualBasic.Right(scores(i, 1), len - (pos + 1))
            End If
            ' In the above code the code reads the scoreboard textfile
        Next
        Dim high As Integer = 0
        For i = 1 To count
            If scores(i, 3) > high Then
                high = scores(i, 3)
            End If
        Next
        For x = high To 0 Step -100
            For i = 1 To count
                If scores(i, 3) > 0 Then
                    If scores(i, 3) = x Then
                        Leaderboard.Items.Add(scores(i, 2) & "," & scores(i, 3))
                    End If
                End If
            Next
        Next
        Leaderboard.Enabled = False
        ' In the above code the users name and scores are added to the listbox to show the user the scoreboard
    End Sub
    Private Sub TetrisGame_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'In the below code the movement for the blocks is created with the pause button and a reset for the game
        Select Case e.KeyCode
            Case Keys.Left, Keys.Right, Keys.Down, Keys.Up
                If Status = GameStatus.Running Then
                    With FallingBlock
                        Select Case e.KeyCode
                            Case Keys.Left
                                If .CanMove(MoveDirection.Left) Then .Move(MoveDirection.Left)
                            Case Keys.Right
                                If .CanMove(MoveDirection.Right) Then .Move(MoveDirection.Right)
                            Case Keys.Down
                                If .CanMove(MoveDirection.Down) Then .Move(MoveDirection.Down)
                            Case Keys.Up
                                If .CanRotate Then .Rotate()
                        End Select
                    End With
                End If
            Case Keys.P
                If Status <> GameStatus.Stopped Then TogglePauseGame()
            Case Keys.N
                If Status = GameStatus.Stopped Then
                    StartNewGame()
                ElseIf DialogResult.Yes = MessageBox.Show("Restart Game?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) Then
                    StartNewGame()
                End If
        End Select
    End Sub
    Private Sub MessageLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MessageLabel.Click
        Revison_Timer.Enabled = True
        Select Case Status
            Case GameStatus.Stopped
                StartNewGame()
            Case GameStatus.Paused
                TogglePauseGame()
        End Select
        ' In the above code the revision timer is set when the game restarts and is stopped when the game is paused
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_1.Tick
        'In the below code the blocks fall and when the rows are completed it updates the gameboard so that the lines are completed and then disappear
        If FallingBlock.CanMove(MoveDirection.Down) Then
            FallingBlock.Move(MoveDirection.Down)
        Else
            For Each cell As TetrisCell In FallingBlock.Cells
                cell.IsEmpty = False
            Next
            Dim checkRows = From cell In FallingBlock.Cells _
                            Order By cell.Row _
                            Select cell.Row Distinct
            Dim rowsRemoved As Integer = 0
            For Each row In checkRows
                If GameBoard.IsRowComplete(row) Then
                    GameBoard.RemoveRow(row)
                    rowsRemoved += 1
                End If
                ' In the below code the maths is daaded in so that the score is updated and the speed can change depending on how well the player is doing
            Next
            Score += Math.Pow(rowsRemoved, 2) * 100
            Lines += rowsRemoved
            Speed = 1 + Lines \ 10
            If Speed Mod 10 = 0 Then Level += 1 : Speed = 1
            Timer_1.Interval = (10 - Speed) * 100
            UpdateStatistics()
            DropNextFallingBlock()
            If Not FallingBlock.CanMove(FallingBlock.CenterCell) Then EndGame()
        End If
    End Sub
#End Region
#Region "Private Methods"
    Private Function GetRandomShape() As TetrisBlock.Shapes
        Dim number As Integer = RandomNumbers.Next(Shapes.I1, Shapes.Z4 + 1)
        Return CType(number, TetrisBlock.Shapes)
    End Function
    Private Sub DropNextFallingBlock()
        ' In the below code the next block is falling as the preview block refeshes and the block is dropped
        FallingBlock.CenterCell = GameBoard.Cells(2, GameBoard.Columns \ 2)
        FallingBlock.Shape = PreviewBlock.Shape
        PreviewBlock.Shape = GetRandomShape()
        PreviewBlock.RefreshBackGround()
        PreviewBlock.Refresh()
    End Sub
    Private Sub StartNewGame()
        'In the below code the game has started again and the statistics have been re-set
        Score = 0
        Lines = 0
        Speed = 1
        Level = 1
        Timer_1.Interval = 1000
        UpdateStatistics()
        Time_Label.Text = 10
        GameBoard.SetupBoard()
        DropNextFallingBlock()
        MessageLabel.Visible = False
        Timer_1.Enabled = True
        Status = GameStatus.Running
    End Sub
    Private Sub EndGame()
        ' This is the end game code.
        Timer_1.Enabled = False
        Status = GameStatus.Stopped
        Dim len As Integer
        Dim pos As Integer
        For i As Integer = 1 To 5
            Dim textline As String = CStr(Leaderboard.Items(i - 1))
            len = textline.Length
            pos = textline.IndexOf(",")
            scores(i, 1) = textline
            scores(i, 2) = Microsoft.VisualBasic.Left(scores(i, 1), pos)
            scores(i, 3) = Microsoft.VisualBasic.Right(scores(i, 1), len - (pos + 2))
            'In the above code i have stopped the game and also have stoppped the timer from continuing. I have also declared integers named len, pos and i. I have used these integers to order the playername and the score for the player
        Next
        If Score > scores(1, 3) Then
            MsgBox("New High Score")
        ElseIf Score > scores(2, 3) Then
            MsgBox("2nd Place")
        ElseIf Score > scores(3, 3) Then
            MsgBox("3rd Place")
        ElseIf Score > scores(4, 3) Then
            MsgBox("4th Place")
        ElseIf Score > scores(5, 3) Then
            MsgBox("5th Place")
            ' In the above code i have added in each place and what message comes up for the player.
        End If
        File.Create("n:\scoreboard2.txt").Dispose()
        Dim newpos(5) As String
        For i = 0 To 4
            newpos(i) = Leaderboard.Items(i)
        Next
        newpos(5) = playername & "," & Score
        For i = 0 To 5
            Using writer As StreamWriter = New StreamWriter("n:\scoreboard2.txt", True)
                writer.Write(newpos(i) & Environment.NewLine)
            End Using
        Next
        Leaderboard.Items.Clear()
        scoreboard()
        ShowMessage(String.Format("{0}{0}GAME OVER{0}{0}{0}{0}Click here to start new game", vbCrLf))
        Revison_Timer.Stop()
        'Finally i have disposed of the old scoreboard and re wrote the new score board and cleared the listbox and stopped the timer continuing to show the game over message. 
    End Sub
    Private Sub UpdateStatistics()
        Label_score.Text = Score.ToString("000")
        LinesLabel.Text = Lines.ToString
        LevelLabel.Text = Level.ToString
        SpeedLabel.Text = Speed.ToString
        'In the above code the statistics have been updated so that when the programme starts the statistics are changed.
    End Sub
    Private Sub TogglePauseGame()
        'In the below code when the game is paused by the user the timer stops and the the game stops running.
        If Status = GameStatus.Paused Then
            Status = GameStatus.Running
            MessageLabel.Visible = False
            Timer_1.Enabled = True
            Revison_Timer.Start()
        Else
            Status = GameStatus.Paused
            ShowMessage(String.Format("{0}{0}GAME PAUSED{0}{0}{0}{0}Click here to resume.", vbCrLf))
            Revison_Timer.Stop()
        End If
    End Sub
    Private Sub ShowMessage(ByVal message As String)
        'In the below code the timer is stopped and a message is sent to the user
        MessageLabel.Text = message
        MessageLabel.Visible = True
        Timer_1.Enabled = False
    End Sub
#End Region
    Private Sub Revison_Timer_Tick(sender As Object, e As EventArgs) Handles Revison_Timer.Tick
        Dim temp As Integer = Time_Label.Text
        If temp > 0 Then
            Time_Label.Text = temp - 1
        End If
        If temp = 0 Then
            Revison_Timer.Stop()
            Dim message, title As String
            Dim answer As Integer = 0
            Dim num As Integer
            Dim num2 As Integer
            Dim correct As Integer = 1000
            Dim optype As Integer
            Randomize()
            num = Int(Rnd() * 12) + 1
            Randomize()
            num2 = Int(Rnd() * 12) + 1
            Randomize()
            optype = Int(Rnd() * 3) + 1
            If optype = 1 Then
                message = num & " + " & num2
                correct = num + num2
            ElseIf optype = 2 Then
                message = num & " - " & num2
                correct = num - num2
            Else
                message = num & " x " & num2
                correct = num * num2
            End If
            title = "Tetris Revision"
            Do While answer <> correct
                answer = InputBox(message, title)
            Loop
            If Status = GameStatus.Stopped Then
            Else
                Time_Label.Text = 10
                Revison_Timer.Start()
                'In the above code the revision timer is stopped and two random numbers are generated with three potential operations for the user to complete before they can continue with the game
            End If
        End If
    End Sub
End Class
