Public Class Form1
    Private xList As New List(Of GridItem)
    Private bSuppressEvents As Boolean = True

    Public Sub New()
        InitializeComponent()
        mapControls()
        initializeValues()
        bSuppressEvents = False
    End Sub
    Private Sub mapControls()
        For Each cnt In Me.Controls
            cnt.tabindex = 0
            mapControl(TryCast(cnt, Panel))
        Next
    End Sub

    Private Sub mapControl(cnt As Panel)
        If IsNothing(cnt) Then Exit Sub
        If Not IsNumeric(Strings.Right(cnt.Name, 2)) Then Exit Sub

        For Each x As TextBox In cnt.Controls

            xList.Add(New GridItem(x))
            AddHandler x.TextChanged, New EventHandler(AddressOf validateText)
        Next
    End Sub

    Private Function solve(index As Integer, Optional ByRef iCount As Integer = -1, Optional iMax As Integer = 2) As Boolean
        For x As Integer = index To xList.Count - 1
            If Not xList(x).isLocked Then
                For i = 1 To 9
                    If isValid(xList(x), i) Then
                        xList(x).currVal = i
                        If solve(x + 1, iCount, iMax) Then Return True
                        xList(x).currVal = 0
                    End If
                Next
                Return False
            End If
        Next

        If iCount >= 0 AndAlso iMax > 0 Then
            iCount += 1
            If iCount = iMax Then Return True
            Return False
        End If

        Return True
    End Function
    Private Function getSolutionCount() As Integer
        Dim iCount As Integer = 0
        If solve(0, iCount) Then
            Return iCount
        Else
            Return 0
        End If
    End Function

    Private Sub writeResult()
        Dim bX As Boolean = bSuppressEvents
        bSuppressEvents = True
        For Each x In xList
            x.writeResult()
        Next
        bSuppressEvents = bX
    End Sub

    Private Function reset(start As Integer) As Boolean
        Dim bX As Boolean = bSuppressEvents
        bSuppressEvents = True
        For x As Integer = start To xList.Count - 1
            xList(x).resetResult()
        Next
        bSuppressEvents = bX
    End Function
    Private Function isValidStartingPoint() As Boolean
        For Each x In xList
            If x.origVal <> 0 Then
                If Not isValid(x, x.origVal) Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Private Function isValid(xGrid As GridItem, iVal As Integer) As Boolean
        For Each x As GridItem In xList
            If isSibling(x, xGrid) Then
                If CInt(x.currVal) = iVal Then Return False
            End If
        Next

        Return True
    End Function

    Private Function getColor(x As TextBox) As Color
        If Not IsNumeric(x.Text) Then Return Color.Red
        If CInt(x.Text) < 1 OrElse CInt(x.Text) > 9 Then Return Color.Red

        For Each itm In xList
            If x.Equals(itm.xTxt) Then
                If isValid(itm, CInt(itm.xTxt.Text)) Then
                    Return Color.White
                Else
                    Return Color.Red
                End If
            End If
        Next
    End Function

    Private Function isSibling(x As GridItem, y As GridItem)
        If x.column = y.column AndAlso x.row = y.row Then Return False

        If x.column = y.column Then Return True
        If x.row = y.row Then Return True
        If x.panel = y.panel Then Return True

        Return False
    End Function


    Private Sub btnSolve_Click(sender As Object, e As EventArgs) Handles btnSolve.Click
        initializeValues()

        If Not isValidStartingPoint() Then
            MsgBox("ERROR - this grid is invalid")
            Exit Sub
        End If

        If solve(0) Then
            writeResult()
            MsgBox("DONE")
        Else
            MsgBox("ERROR - No valid solutions")
        End If

    End Sub
    Private Sub initializeValues()
        For i = 0 To xList.Count - 1
            With xList(i)
                .initialize()
            End With
        Next
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        For i = 0 To xList.Count - 1
            With xList(i)
                .resetResult()
            End With
        Next
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        clearTable()
    End Sub
    Private Sub clearTable()
        For i = 0 To xList.Count - 1
            With xList(i)
                .clearResult()
            End With
        Next
    End Sub

    Private Sub btnRandom_Click(sender As Object, e As EventArgs) Handles btnRandom.Click
        bSuppressEvents = True

        Dim rand As Random = New Random
        clearTable()

        For i = 0 To 5
            setValue(rand)
        Next

        If getSolutionCount() = 0 Then Exit Sub
        writeResult()

        initializeValues()

        Dim iCnt As Integer = 0
        For i = 0 To 60
            If removeValue(rand) Then iCnt += 1
            If iCnt = 2 Then Exit For
        Next

        initializeValues()
        bSuppressEvents = False

    End Sub
    Private Function removeValue(rand As Random) As Boolean
        Dim index As Integer = getIndex(rand, False)
        Dim val As String = xList(index).xTxt.Text
        xList(index).initialize("")

        If getSolutionCount() > 1 Then
            xList(index).initialize(val)
            Return True
        End If
    End Function
    Private Sub setValue(rand As Random)
        Dim index As Integer = getIndex(rand)
        xList(index).initialize(getValidValue(rand, index))
    End Sub
    Private Function getIndex(ByRef rand As Random, Optional bReturnEmpty As Boolean = True) As Integer
        Dim index As Integer
        Do
            index = rand.Next(0, 80)
            If bReturnEmpty Then
                If xList(index).xTxt.Text = "" Then Return index
            Else
                If xList(index).xTxt.Text <> "" Then Return index
            End If

        Loop
    End Function
    Private Function getValidValue(ByRef rand As Random, index As Integer) As Integer
        Dim val As Integer = rand.Next(1, 9)
        While Not (isValid(xList(index), val))
            val = rand.Next(1, 9)
        End While
        Return val
    End Function

    Friend Sub validateText(sender As TextBox, e As EventArgs)
        If bSuppressEvents Then Exit Sub

        If sender.Text = "" Then
            sender.BackColor = Color.White
        Else
            sender.BackColor = getColor(sender)
        End If
        If isComplete() Then
            If isCorrect() Then
                MsgBox("YAY!!! CONGRATS")
            Else

            End If
        End If
    End Sub
    Private Function isCorrect() As Boolean
        'TODO:  need to parse row, column, and block values
        Return True

    End Function
    Private Function isComplete() As Boolean
        For Each x In xList
            If x.xTxt.Text = "" Then Return False
            If x.xTxt.BackColor = Color.Red Then Return False
        Next
        Return True
    End Function
End Class
