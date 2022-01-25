Public Class GridItem
    Friend column As Integer
    Friend row As Integer
    Friend panel As Integer
    Friend xTxt As TextBox
    Friend origVal As Integer
    Friend currVal As Integer

    Public Sub New(x As TextBox)
        xTxt = x

        panel = Strings.Right(x.Parent.Name, 2)
        Dim iX As Integer = Strings.Right(x.Tag, 2)

        column = (iX Mod 10) + 3 * (panel Mod 10)
        row = CInt(iX / 10) + 3 * CInt(panel / 10)

        xTxt.TabIndex = row * 9 + column + 1
    End Sub
    Friend Function isLocked() As Boolean
        Return Not origVal = 0
    End Function
    Friend Sub resetResult()
        If origVal = 0 Then
            clearResult()
        Else
            currVal = origVal
            xTxt.Text = origVal
        End If
    End Sub
    Friend Sub writeResult()
        xTxt.Text = currVal
    End Sub
    Friend Sub clearResult()
        origVal = 0
        currVal = 0
        xTxt.Text = ""
    End Sub
    Friend Sub initialize()
        xTxt.Enabled = xTxt.Text = ""

        If xTxt.Text = "" Then

            origVal = 0
        Else
            origVal = CInt(xTxt.Text)
        End If

        currVal = origVal
    End Sub
    Friend Sub initialize(val As String)
        xTxt.Text = val
        initialize()
    End Sub
End Class
