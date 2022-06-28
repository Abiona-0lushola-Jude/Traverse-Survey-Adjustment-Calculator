Public Class Form2
    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim lat1, lat2, lat3, lat4 As Double
        Dim dep1, dep2, dep3, dep4 As Double
        Dim EL As Double
        Dim EDP As Double
        lat1 = (Val(txtl1.Text) * (Math.Cos(Val(txtbg1.Text))))
        lat2 = (Val(txtl2.Text) * (Math.Cos(Val(txtbg2.Text))))
        lat3 = (Val(txtl3.Text) * (Math.Cos(Val(txtbg3.Text))))
        lat4 = (Val(txtl4.Text) * (Math.Cos(Val(txtbg4.Text))))
        dep1 = (Val(txtl1.Text) * (Math.Sin(Val(txtbg1.Text))))
        dep2 = (Val(txtl2.Text) * (Math.Sin(Val(txtbg2.Text))))
        dep3 = (Val(txtl3.Text) * (Math.Sin(Val(txtbg3.Text))))
        dep4 = (Val(txtl4.Text) * (Math.Sin(Val(txtbg4.Text))))
        'THIS IS FOR THE AREA
        txtArea.Text = (Val(txtl1.Text) * Val(txtl2.Text) * Val(txtl3.Text) * Val(txtl4.Text))
        MsgBox("The Length = " & txtArea.Text)
        'THIS IS FOR THE LINEAR ACCURACY
        txtLA.Text = Math.Sqrt(Math.Pow(EL, 2) + Math.Pow(EDP, 2))
        MsgBox("Linear Misclosure =" & txtLA.Text)
        'GETTING THE NORTHING AND EASTING
        EL = lat1 + lat2 + lat3 + lat4
        EDP = dep1 + dep2 + dep3 + dep4
        If EL = 0 Then
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        txtArea.Text = ""
        txtbg1.Text = ""
        txtbg2.Text = ""
        txtbg3.Text = ""
        txtbg4.Text = ""
        txtcntE.Text = ""
        txtcntN.Text = ""
        txtl1.Text = ""
        txtl2.Text = ""
        txtl3.Text = ""
        txtl4.Text = ""
        txtLA.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Text = MsgBox("Do you want to exit?", vbYesNo)
        If Button1.Text = vbYes Then
            Application.Exit()
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
