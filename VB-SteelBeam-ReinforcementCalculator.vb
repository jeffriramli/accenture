# Part1

Imports System.IO

Public Class Form1
    'Declare global variables
    Dim b, d, l, w, K, Z As Double
    Dim A As Integer

    Private Sub btncalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncalculate.Click
        'CODE TO CALCULATE INPUT DATA
        b = Val(txtbreadth.Text)
        d = Val(txtdepth.Text)
        l = Val(txtlength.Text)
        w = Val(txtload.Text)

        'If statements for "calculating errors"
        If txtbreadth.Text = "" Or txtlength.Text = "" Or txtdepth.Text = "" Or txtload.Text = "" Then
            MsgBox("Please enter a value in the input boxes.", MsgBoxStyle.Exclamation, "Error")
        ElseIf b <= 0 Or d <= 0 Or l <= 0 Or w <= 0 Then
            MsgBox("Please enter a value greater than 0, negative values are not acceptable.", MsgBoxStyle.Exclamation, "Error")
        Else
            A = Steel(b, l, w, d)
            If A = -99 Then
                MsgBox("Beam cannot take the load.")
            End If
            txtarea.Text = A
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        'CODE TO SAVE THE RESULTS
        'Declare variables
        Dim sname As String
        Dim sr As StreamWriter

        'If statements for "saving errors"
        If txtbreadth.Text = "" Or txtlength.Text = "" Or txtdepth.Text = "" Or txtload.Text = "" Then
            MsgBox("Please enter and process data before saving.", MsgBoxStyle.Exclamation, "Error")
        Else
            SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt"
            If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                sname = SaveFileDialog1.FileName
                sr = New StreamWriter(sname)
                sr.WriteLine("STEEL REINFORCEMENT CALCULATION")
                sr.WriteLine("")
                sr.WriteLine(Date.Now)
                sr.WriteLine("")
                sr.WriteLine("Breadth, b = " & txtbreadth.Text & " mm")
                sr.WriteLine("Depth, d = " & txtdepth.Text & " mm")
                sr.WriteLine("Length, l = " & txtlength.Text & " mm")
                sr.WriteLine("Load, w = " & txtload.Text & " kN/m")
                sr.WriteLine("")
                sr.WriteLine("Steel Reinforcement Required = " & txtarea.Text & " mm^2")
                sr.WriteLine("")
                sr.WriteLine("Programmed by Jeffri Ramli")
                sr.Close()
            End If
        End If
    End Sub

    Private Sub btnclear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclear.Click
        'Clear all textboxes
        txtbreadth.Text = ""
        txtlength.Text = ""
        txtarea.Text = ""
        txtdepth.Text = ""
        txtload.Text = ""
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'Exit the program
        End
    End Sub
End Class

# Part2

Module AreaCalculation
    Public Function Steel(ByVal b As Double, ByVal l As Double, ByVal w As Double, ByVal d As Double) As Double
        'Declare variables
        Dim M, Mu, K, Z, fcu, fy As Double
        Dim A As Integer

        fcu = 30
        fy = 500

        'Calculate the data
        M = w * (l ^ 2) / 8
        Mu = (0.156 * fcu * b * (d ^ 2)) / 1000000

        If Mu < M Then
            Return -99
        Else
            K = (M * 1000000) / (fcu * b * (d ^ 2))
            Z = d * (0.5 + ((0.25 - (K / 0.9)) ^ 0.5))
            A = (M * 1000000) / (0.87 * fy * Z)
            Return A
        End If
    End Function
End Module
