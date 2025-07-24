Imports System.IO

Public Class Form1
    'Declare the global variables
    Dim area As ArrayList
    Dim sd, vol As Double

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        'CODE TO OPEN THE FILE
        'Declare variables
        Dim sr As StreamReader
        Dim fname, line As String

        area = New ArrayList

        'Open a Stream using the selection from a FileOpen Dialog
        Try
            If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                fname = OpenFileDialog1.FileName
                sr = New StreamReader(fname)

                'To separate the separation distance from the areas
                sd = sr.ReadLine
                txtdistance.Text = sd
                line = sr.ReadLine

                'Check for nothing from the stream
                While line <> Nothing
                    area.Add(Val(line))
                    txtarea.Text += line & vbNewLine
                    line = sr.ReadLine
                End While
                sr.Close()
            End If

        'Code for "opening errors"
        Catch fnf As FileNotFoundException
            MsgBox("File not found, try typing it correctly next time", MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub btncalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncalculate.Click
        'CODE TO CALCULATE INPUT DATA
        'Declare variables
        Dim i As Integer
        Dim sum As Double

        'Calculate the data
        Try
            If sd < 0 Then
                MsgBox("Separation distance must be bigger than 0", MsgBoxStyle.Exclamation, "Error")
            Else
                For i = 1 To area.Count - 2
                    sum = sum + area(i)
                Next

                'Formula to calculate the estimated volume
                vol = (sd * (area(0) + area(area.Count - 1) + 2 * sum)) / 2

                'Display the result in the textbox
                txtresult.Text = vol
            End If

        'Code for "calculating errors"
        Catch null As System.NullReferenceException
            MsgBox("No data found. Please import data.", MsgBoxStyle.Exclamation, "Error")
        Catch ex As System.ArgumentOutOfRangeException
            MsgBox("Invalid Data. Please check the file contents.", MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        'CODE TO SAVE THE RESULT
        'Declare variables
        Dim sname As String
        Dim sr As StreamWriter

        'Code for "saving errors"
        If txtdistance.Text = "" Or txtarea.Text = "" Or txtresult.Text = "" Then
            MsgBox("Please import a file and process the data before saving.", MsgBoxStyle.Exclamation, "Error")
        Else
            SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt"
            'Save the results into a text file
            Try
                If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                    sname = SaveFileDialog1.FileName
                    sr = New StreamWriter(sname)
                    sr.WriteLine("** JEFFRI'S ESTIMATED VOLUME CALCULATION **")
                    sr.WriteLine(Date.Now)
                    sr.WriteLine("Separation distance: " & sd.ToString & " metres")
                    sr.WriteLine("No. of cross sectional area are: " & area.Count.ToString)
                    sr.WriteLine("Volume is: " & vol & " cubic metres")
                    sr.Close()
                End If
            Catch ex As EndOfStreamException
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
            End Try
        End If
    End Sub

    Private Sub btnclear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclear.Click
        'Clear all textboxes
        txtdistance.Text = ""
        txtarea.Text = ""
        txtresult.Text = ""
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'Exit the program
        End
    End Sub
End Class
