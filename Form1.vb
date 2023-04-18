Imports System.Net
Imports System.IO
Imports System.Collections.ObjectModel
Imports System.Data
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word

Public Class Form1
    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    'Bottone Crea Report
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim S() As String = OpenFileDialog1.FileNames 'un array che contiene i nomi dei file scelti

        Dim File As String = OpenFileDialog1.FileName
        Dim objWrd As Microsoft.Office.Interop.Word.Application
        objWrd = New Microsoft.Office.Interop.Word.Application
        objWrd.Visible = True
        objWrd.DisplayAlerts = False
        If RadioButton4.Checked = True Then
            If MsgBox("Premi il pulsante Crea Report SEM-EDS!", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
                objWrd.Quit()
            End If
        End If

        If ListBox1.Items.Count = 0 Then
            If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
            End If
            objWrd.Quit()
            Exit Sub
        End If
        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False And RadioButton6.Checked = False And RadioButton7.Checked = False Then

            If MsgBox("Scegli prima il tipo di report!!", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
            End If
            objWrd.Quit()
            Exit Sub

            '***************************************************************************************************************
            'Report Tal quale
            '*************************************************************************************************************

        ElseIf RadioButton1.Checked = True Then
            Dim picture As String
            Dim percorso As String
            Dim immagine As String
            For Each f As String In ListBox1.Items

                percorso = Me.ListBox1.GetItemText(f)

                For Each picture In My.Computer.FileSystem.GetFiles(f)
                    If InStr(picture, ".jpg") > 0 Or InStr(immagine, ".JPG") > 0 Then

                        ListBox4.Items.Add(picture)

                    End If
                Next picture
                '**************************************************************************************************************************
                'Ordinamento listbox5
                '**************************************************************************************************************************
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X fr") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X r") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X fr") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X r") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X fr") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X fr") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X fr") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X r") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X r") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X r") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X fr") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X fr") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X fr") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X r") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X r") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X r") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X fr") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "10X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "10X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X r") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "10X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X fr") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "10X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "10X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X r") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "10X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "20X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X fr") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "20X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "20X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X r") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "20X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X fr") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "20X fr") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "20X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X r") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "20X r") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                With objWrd

                    .Documents.Open("C:\Format Tal quale")
                    .DisplayAlerts = False
                    Dim ncamp, camp As String
                    ncamp = InStrRev(f, "\")
                    camp = Mid(f, ncamp + 1)
                    .Selection.MoveRight(Count:=1, Extend:=1)
                    .Selection.TypeText("Tal quale: ")
                    .Selection.Font.Color = WdColor.wdColorRed
                    .Selection.TypeText(camp)
                    .Selection.Font.Color = WdColor.wdColorAutomatic

                    For a = 1 To 5 Step 1

                        .Selection.MoveDown()

                    Next

                    For Each immagine In ListBox5.Items

                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                        .Selection.MoveLeft(Count:=1, Extend:=1)
                        .Selection.InlineShapes(1).Height = 283.7
                        .Selection.InlineShapes(1).Width = 425.3
                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)


                        If InStr(immagine, "2,5X fr") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 25X, fronte.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "2,5X r") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 25X, retro.")
                            .Selection.MoveDown(Count:=6)
                        End If

                        If InStr(immagine, "2,5X fr uva") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 25X, filtro A, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "2,5X r uva") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 25X, filtro A, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                        If InStr(immagine, "2,5X fr uvd") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 25X, filtro D, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "2,5X r uvd") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 25X, filtro D, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                        If InStr(immagine, "5X fr") > 0 And InStr(immagine, ",") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 50X, fronte.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "5X r") > 0 And InStr(immagine, ",") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 50X, retro.")
                            .Selection.MoveDown(Count:=6)
                        End If

                        If InStr(immagine, "5X fr uva") > 0 And InStr(immagine, ",") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 50X, filtro A, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "5X r uva") > 0 And InStr(immagine, ",") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 50X, filtro A, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                        If InStr(immagine, "5X fr uvd") > 0 And InStr(immagine, ",") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 50X, filtro D, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "5X r uvd") > 0 And InStr(immagine, ",") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 50X, filtro D, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                        If InStr(immagine, "10X fr") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 100X, fronte.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "10X r") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 100X, retro.")
                            .Selection.MoveDown(Count:=6)
                        End If

                        If InStr(immagine, "10X fr uva") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 100X, filtro A, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "10X r uva") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 100X, filtro A, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                        If InStr(immagine, "10X fr uvd") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 100X, filtro D, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "10X r uvd") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 100X, filtro D, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                        If InStr(immagine, "20X fr") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 200X, fronte.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "20X r") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 200X, retro.")
                            .Selection.MoveDown(Count:=6)
                        End If

                        If InStr(immagine, "20X fr uva") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 200X, filtro A, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "20X r uva") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 200X, filtro A, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                        If InStr(immagine, "20X fr uvd") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 200X, filtro D, fronte.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "20X r uvd") > 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 200X, filtro D, retro.")
                            .Selection.MoveDown(Count:=5)

                        End If
                    Next
                    objWrd.ActiveDocument.SaveAs(FileName:=f + "\Tal quale " + camp)
                    objWrd.ActiveDocument.Close()
                    ListBox4.Items.Clear()
                    ListBox5.Items.Clear()

                End With
                ListBox5.Items.Clear()
            Next f





            '**************************************************************************************************************
            'Report Analisi mineralogica
            '******************************************************************************************************************

        ElseIf RadioButton2.Checked = True Then
            Dim a As Integer
            For Each f As String In ListBox1.Items
                Dim percorso As String
                Dim immagine As String
                Dim picture As String
                percorso = Me.ListBox1.GetItemText(f)

                For Each picture In My.Computer.FileSystem.GetFiles(f)
                    If InStr(picture, ".jpg") > 0 Or InStr(immagine, ".JPG") > 0 Then

                        ListBox4.Items.Add(picture)

                    End If
                Next picture

                '**************************************************************************************************************************
                'Ordinamento listbox5
                '**************************************************************************************************************************

                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "2,5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") = 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") = 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "5X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, ",") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "10X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "20X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "40X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_4") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") > 0 And InStr(picture, "NX") = 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                    If InStr(picture, "100X") > 0 And InStr(picture, "NP") = 0 And InStr(picture, "NX") > 0 And InStr(picture, "_5") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next






                objWrd.DisplayAlerts = False
                With objWrd
                    .Documents.Open("C:\Format analisi mineralogica")


                    Dim ncamp, camp As String
                    ncamp = InStrRev(f, "\")
                    camp = Mid(f, ncamp + 1)
                    .Selection.MoveRight(Count:=1, Extend:=1)
                    .Selection.TypeText("Analisi mineralogica: ")
                    .Selection.Font.Color = WdColor.wdColorRed
                    .Selection.TypeText(camp)
                    .Selection.Font.Color = WdColor.wdColorAutomatic

                    For a = 1 To 48 Step 1

                        .Selection.MoveDown()

                    Next

                    For Each immagine In ListBox5.Items

                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                        .Selection.MoveLeft(Count:=1, Extend:=1)
                        .Selection.InlineShapes(1).Height = 283.7
                        .Selection.InlineShapes(1).Width = 425.3
                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)


                        If InStr(immagine, "40X") > 0 Then
                            If InStr(immagine, "NX") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 400X, Nicol Incrociati.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "NP") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 400X, Nicol Paralleli.")
                                .Selection.MoveDown(Count:=6)
                            End If
                        End If
                        If InStr(immagine, "20X") > 0 Then
                            If InStr(immagine, "NX") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 200X, Nicol Incrociati.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "NP") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 200X, Nicol Paralleli.")
                                .Selection.MoveDown(Count:=6)
                            End If
                        End If
                        If InStr(immagine, "2,5X") > 0 Then
                            If InStr(immagine, "NX") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 25X, Nicol Incrociati.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "NP") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 25X, Nicol Paralleli.")
                                .Selection.MoveDown(Count:=6)
                            End If
                        End If

                        If InStr(immagine, "5X") > 0 And InStr(immagine, ",") = 0 Then
                            If InStr(immagine, "NX") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 50X, Nicol Incrociati.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "NP") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 50X, Nicol Paralleli.")
                                .Selection.MoveDown(Count:=6)
                            End If
                        End If
                        If InStr(immagine, "10X") > 0 Then
                            If InStr(immagine, "NX") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 100X, Nicol Incrociati.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "NP") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 100X, Nicol Paralleli.")
                                .Selection.MoveDown(Count:=6)
                            End If
                        End If
                        If InStr(immagine, "100X") > 0 Then
                            If InStr(immagine, "NX") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 1000X, Nicol Incrociati.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "NP") > 0 Then
                                .Selection.TypeText("Foto in luce trasmessa ingr.reali 1000X, Nicol Paralleli.")
                                .Selection.MoveDown(Count:=6)
                            End If
                        End If

                    Next
                    objWrd.ActiveDocument.SaveAs(FileName:=f + "\Analisi mineralogica " + camp)
                    objWrd.ActiveDocument.Close()
                    ListBox4.Items.Clear()
                    ListBox5.Items.Clear()


                End With
                ListBox5.Items.Clear()
            Next f



            '*********************************************************************************************************
            'Report analisi stratigrafica
            '*********************************************************************************************************

        ElseIf RadioButton3.Checked = True Then
            Dim picture As String
            For Each f As String In ListBox1.Items
                Dim percorso As String
                Dim immagine As String
                percorso = Me.ListBox1.GetItemText(f)

                For Each picture In My.Computer.FileSystem.GetFiles(f)
                    If InStr(picture, ".jpg") > 0 Or InStr(picture, ".JPG") > 0 Then

                        ListBox4.Items.Add(picture)

                    End If
                Next picture

                '**************************************************************************************************************************
                'Ordinamento listbox5
                '**************************************************************************************************************************

                For Each picture In ListBox4.Items
                    If InStr(picture, "STRATI") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "GEO") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "panoramica") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 And InStr(picture, "STRATI") = 0 And InStr(picture, "GEO") = 0 And InStr(picture, "panoramica") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X uva") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X uvd") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next



                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 2,5X") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next






                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 And InStr(picture, "STRATI") = 0 And InStr(picture, "GEO") = 0 And InStr(picture, "IRFC") = 0 And InStr(picture, "panoramica") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") > 0 And InStr(picture, "_") = 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uva") > 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 5X") > 0 And InStr(picture, ",") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 And InStr(picture, "STRATI") = 0 And InStr(picture, "GEO") = 0 And InStr(picture, "panoramica") = 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X uva") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X uvd") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 10X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

               
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 And InStr(picture, "IRFC") = 0 And InStr(picture, "STRATI") = 0 And InStr(picture, "GEO") = 0 And InStr(picture, "panoramica") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_") = 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "_2") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_2") > 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") = 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next
                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uvd") > 0 And InStr(picture, "_3") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next

                For Each picture In ListBox4.Items
                    If InStr(picture, "st 20X") > 0 And InStr(picture, "uva") = 0 And InStr(picture, "uvd") = 0 And InStr(picture, "_3") > 0 And InStr(picture, "IRFC") > 0 Then
                        ListBox5.Items.Add(picture)
                    End If
                Next










                







                With objWrd

                    .Documents.Open("C:\Format analisi stratigrafica")

                    Dim ncamp, camp As String
                    ncamp = InStrRev(f, "\")
                    camp = Mid(f, ncamp + 1)
                    .Selection.MoveRight(Count:=1, Extend:=1)
                    .Selection.TypeText("Analisi stratigrafica: ")
                    .Selection.Font.Color = WdColor.wdColorRed
                    .Selection.TypeText(camp)
                    .Selection.Font.Color = WdColor.wdColorAutomatic

                    For a = 1 To 5 Step 1

                        .Selection.MoveDown()

                    Next

                    For Each immagine In ListBox5.Items

                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                        .Selection.MoveLeft(Count:=1, Extend:=1)
                        .Selection.InlineShapes(1).Height = 283.7
                        .Selection.InlineShapes(1).Width = 425.3
                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)

                        If InStr(immagine, "st 2,5X") > 0 And InStr(immagine, "STRATI") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 25X con successioni stratigrafiche in evidenza.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "st 5X") > 0 And InStr(immagine, "STRATI") > 0 And InStr(immagine, ",") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 50X con successioni stratigrafiche in evidenza.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "st 10X") > 0 And InStr(immagine, "STRATI") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 100X con successioni stratigrafiche in evidenza.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "st 20X") > 0 And InStr(immagine, "STRATI") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 200X con successioni stratigrafiche in evidenza.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "GEO") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Rappresentazione grafica degli strati.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "st 2,5X") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") > 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 25X, panoramica.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "st 5X") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, ",") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") > 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 50X, panoramica.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "st 10X") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") > 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 100X, panoramica.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "st 20X") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "panoramica") > 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 200X, panoramica.")
                            .Selection.MoveDown(Count:=1)
                        End If
                        If InStr(immagine, "st 2,5X") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 25X.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "st 2,5X") > 0 And InStr(immagine, "uva") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 25X, filtro A.")
                            .Selection.MoveDown(Count:=1)

                        End If

                        If InStr(immagine, "st 2,5X") > 0 And InStr(immagine, "uvd") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 25X, filtro D.")
                            .Selection.MoveDown(Count:=1)

                        End If

                        If InStr(immagine, "st 5X") > 0 And InStr(immagine, ",") = 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "IRFC") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 50X.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "st 5X") > 0 And InStr(immagine, "uva") > 0 And InStr(immagine, ",") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 50X, filtro A.")
                            .Selection.MoveDown(Count:=1)

                        End If

                        If InStr(immagine, "st 5X") > 0 And InStr(immagine, "uvd") > 0 And InStr(immagine, ",") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 50X, filtro D.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "st 5X") > 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, ",") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "IRFC") > 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Ripresa in Infrarosso falsi colori, ingr. reali 50X.")
                            .Selection.MoveDown(Count:=1)

                        End If

                        If InStr(immagine, "st 10X") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "IRFC") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 100X.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "st 10X") > 0 And InStr(immagine, "uva") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 100X, filtro A.")
                            .Selection.MoveDown(Count:=1)

                        End If

                        If InStr(immagine, "st 10X") > 0 And InStr(immagine, "uvd") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 100X, filtro D.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "st 10X") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "IRFC") > 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Ripresa in Infrarosso falsi colori, ingr. reali 100X.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "st 20X") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 And InStr(immagine, "IRFC") = 0 Then
                            .Selection.TypeText("Foto in luce riflessa ingr.reali 200X.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "st 20X") > 0 And InStr(immagine, "uva") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 200X, filtro A.")
                            .Selection.MoveDown(Count:=1)

                        End If

                        If InStr(immagine, "st 20X") > 0 And InStr(immagine, "uvd") > 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Foto in fluorescenza ultravioletta indotta ingr. reali 200X, filtro D.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        If InStr(immagine, "st 20X") > 0 And InStr(immagine, "uva") = 0 And InStr(immagine, "uvd") = 0 And InStr(immagine, "STRATI") = 0 And InStr(immagine, "GEO") = 0 And InStr(immagine, "IRFC") > 0 And InStr(immagine, "panoramica") = 0 Then
                            .Selection.TypeText("Ripresa in Infrarosso falsi colori, ingr. reali 200X.")
                            .Selection.MoveDown(Count:=1)

                        End If

                    Next
                    objWrd.ActiveDocument.SaveAs(FileName:=f + "\Analisi stratigrafica " + camp)
                    objWrd.ActiveDocument.Close()
                    ListBox4.Items.Clear()
                    ListBox5.Items.Clear()

                End With
                ListBox5.Items.Clear()
            Next f
            '*************************************************************************************************************************
            'Report Punto di prelievo
            '************************************************************************************************************************
        ElseIf RadioButton5.Checked = True Then
            For Each f As String In ListBox1.Items
                Dim percorso As String
                Dim immagine As String
                percorso = Me.ListBox1.GetItemText(f)

                For Each immagine In My.Computer.FileSystem.GetFiles(f)
                    If InStr(immagine, ".jpg") > 0 Or InStr(immagine, ".JPG") > 0 Then

                        ListBox5.Items.Add(immagine)

                    End If
                Next immagine

                With objWrd

                    .Documents.Open("C:\Format punto di prelievo")

                    Dim ncamp, camp As String
                    ncamp = InStrRev(f, "\")
                    camp = Mid(f, ncamp + 1)
                    .Selection.MoveRight(Count:=1, Extend:=1)
                    .Selection.TypeText("Punto di prelievo")
                    .Selection.MoveDown(Count:=1)



                    For Each immagine In ListBox5.Items

                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                        .Selection.MoveLeft(Count:=1, Extend:=1)
                        .Selection.InlineShapes(1).Height = 283.7
                        .Selection.InlineShapes(1).Width = 425.3
                        .Selection.MoveDown(Count:=1)
                        .Selection.MoveRight(Count:=1, Extend:=1)


                        If InStr(immagine, "gen") > 0 Then
                            .Selection.TypeText("Foto del punto di prelievo del campione " + camp + ", generale.")
                            .Selection.MoveDown(Count:=1)
                        End If

                        If InStr(immagine, "macro") > 0 Then
                            .Selection.TypeText("Foto del punto di prelievo del campione " + camp + ", macro.")
                            .Selection.MoveDown(Count:=1)

                        End If
                        .Selection.MoveDown(Count:=3)
                    Next
                    objWrd.ActiveDocument.SaveAs(FileName:=f + "\Punto di prelievo " + camp)
                    objWrd.ActiveDocument.Close()
                    ListBox4.Items.Clear()
                    ListBox5.Items.Clear()

                End With
                ListBox5.Items.Clear()
            Next f
            '********************************************************************************************************************
            'Report XRF (Artax)
            '*******************************************************************************************************************
        ElseIf RadioButton6.Checked = True And RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False Then
            For Each f As String In ListBox1.Items
                Dim objXls As Microsoft.Office.Interop.Excel.Application
                objXls = New Microsoft.Office.Interop.Excel.Application
                objXls.Visible = True
                Dim percorso As String
                Dim immagine As String
                Dim testo As String
                Dim path2 As String
                Dim g As String
                percorso = Me.ListBox1.GetItemText(f)

                For Each immagine In My.Computer.FileSystem.GetFiles(f)
                    If InStr(immagine, ".jpg") > 0 Or InStr(immagine, ".JPG") > 0 Then

                        ListBox5.Items.Add(immagine)

                    End If
                Next immagine
                Dim npunt0 As String
                Dim descrizione0 As String
                Dim npunt As String
                Dim descrizione As String
                Dim descrizione2 As String
                Dim estensione As String
                For Each testo In My.Computer.FileSystem.GetFiles(f)
                    If InStr(testo, ".txt") > 0 Then

                        ListBox4.Items.Add(testo)

                    End If
                Next
                For Each g In ListBox4.Items
                    path2 = Me.ListBox4.GetItemText(g)

                Next g
                npunt0 = InStrRev(g, "\")
                descrizione0 = InStrRev(g, "-")
                estensione = InStrRev(g, ".")
                descrizione = Mid(g, descrizione0 + 1, estensione - descrizione0 - 1)
                npunt = Mid(f, InStrRev(f, "\") + 1)



                With objWrd

                    .Documents.Open("C:\Format XRF")
                    .Selection.MoveDown(Count:=3)
                    .Selection.Font.Bold = True
                    .Selection.Font.Name = "Book Antiqua"
                    .Selection.Font.Size = 12
                    .Selection.TypeText(npunt + ". " + descrizione)
                    .Selection.Font.Bold = False
                    .Selection.MoveDown(Count:=2)

                    For Each immagine In ListBox5.Items


                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                        .Selection.MoveLeft(Count:=1, Extend:=1)
                        .Selection.InlineShapes(1).Height = 158.8
                        .Selection.InlineShapes(1).Width = 237.2
                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.TypeText("  ")
                        .Selection.MoveRight(Count:=1, Extend:=1)

                    Next




                End With
                '********************
                'pre-elaborazione

                For Each testo In ListBox4.Items

                    With objXls
                        objXls.Visible = True
                        .Workbooks.Open(testo)

                        .Cells.EntireColumn("A").Delete()
                        .Range("E1").Select()
                        .ActiveCell.Value = "Conc"
                        .Cells.EntireColumn("F:Z").Delete()
                        .Range("A1").Select()
                        .ActiveCell.Value = "Elt."
                        .Range("A2").Select()
                        Do Until .ActiveCell.Value = ""
                            If .ActiveCell.Value <> .ActiveCell.Offset(1, 0).Value Then
                                .ActiveCell.Offset(1, 0).Select()
                            End If
                            If .ActiveCell.Value = .ActiveCell.Offset(1, 0).Value Then
                                If .ActiveCell.Value = "Al" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Si" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "P" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If

                                If .ActiveCell.Value = "S" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Cl" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "K" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Ca" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Ti" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Cr" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Mn" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Fe" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Co" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Ni" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Cu" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Zn" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "As" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Se" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Br" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Sr" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Ag" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If

                                If .ActiveCell.Value = "Cd" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If

                                If .ActiveCell.Value = "Sn" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Sb" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "K12" Then
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Ba" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "L1" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Pt" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "L1" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Au" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "L1" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Hg" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "L1" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Pb" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "L1" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = "Bi" Then
                                    .ActiveCell.Offset(0, 1).Select()
                                    If .ActiveCell.Value = "L1" Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -1).Select()
                                    End If
                                End If
                                If .ActiveCell.Value = .ActiveCell.Offset(1, 0).Value Then
                                    .ActiveCell.Offset(0, 4).Select()
                                    If .ActiveCell.Value > .ActiveCell.Offset(1, 0).Value Then
                                        .ActiveCell.Offset(1, 0).Select()
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -4).Select()
                                    End If
                                    If .ActiveCell.Value < .ActiveCell.Offset(1, 0).Value Then
                                        .ActiveCell.EntireRow.Delete()
                                        .ActiveCell.Offset(0, -4).Select()
                                    End If
                                End If


                            End If

                        Loop


                        .Range("A2").Select()
                        Do Until .ActiveCell.Value = ""
                            If .ActiveCell.Value <> "Mo" Then
                                .ActiveCell.Offset(1, 0).Select()
                            End If
                            If .ActiveCell.Value = "Mo" Then
                                .ActiveCell.EntireRow.Delete()
                            End If

                        Loop


                        .Range("A2").Select()
                        Do Until .ActiveCell.Value = ""
                            If .ActiveCell.Value <> "Nb" Then
                                .ActiveCell.Offset(1, 0).Select()
                            End If
                            If .ActiveCell.Value = "Nb" Then
                                .ActiveCell.EntireRow.Delete()
                            End If

                        Loop
                        .ActiveWorkbook.Close(SaveChanges:=True)
                        .Application.Quit()

                    End With


                Next
                '*********************************
                'Elaborazione

                objXls.Visible = True
                If ListBox4.Items.Count = 0 Then
                    If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                        Exit Sub
                    End If
                    objXls.Quit()
                    Exit Sub
                End If
                objXls.Workbooks.Open("C:\Elaborato Excel.xls")


                For Each testo In ListBox4.Items

                    Dim Riga As String, M As New Collections.ArrayList
                    Dim R As Long, C As Long, NomeColonna As String, NumColonna As Long, Inizio As Long, Misura As Long = 2

                    NomeColonna = "Conc"  '<<----Specificare qua il nome della colonna che contiene i valori da estrarre
                    NumColonna = -1
                    '*****************************************************************
                    'lettura file e salva i valori  in arraylist (M)
                    FileOpen(1, testo, OpenMode.Input)
                    Do While Not EOF(1)
                        Riga = LineInput(1)
                        If Len(Trim(Riga)) > 0 Then
                            R = M.Add(Split(Riga, vbTab))
                            For C = 0 To UBound(M(R))
                                If M(R)(C) = NomeColonna Then NumColonna = C : Inizio = R + 1
                                If Trim(LCase(M(R)(C))) = "total" Then Exit Do
                            Next C
                        End If
                    Loop
                    FileClose(1)

                    If NumColonna < 0 Then 'caso la colonna interessata non sia presente nel file, quindi termina..
                        MsgBox("Colonna  '" + NomeColonna + "'  non è stata trovata nel file report:" + _
                            vbCr + "'" + testo + "'." + vbCr + vbCr + "Controllare che il nome della colonna da analizzare sia corretto", vbCritical, "ATTENZIONE!")

                    End If
                    '*********************************************************************
                    'Inserimento valori nel foglio della tabella preformattata in excel

                    With objXls



                        If .Selection.Rows.Count > 0 Then
                            Misura = .Selection.Row ' se vi è almeno una riga selezionata , allora utilizza la riga selezionata come posizione d'inserimento. 
                            .Rows(Misura).Select() ' riseleziona l'intera riga di riferimento 
                            .ActiveCell.Offset(1, 0).Select()
                            .ActiveCell.EntireRow.Select()
                        End If


                        Misura = .Selection.Row ' prende il riferimento riga dalla riga selezionata 


                        For R = Inizio To M.Count - 1
                            If Len(Trim(M(R)(0))) > 0 Then
                                For C = 2 To .Columns.Count
                                    If Len(Trim(.Cells(1, C).value.ToString)) = 0 Then 'controllo opzionale:se l'elemento non esiste nella tabella chiede di aggiungerlo automaticamente
                                        If MsgBox("Elemento:   ' " & M(R)(0) & " '   non esiste nella tabella-foglio di excel..." & vbCr & vbCr & _
                                                "Si vuole aggiungere adesso questo nuovo elemento alla tabella?", vbYesNo, "File corrente: '" & testo & "'") = vbYes Then
                                            .Cells(1, C) = M(R)(0)
                                            .Cells(Misura, C) = M(R)(NumColonna)
                                        End If
                                        Exit For
                                    ElseIf Trim(.Cells(1, C).Value.ToString) = Trim(M(R)(0)) Then

                                        .Cells(Misura, C) = M(R)(NumColonna)

                                        Exit For

                                    End If

                                Next C

                            End If

                        Next R


                        NumColonna = -1 : Inizio = 0 : M.Clear() ' queste variabili vengono azzerate per ogni nuovo ciclo
                    End With
                Next testo
                '******************************************************************************************************************
                '-------------------------------------------------------------------------------------------------------------------
                '*******************************************************************************************************************


                Dim h As Integer
                Dim p As Integer


                With objXls

                    'Eliminazione colonne vuote 
                    .Cells.Range("B2:B401").Select()
                    For h = 1 To 93 Step 1
                        If .WorksheetFunction.CountBlank(.Selection) = 400 Then
                            .ActiveCell.EntireColumn.Delete()
                        End If
                        If .WorksheetFunction.CountBlank(.Selection) < 400 Then
                            .Selection.Offset(0, 1).Select()
                        End If
                    Next h

                    'Eliminazione righe vuote
                    .Cells.Range("B2:CO2").Select()
                    For p = 1 To 40 Step 1
                        If .WorksheetFunction.CountBlank(.Selection) = 92 Then
                            .ActiveCell.EntireRow.Delete()
                        End If
                        If .WorksheetFunction.CountBlank(.Selection) < 92 Then
                            .Selection.Offset(1, 0).Select()
                        End If
                    Next p

                    .Cells.Range("A1").Select()

                    'Selezione tabella

                    .ActiveSheet.UsedRange.Select()



                    'Conversione celle da testo a numero
                    For Each xcell In .Selection
                        If IsNumeric(xcell.Value) Then
                            xcell.Value = xcell.Value * 1
                        End If
                    Next xcell

                    'Formattazione tabella
                    .Cells.Range("A1").Select()
                    .ActiveCell.Value = "Elementi"
                    .Cells.Range("A2").Select()
                    .ActiveCell.Value = "Cnts"
                    .ActiveSheet.UsedRange.Select()
                    .Cells.Font.Name = "Book Antiqua"
                    .Cells.Font.Size = "10"
                    .Cells.Range("A1:B2").Select()
                    .Selection.Font.Bold = True
                    .ActiveSheet.UsedRange.Select()

                    'Bordi tabella
                    .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlInsideVertical).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlInsideVertical).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlInsideHorizontal).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlInsideHorizontal).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeTop).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeRight).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeRight).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeLeft).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeLeft).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeBottom).ColorIndex = XlColorIndex.xlColorIndexAutomatic






                    Dim directory As String
                    directory = My.Computer.FileSystem.GetParentPath(testo)
                    .ActiveWorkbook.SaveAs(Filename:=directory + "\Tabella XRF")
                    .Selection.Copy()

                End With
                With objWrd
                    .Selection.MoveRight(Count:=1, Extend:=1)
                    .Selection.PasteExcelTable(False, False, True)

                    .Selection.MoveDown(Count:=10)
                    For nt As Integer = 1 To .ActiveDocument.Tables.Count
                        .ActiveDocument.Tables(nt).AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent)
                        .ActiveDocument.Tables(nt).AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
                    Next nt

                End With

                objXls.ActiveWorkbook.Close(SaveChanges:=vbYes)
                objXls.Quit()
                ListBox4.Items.Clear()
                ListBox5.Items.Clear()

            Next f
            'Salva
            With objWrd
                If SaveFileDialog6.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                    .ActiveDocument.SaveAs(SaveFileDialog6.FileName)
                    .Application.DisplayAlerts = True
                Else
                    .ActiveDocument.Close(SaveChanges:=False)
                    .Application.Quit()

                    Exit Sub
                End If
            End With
            ListBox4.Items.Clear()
            ListBox5.Items.Clear()
            objWrd.Application.Quit()

            '*********************************************************************************************************************
            'Report XRF (Portatile)
            '********************************************************************************************************************
        ElseIf RadioButton7.Checked = True And RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False And RadioButton6.Checked = False Then
            For Each f As String In ListBox1.Items
                Dim objXls As Microsoft.Office.Interop.Excel.Application
                objXls = New Microsoft.Office.Interop.Excel.Application
                objXls.Visible = True
                Dim percorso As String
                Dim immagine As String
                Dim testo As String
                Dim path2 As String
                Dim g As String
                percorso = Me.ListBox1.GetItemText(f)

                For Each immagine In My.Computer.FileSystem.GetFiles(f)
                    If InStr(immagine, ".jpg") > 0 Or InStr(immagine, ".JPG") > 0 Then

                        ListBox5.Items.Add(immagine)

                    End If
                Next immagine
                Dim npunt0 As String
                Dim descrizione0 As String
                Dim npunt As String
                Dim descrizione As String
                Dim descrizione2 As String
                Dim estensione As String
                For Each testo In My.Computer.FileSystem.GetFiles(f)
                    If InStr(testo, ".txt") > 0 Then

                        ListBox4.Items.Add(testo)

                    End If
                Next
                For Each g In ListBox4.Items
                    path2 = Me.ListBox4.GetItemText(g)

                Next g
                npunt0 = InStrRev(g, "\")
                descrizione0 = InStrRev(g, "-")
                estensione = InStrRev(g, ".")
                descrizione = Mid(g, descrizione0 + 1, estensione - descrizione0 - 1)
                npunt = Mid(f, InStrRev(f, "\") + 1)



                With objWrd

                    .Documents.Open("C:\Format XRF")
                    .Selection.MoveDown(Count:=3)
                    .Selection.Font.Bold = True
                    .Selection.Font.Name = "Book Antiqua"
                    .Selection.Font.Size = 12
                    .Selection.TypeText(npunt + ". " + descrizione)
                    .Selection.Font.Bold = False
                    .Selection.MoveDown(Count:=2)
                    For Each immagine In ListBox5.Items


                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                        .Selection.MoveLeft(Count:=1, Extend:=1)
                        .Selection.InlineShapes(1).Height = 158.8
                        .Selection.InlineShapes(1).Width = 237.2
                        .Selection.MoveRight(Count:=1, Extend:=1)
                        .Selection.TypeText("  ")
                        .Selection.MoveRight(Count:=1, Extend:=1)

                    Next




                End With
                'pre-elaborazione
                For Each testo In ListBox4.Items

                    With objXls
                        objXls.Visible = True
                        .Workbooks.OpenText(testo, DataType:=XlTextParsingType.xlDelimited, Space:=True, ConsecutiveDelimiter:=True)
                        ProgressBar2.Maximum = 10
                        ProgressBar2.Value = 3
                        Label2.Text = "Inizio formattazione files di testo ..."
                        For ab As Integer = 1 To 24
                            .ActiveCell.Rows.EntireRow.Delete()
                        Next
                        .ActiveCell.Offset(0, 2).Select()
                        .ActiveCell.Columns.EntireColumn.Delete()
                        .Cells.Range("B1").Select()
                        Do Until .ActiveCell.Value = ""
                            If Len(.ActiveCell.Value) = 1 Then
                                .ActiveCell.Offset(0, 1).Value = .ActiveCell.Offset(0, 2).Value
                            End If
                            .ActiveCell.Offset(1, 0).Select()
                        Loop
                        .Cells.Range("B1").Select()
                        .ActiveCell.Offset(0, 2).Select()
                        .ActiveCell.Columns.EntireColumn.Delete()
                        .ActiveCell.Columns.EntireColumn.Delete()
                        .ActiveCell.Columns.EntireColumn.Delete()
                        .ActiveCell.Columns.EntireColumn.Delete()
                        .ActiveCell.Offset(0, -2).Select()
                        Do Until .ActiveCell.Value = ""
                            If Len(.ActiveCell.Value) = 4 Then
                                .ActiveCell.Offset(0, -1).Value = Microsoft.VisualBasic.Left(.ActiveCell.Value, 2)

                            ElseIf Len(.ActiveCell.Value) = 3 Then
                                .ActiveCell.Offset(0, -1).Value = Microsoft.VisualBasic.Left(.ActiveCell.Value, 1)

                            ElseIf Len(.ActiveCell.Value) = 1 Then
                                .ActiveCell.Offset(0, -1).Value = Microsoft.VisualBasic.Left(.ActiveCell.Value, 1)
                            End If
                            .ActiveCell.Offset(1, 0).Select()
                        Loop
                        .ActiveCell.Columns.EntireColumn.Delete()
                        .Cells.Range("A1").Select()
                        .ActiveCell.Rows.EntireRow.Insert()
                        .Cells.Range("A1").Value = "Elt."
                        .Cells.Range("B1").Value = "Conc"
                        ProgressBar2.Value = 7
                        Label2.Text = "Formattazione in corso ..."
                        If CheckBox1.Checked Then
                            .Range("A2").Select()
                            Do Until .ActiveCell.Value = ""
                                If .ActiveCell.Value <> "Mo" Then
                                    .ActiveCell.Offset(1, 0).Select()
                                End If
                                If .ActiveCell.Value = "Mo" Then
                                    .ActiveCell.EntireRow.Delete()
                                End If

                            Loop
                        End If
                        If CheckBox2.Checked Then
                            .Range("A2").Select()
                            Do Until .ActiveCell.Value = ""
                                If .ActiveCell.Value <> "Nb" Then
                                    .ActiveCell.Offset(1, 0).Select()
                                End If
                                If .ActiveCell.Value = "Nb" Then
                                    .ActiveCell.EntireRow.Delete()
                                End If

                            Loop
                        End If


                        If CheckBox3.Checked Then
                            .Range("A2").Select()
                            Do Until .ActiveCell.Value = ""
                                If .ActiveCell.Value <> "W" Then
                                    .ActiveCell.Offset(1, 0).Select()
                                End If
                                If .ActiveCell.Value = "W" Then
                                    .ActiveCell.EntireRow.Delete()
                                End If

                            Loop
                        End If

                        If CheckBox4.Checked Then
                            .Range("A2").Select()
                            Do Until .ActiveCell.Value = ""
                                If .ActiveCell.Value <> TextBox2.Text Then
                                    .ActiveCell.Offset(1, 0).Select()
                                End If
                                If .ActiveCell.Value = TextBox2.Text Then
                                    .ActiveCell.EntireRow.Delete()
                                End If

                            Loop
                        End If
                        If CheckBox5.Checked Then
                            .Range("A2").Select()
                            Do Until .ActiveCell.Value = ""
                                If .ActiveCell.Value <> TextBox3.Text Then
                                    .ActiveCell.Offset(1, 0).Select()
                                End If
                                If .ActiveCell.Value = TextBox3.Text Then
                                    .ActiveCell.EntireRow.Delete()
                                End If

                            Loop
                        End If
                        If CheckBox6.Checked Then
                            .Range("A2").Select()
                            Do Until .ActiveCell.Value = ""
                                If .ActiveCell.Value <> TextBox4.Text Then
                                    .ActiveCell.Offset(1, 0).Select()
                                End If
                                If .ActiveCell.Value = TextBox4.Text Then
                                    .ActiveCell.EntireRow.Delete()
                                End If

                            Loop
                        End If

                        .ActiveWorkbook.Close(SaveChanges:=True)
                        .Application.Quit()



                    End With
                Next testo


                ProgressBar2.Value = 10
                Label2.Text = "Formattazione completata!!"
                '*********************************
                'Elaborazione

                objXls.Visible = True
                If ListBox4.Items.Count = 0 Then
                    If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                        Exit Sub
                    End If
                    objXls.Quit()
                    Exit Sub
                End If
                objXls.Workbooks.Open("C:\Elaborato Excel.xls")


                For Each testo In ListBox4.Items

                    Dim Riga As String, M As New Collections.ArrayList
                    Dim R As Long, C As Long, NomeColonna As String, NumColonna As Long, Inizio As Long, Misura As Long = 2

                    NomeColonna = "Conc"  '<<----Specificare qua il nome della colonna che contiene i valori da estrarre
                    NumColonna = -1
                    '*****************************************************************
                    'lettura file e salva i valori  in arraylist (M)
                    FileOpen(1, testo, OpenMode.Input)
                    Do While Not EOF(1)
                        Riga = LineInput(1)
                        If Len(Trim(Riga)) > 0 Then
                            R = M.Add(Split(Riga, vbTab))
                            For C = 0 To UBound(M(R))
                                If M(R)(C) = NomeColonna Then NumColonna = C : Inizio = R + 1
                                If Trim(LCase(M(R)(C))) = "total" Then Exit Do
                            Next C
                        End If
                    Loop
                    FileClose(1)

                    If NumColonna < 0 Then 'caso la colonna interessata non sia presente nel file, quindi termina..
                        MsgBox("Colonna  '" + NomeColonna + "'  non è stata trovata nel file report:" + _
                            vbCr + "'" + testo + "'." + vbCr + vbCr + "Controllare che il nome della colonna da analizzare sia corretto", vbCritical, "ATTENZIONE!")

                    End If
                    '*********************************************************************
                    'Inserimento valori nel foglio della tabella preformattata in excel

                    With objXls



                        If .Selection.Rows.Count > 0 Then
                            Misura = .Selection.Row ' se vi è almeno una riga selezionata , allora utilizza la riga selezionata come posizione d'inserimento. 
                            .Rows(Misura).Select() ' riseleziona l'intera riga di riferimento 
                            .ActiveCell.Offset(1, 0).Select()
                            .ActiveCell.EntireRow.Select()
                        End If


                        Misura = .Selection.Row ' prende il riferimento riga dalla riga selezionata 


                        For R = Inizio To M.Count - 1
                            If Len(Trim(M(R)(0))) > 0 Then
                                For C = 2 To .Columns.Count
                                    If Len(Trim(.Cells(1, C).value.ToString)) = 0 Then 'controllo opzionale:se l'elemento non esiste nella tabella chiede di aggiungerlo automaticamente
                                        If MsgBox("Elemento:   ' " & M(R)(0) & " '   non esiste nella tabella-foglio di excel..." & vbCr & vbCr & _
                                                "Si vuole aggiungere adesso questo nuovo elemento alla tabella?", vbYesNo, "File corrente: '" & testo & "'") = vbYes Then
                                            .Cells(1, C) = M(R)(0)
                                            .Cells(Misura, C) = M(R)(NumColonna)
                                        End If
                                        Exit For
                                    ElseIf Trim(.Cells(1, C).Value.ToString) = Trim(M(R)(0)) Then

                                        .Cells(Misura, C) = M(R)(NumColonna)

                                        Exit For

                                    End If

                                Next C

                            End If

                        Next R


                        NumColonna = -1 : Inizio = 0 : M.Clear() ' queste variabili vengono azzerate per ogni nuovo ciclo
                    End With
                Next testo
                '******************************************************************************************************************
                '-------------------------------------------------------------------------------------------------------------------
                '*******************************************************************************************************************


                Dim h As Integer
                Dim p As Integer


                With objXls

                    'Eliminazione colonne vuote 
                    .Cells.Range("B2:B401").Select()
                    For h = 1 To 93 Step 1
                        If .WorksheetFunction.CountBlank(.Selection) = 400 Then
                            .ActiveCell.EntireColumn.Delete()
                        End If
                        If .WorksheetFunction.CountBlank(.Selection) < 400 Then
                            .Selection.Offset(0, 1).Select()
                        End If
                    Next h

                    'Eliminazione righe vuote
                    .Cells.Range("B2:CO2").Select()
                    For p = 1 To 40 Step 1
                        If .WorksheetFunction.CountBlank(.Selection) = 92 Then
                            .ActiveCell.EntireRow.Delete()
                        End If
                        If .WorksheetFunction.CountBlank(.Selection) < 92 Then
                            .Selection.Offset(1, 0).Select()
                        End If
                    Next p

                    .Cells.Range("A1").Select()

                    'Selezione tabella

                    .ActiveSheet.UsedRange.Select()



                    'Conversione celle da testo a numero
                    For Each xcell In .Selection
                        If IsNumeric(xcell.Value) Then
                            xcell.Value = xcell.Value * 1
                        End If
                    Next xcell

                    'Formattazione tabella
                    .Cells.Range("A1").Select()
                    .ActiveCell.Value = "Elementi"
                    .Cells.Range("A2").Select()
                    .ActiveCell.Value = "Cnts"
                    .ActiveSheet.UsedRange.Select()
                    .Cells.Font.Name = "Book Antiqua"
                    .Cells.Font.Size = "10"
                    .Cells.Range("A1:B2").Select()
                    .Selection.Font.Bold = True
                    .ActiveSheet.UsedRange.Select()

                    'Bordi tabella
                    .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlInsideVertical).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlInsideVertical).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlInsideHorizontal).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlInsideHorizontal).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeTop).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeRight).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeRight).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeLeft).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeLeft).ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    .Selection.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    .Selection.Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
                    .Selection.Borders(XlBordersIndex.xlEdgeBottom).ColorIndex = XlColorIndex.xlColorIndexAutomatic






                    Dim directory As String
                    directory = My.Computer.FileSystem.GetParentPath(testo)
                    .ActiveWorkbook.SaveAs(Filename:=directory + "\Tabella XRF")
                    .Selection.Copy()

                End With
                With objWrd
                    .Selection.MoveRight(Count:=1, Extend:=1)
                    .Selection.PasteExcelTable(False, False, True)

                    .Selection.MoveDown(Count:=10)
                    For nt As Integer = 1 To .ActiveDocument.Tables.Count
                        .ActiveDocument.Tables(nt).AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent)
                        .ActiveDocument.Tables(nt).AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
                    Next nt

                End With
                objXls.ActiveWorkbook.Close(SaveChanges:=vbYes)
                objXls.Quit()
                ListBox4.Items.Clear()
                ListBox5.Items.Clear()

            Next f
            'Salva
            With objWrd
                If SaveFileDialog6.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                    .ActiveDocument.SaveAs(SaveFileDialog6.FileName)
                    .Application.DisplayAlerts = True
                Else
                    .ActiveDocument.Close(SaveChanges:=False)
                    .Application.Quit()

                    Exit Sub
                End If
            End With
        End If
        

        ListBox4.Items.Clear()
        ListBox5.Items.Clear()

        objWrd.Application.Quit()


    End Sub

    'Bottone Cancella

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ListBox1.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
    End Sub
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Tab 2 XRF
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '******************************************************************************************************************
        'Finestra di dialogo Apri Files e aggiungi files in Listbox2
        '******************************************************************************************************************

        ProgressBar1.Value = 0
        Label1.Text = ""
        Label2.Text = ""
        If OpenFileDialog2.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim S() As String = OpenFileDialog2.FileNames 'un array che contiene i nomi dei file scelti
            Dim File As String

            For Each File In S
                ListBox2.Items.Add(File)
            Next
        End If
    End Sub

    '***********************************************************************************************************************
    'Tasto Cancella
    '************************************************************************************************************************
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ProgressBar1.Value = 0
        Label1.Text = ""
        ProgressBar2.Value = 0
        Label2.Text = ""
        ListBox2.Items.Clear()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim File As String = OpenFileDialog1.FileName
        Dim objXls As Microsoft.Office.Interop.Excel.Application
        objXls = New Microsoft.Office.Interop.Excel.Application
        objXls.Visible = True
        If ListBox2.Items.Count = 0 Then
            If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
            End If
            objXls.Quit()
            Exit Sub
        End If


        For Each File In ListBox2.Items

            With objXls
                objXls.Visible = True
                .Workbooks.Open(File)
                ProgressBar2.Maximum = 10
                ProgressBar2.Value = 3
                Label2.Text = "Inizio formattazione files di testo ..."



                .Cells.EntireColumn("A").Delete()
                .Range("E1").Select()
                .ActiveCell.Value = "Conc"
                .Cells.EntireColumn("F:Z").Delete()
                .Range("A1").Select()
                .ActiveCell.Value = "Elt."
                .Range("A2").Select()
                Do Until .ActiveCell.Value = ""
                    If .ActiveCell.Value <> .ActiveCell.Offset(1, 0).Value Then
                        .ActiveCell.Offset(1, 0).Select()
                    End If
                    If .ActiveCell.Value = .ActiveCell.Offset(1, 0).Value Then
                        If .ActiveCell.Value = "Al" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Si" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "P" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If

                        If .ActiveCell.Value = "S" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Cl" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "K" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Ca" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Ti" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Cr" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Mn" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Fe" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Co" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Ni" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Cu" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Zn" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "As" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Se" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Br" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Sr" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Ag" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If

                        If .ActiveCell.Value = "Cd" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If

                        If .ActiveCell.Value = "Sn" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Sb" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "K12" Then
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Ba" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "L1" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Pt" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "L1" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Au" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "L1" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Hg" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "L1" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Pb" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "L1" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = "Bi" Then
                            .ActiveCell.Offset(0, 1).Select()
                            If .ActiveCell.Value = "L1" Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -1).Select()
                            End If
                        End If
                        If .ActiveCell.Value = .ActiveCell.Offset(1, 0).Value Then
                            .ActiveCell.Offset(0, 4).Select()
                            If .ActiveCell.Value > .ActiveCell.Offset(1, 0).Value Then
                                .ActiveCell.Offset(1, 0).Select()
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -4).Select()
                            End If
                            If .ActiveCell.Value < .ActiveCell.Offset(1, 0).Value Then
                                .ActiveCell.EntireRow.Delete()
                                .ActiveCell.Offset(0, -4).Select()
                            End If
                        End If


                    End If
                    ProgressBar2.Value = 7
                    Label2.Text = "Formattazione in corso ..."
                Loop

                If CheckBox1.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> "Mo" Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = "Mo" Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If
                If CheckBox2.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> "Nb" Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = "Nb" Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If


                If CheckBox3.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> "W" Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = "W" Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If

                If CheckBox4.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> TextBox2.Text Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = TextBox2.Text Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If
                If CheckBox5.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> TextBox3.Text Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = TextBox3.Text Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If
                If CheckBox6.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> TextBox4.Text Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = TextBox4.Text Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If

                .ActiveWorkbook.Close(SaveChanges:=True)
                .Application.Quit()
            End With

        Next File

        ProgressBar2.Value = 10
        Label2.Text = "Formattazione completata!!"
    End Sub



    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        '****************************************************************************************************************
        'Tasto Elabora
        '****************************************************************************************************************
        ProgressBar1.Maximum = 15
        ProgressBar1.Value = 1
        Label1.Text = "Caricamento files..."

        Dim File As String = OpenFileDialog1.FileName
        Dim objXls As Microsoft.Office.Interop.Excel.Application
        objXls = New Microsoft.Office.Interop.Excel.Application
        objXls.Visible = False
        If ListBox2.Items.Count = 0 Then
            If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
            End If
            objXls.Quit()
            Exit Sub
        End If
        objXls.Workbooks.Open("C:\Elaborato Excel per Artax.xls")

        ProgressBar1.Value = 3
        Label1.Text = "Inizio importazione dati..."
        For Each File In ListBox2.Items

            Dim Riga As String, M As New Collections.ArrayList
            Dim R As Long, C As Long, NomeColonna As String, NumColonna As Long, Inizio As Long, Misura As Long = 2

            NomeColonna = "Conc"  '<<----Specificare qua il nome della colonna che contiene i valori da estrarre
            NumColonna = -1
            '*****************************************************************
            'lettura file e salva i valori  in arraylist (M)
            FileOpen(1, File, OpenMode.Input)
            Do While Not EOF(1)
                Riga = LineInput(1)
                If Len(Trim(Riga)) > 0 Then
                    R = M.Add(Split(Riga, vbTab))
                    For C = 0 To UBound(M(R))
                        If M(R)(C) = NomeColonna Then NumColonna = C : Inizio = R + 1
                        If Trim(LCase(M(R)(C))) = "total" Then Exit Do
                    Next C
                End If
            Loop
            FileClose(1)

            If NumColonna < 0 Then 'caso la colonna interessata non sia presente nel file, quindi termina..
                MsgBox("Colonna  '" + NomeColonna + "'  non è stata trovata nel file report:" + _
                    vbCr + "'" + File + "'." + vbCr + vbCr + "Controllare che il nome della colonna da analizzare sia corretto", vbCritical, "ATTENZIONE!")

            End If
            '*********************************************************************
            'Inserimento valori nel foglio della tabella preformattata in excel

            With objXls

                ProgressBar1.Value = 6
                Label1.Text = "Inizio scrittura dati..."

                If .Selection.Rows.Count > 0 Then
                    Misura = .Selection.Row ' se vi è almeno una riga selezionata , allora utilizza la riga selezionata come posizione d'inserimento. 
                    .Rows(Misura).Select() ' riseleziona l'intera riga di riferimento 
                    .ActiveCell.Offset(1, 0).Select()
                    .ActiveCell.EntireRow.Select()
                End If


                Misura = .Selection.Row ' prende il riferimento riga dalla riga selezionata 


                For R = Inizio To M.Count - 1
                    If Len(Trim(M(R)(0))) > 0 Then
                        For C = 2 To .Columns.Count
                            If Len(Trim(.Cells(1, C).value.ToString)) = 0 Then 'controllo opzionale:se l'elemento non esiste nella tabella chiede di aggiungerlo automaticamente
                                If MsgBox("Elemento:   ' " & M(R)(0) & " '   non esiste nella tabella-foglio di excel..." & vbCr & vbCr & _
                                        "Si vuole aggiungere adesso questo nuovo elemento alla tabella?", vbYesNo, "File corrente: '" & File & "'") = vbYes Then
                                    .Cells(1, C) = M(R)(0)
                                    .Cells(Misura, C) = M(R)(NumColonna)
                                End If
                                Exit For
                            ElseIf Trim(.Cells(1, C).Value.ToString) = Trim(M(R)(0)) Then

                                .Cells(Misura, C) = M(R)(NumColonna)

                                Exit For

                            End If

                        Next C

                    End If

                Next R


                NumColonna = -1 : Inizio = 0 : M.Clear() ' queste variabili vengono azzerate per ogni nuovo ciclo
            End With
        Next File
        '******************************************************************************************************************
        '-------------------------------------------------------------------------------------------------------------------
        '*******************************************************************************************************************


        Dim h As Integer
        Dim f As Integer

        ProgressBar1.Value = 9
        Label1.Text = "Formattazione tabella: Eliminazione righe e colonne vuote..."
        With objXls

            'Eliminazione colonne vuote 
            .Cells.Range("B2:B401").Select()
            For h = 1 To 93 Step 1
                If .WorksheetFunction.CountBlank(.Selection) = 400 Then
                    .ActiveCell.EntireColumn.Delete()
                End If
                If .WorksheetFunction.CountBlank(.Selection) < 400 Then
                    .Selection.Offset(0, 1).Select()
                End If
            Next h

            'Eliminazione righe vuote
            .Cells.Range("B2:CO2").Select()
            For f = 1 To 400 Step 1
                If .WorksheetFunction.CountBlank(.Selection) = 92 Then
                    .ActiveCell.EntireRow.Delete()
                End If
                If .WorksheetFunction.CountBlank(.Selection) < 92 Then
                    .Selection.Offset(1, 0).Select()
                End If
            Next f

            .Cells.Range("A1").Select()

            'Selezione tabella

            .ActiveSheet.UsedRange.Select()

            ProgressBar1.Value = 10
            Label1.Text = "Conversione valori...."

            'Conversione celle da testo a numero
            For Each xcell In .Selection
                If IsNumeric(xcell.Value) Then
                    xcell.Value = xcell.Value * 1
                End If
            Next xcell

            'Selezione tabella
            .ActiveSheet.UsedRange.Select()
            ProgressBar1.Value = 13
            Label1.Text = "Creazione bordi..."

            'Bordi tabella
            .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlInsideVertical).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlInsideVertical).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeTop).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeRight).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeRight).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeBottom).ColorIndex = XlColorIndex.xlColorIndexAutomatic



            .Cells.Font.Name = "Book Antiqua"
            .Cells.Font.Size = "10"




            ProgressBar1.Value = 13
            Label1.Text = "Salvataggio tabella..."


            '**************************************************************************
            '**************************************************************************
            If SaveFileDialog4.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                .Application.DisplayAlerts = False
                .ActiveWorkbook.SaveAs(SaveFileDialog4.FileName)
                .Application.DisplayAlerts = True
            Else
                .ActiveWorkbook.Close(SaveChanges:=False)
                .Application.Quit()
                ProgressBar1.Value = 15
                Label1.Text = "Creazione tabella abortita!"
                Exit Sub
            End If

            '*******************************************************************************
            '*******************************************************************************


            ProgressBar1.Value = 15
            Label1.Text = "Tabella creata con successo!"


            .ActiveWorkbook.Close()
            .Application.Quit()
            Exit Sub
        End With
    End Sub
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Tab3 SEM-EDS
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '******************************************************************************************************************
        'Finestra di dialogo Apri Files e aggiungi files in Listbox1
        '******************************************************************************************************************

        ProgressBar3.Value = 0
        Label3.Text = ""

        If OpenFileDialog3.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim S() As String = OpenFileDialog3.FileNames 'un array che contiene i nomi dei file scelti
            Dim File As String

            For Each File In S
                ListBox3.Items.Add(File)
            Next



        End If

    End Sub


    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        '***********************************************************************************************************************
        'Tasto Cancella
        '************************************************************************************************************************

        ProgressBar3.Value = 0
        Label3.Text = ""
        ListBox3.Items.Clear()

    End Sub


    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        ProgressBar3.Maximum = 15
        ProgressBar3.Value = 1
        Label3.Text = "Caricamento files..."

        Dim File As String = OpenFileDialog1.FileName
        Dim objXls As Microsoft.Office.Interop.Excel.Application
        objXls = New Microsoft.Office.Interop.Excel.Application
        objXls.Visible = True
        If ListBox3.Items.Count = 0 Then
            If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
            End If
            objXls.Quit()
            Exit Sub
        End If
        objXls.Workbooks.Open("C:\Elaborato Excel.xls")

        ProgressBar3.Value = 3
        Label3.Text = "Inizio importazione dati..."
        For Each File In ListBox3.Items

            Dim Riga As String, M As New Collections.ArrayList
            Dim R As Long, C As Long, NomeColonna As String, NumColonna As Long, Inizio As Long, Misura As Long = 2

            NomeColonna = "Conc"  '<<----Specificare qua il nome della colonna che contiene i valori da estrarre
            NumColonna = -1
            '*****************************************************************
            'lettura file e salva i valori  in arraylist (M)
            FileOpen(1, File, OpenMode.Input)
            Do While Not EOF(1)
                Riga = LineInput(1)
                If Len(Trim(Riga)) > 0 Then
                    R = M.Add(Split(Riga, vbTab))
                    For C = 0 To UBound(M(R))
                        If M(R)(C) = NomeColonna Then NumColonna = C : Inizio = R + 1
                        If Trim(LCase(M(R)(C))) = "total" Then Exit Do
                    Next C
                End If
            Loop
            FileClose(1)

            If NumColonna < 0 Then 'caso la colonna interessata non sia presente nel file, quindi termina..
                MsgBox("Colonna  '" + NomeColonna + "'  non è stata trovata nel file report:" + _
                    vbCr + "'" + File + "'." + vbCr + vbCr + "Controllare che il nome della colonna da analizzare sia corretto", vbCritical, "ATTENZIONE!")

            End If
            '*********************************************************************
            'Inserimento valori nel foglio della tabella preformattata in excel

            With objXls

                ProgressBar3.Value = 6
                Label3.Text = "Inizio scrittura dati..."

                If .Selection.Rows.Count > 0 Then
                    Misura = .Selection.Row ' se vi è almeno una riga selezionata , allora utilizza la riga selezionata come posizione d'inserimento. 
                    .Rows(Misura).Select() ' riseleziona l'intera riga di riferimento 
                    .ActiveCell.Offset(1, 0).Select()
                    .ActiveCell.EntireRow.Select()
                End If


                Misura = .Selection.Row ' prende il riferimento riga dalla riga selezionata 


                For R = Inizio To M.Count - 1
                    If Len(Trim(M(R)(0))) > 0 Then
                        For C = 2 To .Columns.Count
                            If Len(Trim(.Cells(1, C).value.ToString)) = 0 Then 'controllo opzionale:se l'elemento non esiste nella tabella chiede di aggiungerlo automaticamente
                                If MsgBox("Elemento:   ' " & M(R)(0) & " '   non esiste nella tabella-foglio di excel..." & vbCr & vbCr & _
                                        "Si vuole aggiungere adesso questo nuovo elemento alla tabella?", vbYesNo, "File corrente: '" & File & "'") = vbYes Then
                                    .Cells(1, C) = M(R)(0)
                                    .Cells(Misura, C) = M(R)(NumColonna)
                                End If
                                Exit For
                            ElseIf Trim(.Cells(1, C).Value.ToString) = Trim(M(R)(0)) Then

                                .Cells(Misura, C) = M(R)(NumColonna)

                                Exit For

                            End If

                        Next C

                    End If

                Next R


                NumColonna = -1 : Inizio = 0 : M.Clear() ' queste variabili vengono azzerate per ogni nuovo ciclo
            End With
        Next File
        '******************************************************************************************************************
        '-------------------------------------------------------------------------------------------------------------------
        '*******************************************************************************************************************


        Dim h As Integer
        Dim f As Integer

        ProgressBar3.Value = 9
        Label3.Text = "Formattazione tabella: Eliminazione righe e colonne vuote..."
        With objXls

            'Eliminazione colonne vuote 
            .Cells.Range("B2:B41").Select()
            For h = 1 To 93 Step 1
                If .WorksheetFunction.CountBlank(.Selection) = 40 Then
                    .ActiveCell.EntireColumn.Delete()
                End If
                If .WorksheetFunction.CountBlank(.Selection) < 40 Then
                    .Selection.Offset(0, 1).Select()
                End If
            Next h

            'Eliminazione righe vuote
            .Cells.Range("B2:CO2").Select()
            For f = 1 To 40 Step 1
                If .WorksheetFunction.CountBlank(.Selection) = 92 Then
                    .ActiveCell.EntireRow.Delete()
                End If
                If .WorksheetFunction.CountBlank(.Selection) < 92 Then
                    .Selection.Offset(1, 0).Select()
                End If
            Next f

            .Cells.Range("A1").Select()

            'Selezione tabella

            .ActiveSheet.UsedRange.Select()

            ProgressBar3.Value = 10
            Label3.Text = "Conversione valori...."

            'Conversione celle da testo a numero
            For Each xcell In .Selection
                If IsNumeric(xcell.Value) Then
                    xcell.Value = xcell.Value * 1
                End If
            Next xcell

            ProgressBar3.Value = 11
            Label3.Text = "Formattazione valori..."

            'Ripristino virgole e formattazione numero
            For Each xcell In .Selection
                If IsNumeric(xcell.Value) Then
                    If xcell.Value >= 1000 Then
                        xcell.Value = xcell.Value / 1000
                    End If
                End If
                .Selection.NumberFormat = "0.00"
            Next xcell

            ProgressBar3.Value = 12
            Label3.Text = "Verifica quantitative..."

            'Verifica Quantitative
            .Cells.Range("CP2").Select()
            .ActiveCell.Formula = "=Sum(B2:CO2)"
            .Cells.Range("CP2:CP41").Select()
            .ActiveCell.AutoFill(Destination:=.Cells.Range("CP2:CP41"))

            For Each xcell In .Selection
                If xcell.Value = 0 Then
                    xcell.Clear()
                End If
            Next
            .Cells.Range("CP2").Select()
            Do Until .ActiveCell.Value = 0
                If .ActiveCell.Value < 99.98 Then
                    .ActiveCell.Offset(0, -93).Select()
                    If MsgBox("Le concentrazioni del File .txt non sono corrette in '" + .ActiveCell.Value + "'; controllare i report di analisi. L'applicazione verrà chiusa.", 1 + 16, "Errore nei file di origine") = vbOK Then
                        .ActiveWorkbook.Close(SaveChanges:=False)
                        .Application.Quit()
                        Exit Sub
                    End If
                    If MsgBox("Le concentrazioni del File .txt non sono corrette in '" + .ActiveCell.Value + "'; controllare i report di analisi. L'applicazione verrà chiusa.", 1 + 16, "Errore nei file di origine") = vbCancel Then

                    End If
                    .ActiveCell.Offset(0, 93).Select()
                    .ActiveCell.Offset(1, 0).Select()
                End If
                If .ActiveCell.Value >= 99.98 Then
                    .ActiveCell.Offset(1, 0).Select()
                End If
            Loop
            .Cells.Range("CP2:CP42").Select()
            .Selection.Clear()

            'Selezione tabella
            .ActiveSheet.UsedRange.Select()
            ProgressBar3.Value = 13
            Label3.Text = "Creazione bordi..."

            'Bordi tabella


            .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlInsideVertical).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlInsideVertical).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeTop).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeRight).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeRight).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).ColorIndex = XlColorIndex.xlColorIndexAutomatic
            .Selection.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            .Selection.Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
            .Selection.Borders(XlBordersIndex.xlEdgeBottom).ColorIndex = XlColorIndex.xlColorIndexAutomatic

            .Selection.Font.Name = "Book Antiqua"
            .Selection.Font.Size = "10"




            ProgressBar3.Value = 13
            Label3.Text = "Salvataggio tabella..."

            Dim d As String
            Dim path As String
            Dim subpath As String
            subpath = Me.ListBox3.GetItemText(File)
            path = My.Computer.FileSystem.GetParentPath(File)



            '**************************************************************************
            'Salvataggio tabella
            '**************************************************************************

            .ActiveWorkbook.SaveAs(Filename:=path + "\Tabella excel")
            ProgressBar3.Value = 15
            Label3.Text = "Tabella creata con successo!"


            .ActiveWorkbook.Close(SaveChanges:=vbYes)
            .Application.Quit()


        End With


    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        Dim File As String = OpenFileDialog1.FileName
        Dim objXls As Microsoft.Office.Interop.Excel.Application
        objXls = New Microsoft.Office.Interop.Excel.Application
        objXls.Visible = True
        If ListBox3.Items.Count = 0 Then
            If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
            End If
            objXls.Quit()
            Exit Sub
        End If
        Dim f As Integer

        For Each File In ListBox3.Items

            With objXls
                objXls.Visible = True
                .Workbooks.OpenText(File, DataType:=XlTextParsingType.xlDelimited, ConsecutiveDelimiter:=True, Tab:=False, Space:=True)

                .Cells.EntireRow("1").Delete()
                .Cells.EntireRow("1").Delete()
                For f = 1 To 50
                    .ActiveCell.Offset(1, 0).Select()
                    If InStr(.ActiveCell.Value, "---") Then
                        .ActiveCell.EntireRow.Delete()
                    End If
                Next
                For f = 1 To 3
                    .Cells.EntireColumn("B").Delete()
                Next f
                For f = 1 To 7
                    .Cells.EntireColumn("C").Delete()
                Next
                .Cells.EntireRow("2").Delete()
                .Cells.Range("A1").Select()
                .ActiveCell.Value = ("Elt.")
                .ActiveCell.Offset(0, 1).Select()
                .ActiveCell.Value = ("Conc")
                .ActiveWorkbook.Close(SaveChanges:=True)
                .Application.Quit()



            End With
        Next File

    End Sub
    '******************************************************************************************************************
    'Bottone Carica Cartelle
    '*******************************************************************************************************************

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim Cartella As String
        Dim subfolder As String
        Dim Folderbrowserdialog1 As New FolderBrowserDialog
        If Folderbrowserdialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Cartella = Folderbrowserdialog1.SelectedPath()
            For Each subfolder In My.Computer.FileSystem.GetDirectories(Cartella)
                ListBox1.Items.Add(subfolder)
            Next


        End If

    End Sub

    '****************************************************************************************************************************************
    'Report analisi SEM-EDS
    '****************************************************************************************************************************************



    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click


        Dim objXls As Microsoft.Office.Interop.Excel.Application
        objXls = New Microsoft.Office.Interop.Excel.Application
        objXls.Visible = False
        Dim objWrd As New Microsoft.Office.Interop.Word.Application
        objWrd.Visible = False
        objWrd.DisplayAlerts = False
        Dim Cartella As String
        Dim g As String
        Dim f As String
        Dim h As String
        Dim percorso As String
        Dim percorso2 As String
        Dim ingrsem, prsem, dpsem, ncamp, camp, ingrsem2, ingrsem3 As String

        If RadioButton4.Checked = True And RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then


            For Each f In ListBox1.Items
                percorso = Me.ListBox1.GetItemText(f)

                Dim tabella As String
                Dim immagine As String
                Dim folder As String




                For Each immagine In My.Computer.FileSystem.GetFiles(f)
                    If InStr(immagine, ".jpg") > 0 Then
                        If InStr(immagine, "bei") Or InStr(immagine, "sei") Or InStr(immagine, "con frame di analisi") > 0 Then
                            ListBox5.Items.Add(immagine)
                        End If
                    End If
                Next immagine
                With objWrd

                    .Documents.Open("C:\Format Analisi SEM-EDS")
                    objWrd.Visible = True
                    .DisplayAlerts = False


                    ncamp = InStrRev(f, "\")
                    camp = Mid(f, ncamp + 1)

                    .Selection.MoveRight(Count:=1, Extend:=1)
                    .Selection.TypeText("Analisi SEM-EDS: ")
                    .Selection.Font.Color = WdColor.wdColorRed
                    .Selection.TypeText(camp)
                    .Selection.Font.Color = WdColor.wdColorAutomatic
                    .Selection.MoveDown(Count:=5)
                    For Each immagine In ListBox5.Items
                        If InStr(immagine, "con frame di analisi") > 0 Then
                            .Selection.MoveUp(Count:=4, Extend:=0)
                            .Selection.MoveRight(Count:=1, Extend:=1)
                            .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                            .Selection.MoveLeft(Count:=1, Extend:=1)
                            .Selection.InlineShapes(1).Height = 283.7
                            .Selection.InlineShapes(1).Width = 425.3
                            .Selection.MoveDown(Count:=1)
                            .Selection.MoveRight(Count:=1, Extend:=1)
                            If InStr(immagine, "st 2,5X") > 0 Then
                                .Selection.TypeText("Foto in luce riflessa ingr.reali 25X con frame di analisi SEM-EDS.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "st 5X") > 0 And InStr(immagine, ",") = 0 Then
                                .Selection.TypeText("Foto in luce riflessa ingr.reali 50X con frame di analisi SEM-EDS.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "st 10X") > 0 Then
                                .Selection.TypeText("Foto in luce riflessa ingr.reali 100X con frame di analisi SEM-EDS.")
                                .Selection.MoveDown(Count:=1)
                            End If
                            If InStr(immagine, "st 20X") > 0 Then
                                .Selection.TypeText("Foto in luce riflessa ingr.reali 200X con frame di analisi SEM-EDS.")
                                .Selection.MoveDown(Count:=1)
                            End If
                        End If
                        If InStr(immagine, "bei") > 0 Then
                            .Selection.MoveRight(Count:=1, Extend:=1)
                            .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                            .Selection.MoveLeft(Count:=1, Extend:=1)
                            .Selection.MoveDown(Count:=1)
                            prsem = InStrRev(immagine, "_")
                            dpsem = InStrRev(immagine, " ")
                            ingrsem = Mid(immagine, prsem + 1, dpsem - prsem - 1)
                            .Selection.TypeText("Immagine SEM del campione " + camp + " " + ingrsem + ", backscattering")

                        End If
                        If InStr(immagine, "sei") > 0 Then
                            .Selection.MoveRight(Count:=1, Extend:=1)
                            .Selection.InlineShapes.AddPicture(immagine, LinkToFile:=False, SaveWithDocument:=True)
                            .Selection.MoveLeft(Count:=1, Extend:=1)
                            .Selection.MoveDown(Count:=1)
                            prsem = InStrRev(immagine, "_")
                            dpsem = InStrRev(immagine, "X")
                            ingrsem = Mid(immagine, prsem - dpsem)
                            .Selection.TypeText("Immagine SEM del campione " + camp + " " + ingrsem + ", elettroni secondari")
                        End If

                    Next immagine

                    .Selection.MoveDown(Count:=2)
                End With


                ListBox5.Items.Clear()



                'aggiunge sottocartelle in listbox4
                For Each Cartella In My.Computer.FileSystem.GetDirectories(f)

                    ListBox4.Items.Add(Cartella)

                    For Each g In ListBox4.Items
                        percorso2 = Me.ListBox1.GetItemText(g)

                    Next g
                    For Each folder In My.Computer.FileSystem.GetFiles(g)
                        If InStr(folder, ".jpg") > 0 Then
                            If InStr(folder, "bei") Or InStr(folder, "sei") > 0 Then
                                ListBox5.Items.Add(folder)
                            End If
                        End If
                        If InStr(folder, "con pa") = 0 Then
                            ListBox5.Items.Remove(folder)
                        End If
                    Next folder
                    ingrsem2 = InStrRev(g, "\")
                    ingrsem3 = Mid(g, ingrsem2 + 1)

                    For Each frame In ListBox5.Items
                        With objWrd

                            If InStr(frame, "bei con pa") > 0 Then
                                .Selection.MoveRight(Count:=1, Extend:=1)
                                .Selection.InlineShapes.AddPicture(frame, LinkToFile:=False, SaveWithDocument:=True)
                                .Selection.MoveLeft(Count:=1, Extend:=1)
                                .Selection.MoveDown(Count:=1)
                                prsem = InStrRev(frame, "_")
                                dpsem = InStrRev(frame, " ")
                                ingrsem = Mid(frame, prsem + 1, dpsem - prsem - 1)
                                .Selection.TypeText("Frame .Immagine SEM del campione " + camp + " " + ingrsem3 + "X" + ", backscattering")

                            End If
                            If InStr(frame, "sei con pa") > 0 Then
                                .Selection.MoveRight(Count:=1, Extend:=1)
                                .Selection.InlineShapes.AddPicture(frame, LinkToFile:=False, SaveWithDocument:=True)
                                .Selection.MoveLeft(Count:=1, Extend:=1)
                                .Selection.MoveDown(Count:=1)
                                prsem = InStrRev(frame, "_")
                                dpsem = InStrRev(frame, "X")
                                ingrsem = Mid(frame, prsem - dpsem)
                                .Selection.TypeText("Frame .Immagine SEM del campione " + camp + " " + ingrsem3 + "X" + ", elettroni secondari")
                            End If
                        End With


                        With objWrd
                            .Selection.MoveDown(Count:=5)
                            tabella = Path.Combine(g, "Tabella excel.xls")
                        End With
                        With objXls
                            .DisplayAlerts = False
                            .Visible = True
                            .Workbooks.Open(tabella)
                            .ActiveSheet.UsedRange.Select()
                            .Selection.Copy()

                        End With
                        With objWrd
                            .Selection.MoveRight(Count:=1, Extend:=1)
                            .Selection.PasteExcelTable(False, False, True)

                            .Selection.MoveDown(Count:=10)
                            For nt As Integer = 1 To .ActiveDocument.Tables.Count
                                .ActiveDocument.Tables(nt).AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
                            Next nt

                        End With

                        objXls.Quit()


                    Next frame
                    ListBox4.Items.Remove(Cartella)
                    ListBox5.Items.Clear()

                Next Cartella

                objWrd.ActiveDocument.SaveAs(FileName:=f + "\Analisi SEM-EDS " + camp)
                objWrd.ActiveDocument.Close()
                ListBox4.Items.Clear()
                ListBox5.Items.Clear()



            Next f

        End If
        objWrd.Quit()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim File As String = OpenFileDialog1.FileName
        Dim objXls As Microsoft.Office.Interop.Excel.Application
        objXls = New Microsoft.Office.Interop.Excel.Application
        objXls.Visible = True
        If ListBox2.Items.Count = 0 Then
            If MsgBox("Prima di procedere all'elaborazione, selezionare i files", 1 + 16, "Errore!") = vbYes Then
                Exit Sub
            End If
            objXls.Quit()
            Exit Sub
        End If


        For Each File In ListBox2.Items

            With objXls
                objXls.Visible = True
                .Workbooks.OpenText(File, DataType:=XlTextParsingType.xlDelimited, Space:=True, ConsecutiveDelimiter:=True)
                ProgressBar2.Maximum = 10
                ProgressBar2.Value = 3
                Label2.Text = "Inizio formattazione files di testo ..."
                For ab As Integer = 1 To 24
                    .ActiveCell.Rows.EntireRow.Delete()
                Next
                .ActiveCell.Offset(0, 2).Select()
                .ActiveCell.Columns.EntireColumn.Delete()
                .Cells.Range("B1").Select()
                Do Until .ActiveCell.Value = ""
                    If Len(.ActiveCell.Value) = 1 Then
                        .ActiveCell.Offset(0, 1).Value = .ActiveCell.Offset(0, 2).Value
                    End If
                    .ActiveCell.Offset(1, 0).Select()
                Loop
                .Cells.Range("B1").Select()
                .ActiveCell.Offset(0, 2).Select()
                .ActiveCell.Columns.EntireColumn.Delete()
                .ActiveCell.Columns.EntireColumn.Delete()
                .ActiveCell.Columns.EntireColumn.Delete()
                .ActiveCell.Columns.EntireColumn.Delete()
                .ActiveCell.Offset(0, -2).Select()
                Do Until .ActiveCell.Value = ""
                    If Len(.ActiveCell.Value) = 4 Then
                        .ActiveCell.Offset(0, -1).Value = Microsoft.VisualBasic.Left(.ActiveCell.Value, 2)

                    ElseIf Len(.ActiveCell.Value) = 3 Then
                        .ActiveCell.Offset(0, -1).Value = Microsoft.VisualBasic.Left(.ActiveCell.Value, 1)

                    ElseIf Len(.ActiveCell.Value) = 1 Then
                        .ActiveCell.Offset(0, -1).Value = Microsoft.VisualBasic.Left(.ActiveCell.Value, 1)
                    End If
                    .ActiveCell.Offset(1, 0).Select()
                Loop
                .ActiveCell.Columns.EntireColumn.Delete()
                .Cells.Range("A1").Select()
                .ActiveCell.Rows.EntireRow.Insert()
                .Cells.Range("A1").Value = "Elt."
                .Cells.Range("B1").Value = "Conc"
                ProgressBar2.Value = 7
                Label2.Text = "Formattazione in corso ..."
                If CheckBox1.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> "Mo" Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = "Mo" Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If
                If CheckBox2.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> "Nb" Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = "Nb" Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If


                If CheckBox3.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> "W" Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = "W" Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If

                If CheckBox4.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> TextBox2.Text Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = TextBox2.Text Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If
                If CheckBox5.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> TextBox3.Text Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = TextBox3.Text Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If
                If CheckBox6.Checked Then
                    .Range("A2").Select()
                    Do Until .ActiveCell.Value = ""
                        If .ActiveCell.Value <> TextBox4.Text Then
                            .ActiveCell.Offset(1, 0).Select()
                        End If
                        If .ActiveCell.Value = TextBox4.Text Then
                            .ActiveCell.EntireRow.Delete()
                        End If

                    Loop
                End If

                .ActiveWorkbook.Close(SaveChanges:=True)
                .Application.Quit()



            End With
        Next File

        ProgressBar2.Value = 10
        Label2.Text = "Formattazione completata!!"
    End Sub

    '**************************************************************************************************************************
    '--------------------------------------------------------------------------------------------------------------------------
    '**************************************************************************************************************************
    'HPLC
    '**************************************************************************************************************************
    '--------------------------------------------------------------------------------------------------------------------------
    '**************************************************************************************************************************

    '***************************************************************************************************************
    'Elaborazione 1
    '****************************************************************************************************************
    Private Sub button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim a As Integer = TextBox1.Text
        Dim b As Integer = TextBox5.Text
        Dim c As Integer = TextBox6.Text
        Dim d As Integer = TextBox7.Text
        Dim t As Integer = TextBox8.Text
        Dim f As Integer = 2
        label5.Text = ""
        If RadioButton9.Checked = True And RadioButton8.Checked = False Then
            If MsgBox("Seleziona la modalità automatica", 48, "Errore!") = vbOK Then
                Exit Sub
            End If
        End If
        If RadioButton8.Checked = True And RadioButton9.Checked = False Then
            Dim objXls As Microsoft.Office.Interop.Excel.Application
            objXls = New Microsoft.Office.Interop.Excel.Application
            objXls.Visible = True
            With objXls

                .Workbooks.Open("C:\HPLC\tabella spettri coloranti.xls")
                .Cells.Range("A2").Select()
                For i = 1 To 80
                    If a >= .ActiveCell.Value - f And a <= .ActiveCell.Value + f And b >= .ActiveCell.Offset(0, 1).Value - f And b <= .ActiveCell.Offset(0, 1).Value + f And c >= .ActiveCell.Offset(0, 2).Value - f And c <= .ActiveCell.Offset(0, 2).Value + f And d >= .ActiveCell.Offset(0, 3).Value - f And d <= .ActiveCell.Offset(0, 3).Value + f And t >= .ActiveCell.Offset(0, 4).Value - f And t <= .ActiveCell.Offset(0, 4).Value + f Then
                        .ActiveCell.Offset(0, 5).Select()
                        ListBox6.Items.Add(.ActiveCell.Value)
                        .ActiveCell.Offset(0, -5).Select()
                        .ActiveCell.Offset(1, 0).Select()
                    Else
                        .ActiveCell.Offset(1, 0).Select()

                    End If
                Next i

                If ListBox6.Items.Count = 0 Then
                    label5.Text = "Nessuna corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 1 Then
                    label5.Text = "1 corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 2 Then
                    label5.Text = "2 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 3 Then
                    label5.Text = "3 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 4 Then
                    label5.Text = "4 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 5 Then
                    label5.Text = "5 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 6 Then
                    label5.Text = "6 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 7 Then
                    label5.Text = "7 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 8 Then
                    label5.Text = "8 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 9 Then
                    label5.Text = "9 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 10 Then
                    label5.Text = "10 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count > 10 Then
                    label5.Text = "Più di 10 corrispondenze trovate!"
                End If


                .ActiveWorkbook.Close(SaveChanges:=False)
                .Application.Quit()

            End With

        End If
    End Sub
    '***********************************************************************************************************
    'Tasto Cancella
    '***********************************************************************************************************
    Private Sub button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        ListBox6.Items.Clear()
        TextBox1.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        PictureBox1.Image = Nothing
        label5.Text = ""
        TextBox9.Clear()
    End Sub

    Private Sub button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '******************************************************************************************************************
        'Elaborazione 2
        '******************************************************************************************************************
        Dim a As Integer = TextBox1.Text
        Dim b As Integer = TextBox5.Text
        Dim c As Integer = TextBox6.Text
        Dim d As Integer = TextBox7.Text
        Dim t As Integer = TextBox8.Text
        Dim f As Integer = 5
        label5.Text = ""
        If RadioButton9.Checked = True And RadioButton8.Checked = False Then
            If MsgBox("Seleziona la modalità automatica", 48, "Errore!") = vbOK Then
                Exit Sub
            End If
        End If

        If RadioButton8.Checked = True And RadioButton9.Checked = False Then
            Dim objXls As Microsoft.Office.Interop.Excel.Application
            objXls = New Microsoft.Office.Interop.Excel.Application
            objXls.Visible = True
            With objXls

                .Workbooks.Open("C:\HPLC\tabella spettri coloranti.xls")
                .Cells.Range("A2").Select()
                For i = 1 To 80
                    If a >= .ActiveCell.Value - f And a <= .ActiveCell.Value + f And b >= .ActiveCell.Offset(0, 1).Value - f And b <= .ActiveCell.Offset(0, 1).Value + f And c >= .ActiveCell.Offset(0, 2).Value - f And c <= .ActiveCell.Offset(0, 2).Value + f And d >= .ActiveCell.Offset(0, 3).Value - f And d <= .ActiveCell.Offset(0, 3).Value + f And t >= .ActiveCell.Offset(0, 4).Value - f And t <= .ActiveCell.Offset(0, 4).Value + f Then
                        .ActiveCell.Offset(0, 5).Select()
                        ListBox6.Items.Add(.ActiveCell.Value)
                        .ActiveCell.Offset(0, -5).Select()
                        .ActiveCell.Offset(1, 0).Select()
                    Else
                        .ActiveCell.Offset(1, 0).Select()

                    End If
                Next i

                If ListBox6.Items.Count = 0 Then
                    label5.Text = "Nessuna corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 1 Then
                    label5.Text = "1 corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 2 Then
                    label5.Text = "2 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 3 Then
                    label5.Text = "3 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 4 Then
                    label5.Text = "4 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 5 Then
                    label5.Text = "5 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 6 Then
                    label5.Text = "6 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 7 Then
                    label5.Text = "7 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 8 Then
                    label5.Text = "8 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 9 Then
                    label5.Text = "9 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 10 Then
                    label5.Text = "10 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count > 10 Then
                    label5.Text = "Più di 10 corrispondenze trovate!"
                End If


                .ActiveWorkbook.Close(SaveChanges:=False)
                .Application.Quit()

            End With
        End If
    End Sub

    Private Sub button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '************************************************************************************************************
        'Elaborazione 3
        '************************************************************************************************************
        Dim a As Integer = TextBox1.Text
        Dim b As Integer = TextBox5.Text
        Dim c As Integer = TextBox6.Text
        Dim d As Integer = TextBox7.Text
        Dim t As Integer = TextBox8.Text
        Dim f As Integer = 10
        Dim k As Integer
        label5.Text = ""
        If RadioButton9.Checked = True And RadioButton8.Checked = False Then
            If MsgBox("Seleziona la modalità automatica", 48, "Errore!") = vbOK Then
                Exit Sub
            End If
        End If
        If RadioButton8.Checked = True And RadioButton9.Checked = False Then
            Dim objXls As Microsoft.Office.Interop.Excel.Application
            objXls = New Microsoft.Office.Interop.Excel.Application
            objXls.Visible = True
            With objXls

                .Workbooks.Open("C:\HPLC\tabella spettri coloranti.xls")
                .Cells.Range("A2").Select()
                For i = 1 To 80
                    If a >= .ActiveCell.Value - f And a <= .ActiveCell.Value + f And b >= .ActiveCell.Offset(0, 1).Value - f And b <= .ActiveCell.Offset(0, 1).Value + f And c >= .ActiveCell.Offset(0, 2).Value - f And c <= .ActiveCell.Offset(0, 2).Value + f And d >= .ActiveCell.Offset(0, 3).Value - f And d <= .ActiveCell.Offset(0, 3).Value + f And t >= .ActiveCell.Offset(0, 4).Value - f And t <= .ActiveCell.Offset(0, 4).Value + f Then
                        .ActiveCell.Offset(0, 5).Select()
                        ListBox6.Items.Add(.ActiveCell.Value)
                        .ActiveCell.Offset(0, -5).Select()
                        .ActiveCell.Offset(1, 0).Select()
                    Else
                        .ActiveCell.Offset(1, 0).Select()

                    End If
                Next i

                If ListBox6.Items.Count = 0 Then
                    label5.Text = "Nessuna corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 1 Then
                    label5.Text = "1 corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 2 Then
                    label5.Text = "2 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 3 Then
                    label5.Text = "3 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 4 Then
                    label5.Text = "4 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 5 Then
                    label5.Text = "5 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 6 Then
                    label5.Text = "6 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 7 Then
                    label5.Text = "7 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 8 Then
                    label5.Text = "8 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 9 Then
                    label5.Text = "9 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 10 Then
                    label5.Text = "10 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count > 10 Then
                    label5.Text = "Più di 10 corrispondenze trovate!"
                End If


                .ActiveWorkbook.Close(SaveChanges:=False)
                .Application.Quit()

            End With
        End If
    End Sub
    '*********************************************************************************************************************
    'Modalità manuale
    '*********************************************************************************************************************

    Private Sub button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If RadioButton8.Checked = True And RadioButton9.Checked = False Then
            If MsgBox("Seleziona la modalità manuale", 48, "Errore!") = vbOK Then
                Exit Sub
            End If
        End If
        If TextBox9.Text = "" Then
            If MsgBox("Inserire il valore di ricerca", 48, "Errore!") = vbOK Then
                Exit Sub
            End If
        End If

        Dim a As Integer = TextBox1.Text
        Dim b As Integer = TextBox5.Text
        Dim c As Integer = TextBox6.Text
        Dim d As Integer = TextBox7.Text
        Dim t As Integer = TextBox8.Text
        Dim p As Integer = TextBox9.Text
        Dim k As Integer
        If RadioButton9.Checked = True And RadioButton8.Checked = False Then
            Dim objXls As Microsoft.Office.Interop.Excel.Application
            objXls = New Microsoft.Office.Interop.Excel.Application
            objXls.Visible = True
            With objXls

                .Workbooks.Open("C:\HPLC\tabella spettri coloranti.xls")
                .Cells.Range("A2").Select()
                For i = 1 To 80
                    If a >= .ActiveCell.Value - p And a <= .ActiveCell.Value + p And b >= .ActiveCell.Offset(0, 1).Value - p And b <= .ActiveCell.Offset(0, 1).Value + p And c >= .ActiveCell.Offset(0, 2).Value - p And c <= .ActiveCell.Offset(0, 2).Value + p And d >= .ActiveCell.Offset(0, 3).Value - p And d <= .ActiveCell.Offset(0, 3).Value + p And t >= .ActiveCell.Offset(0, 4).Value - p And t <= .ActiveCell.Offset(0, 4).Value + p Then
                        .ActiveCell.Offset(0, 5).Select()
                        ListBox6.Items.Add(.ActiveCell.Value)
                        .ActiveCell.Offset(0, -5).Select()
                        .ActiveCell.Offset(1, 0).Select()
                    Else
                        .ActiveCell.Offset(1, 0).Select()

                    End If
                Next i

                If ListBox6.Items.Count = 0 Then
                    label5.Text = "Nessuna corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 1 Then
                    label5.Text = "1 corrispondenza trovata!"
                ElseIf ListBox6.Items.Count = 2 Then
                    label5.Text = "2 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 3 Then
                    label5.Text = "3 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 4 Then
                    label5.Text = "4 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 5 Then
                    label5.Text = "5 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 6 Then
                    label5.Text = "6 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 7 Then
                    label5.Text = "7 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 8 Then
                    label5.Text = "8 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 9 Then
                    label5.Text = "9 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count = 10 Then
                    label5.Text = "10 corrispondenze trovate!"
                ElseIf ListBox6.Items.Count > 10 Then
                    label5.Text = "Più di 10 corrispondenze trovate!"
                End If
                .ActiveWorkbook.Close(SaveChanges:=False)
                .Application.Quit()

            End With
        End If


    End Sub



    Private Sub button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim z As String
        z = ListBox6.SelectedItem
        If InStr(z, "alizarina - robbia") > 0 Then
            PictureBox1.Load("C:\HPLC\alizarina - robbia.jpg")
        End If
        If InStr(z, "rubiadina - robbia") > 0 Then
            PictureBox1.Load("C:\HPLC\rubiadina - robbia.jpg")
        End If
        If InStr(z, "6,6 dibromoindaco - porpora") > 0 Then
            PictureBox1.Load("C:\HPLC\6,6 dibromoindaco - porpora.jpg")
        End If
        If InStr(z, "acido carminico - cocciniglia") > 0 Then
            PictureBox1.Load("C:\HPLC\acido carminico - cocciniglia.jpg")
        End If
        If InStr(z, "acido ellagico - tannini") > 0 Then
            PictureBox1.Load("C:\HPLC\acido ellagico - tannini.jpg")
        End If
        If InStr(z, "acido flavokermesico - cocciniglia") > 0 Then
            PictureBox1.Load("C:\HPLC\acido flavokermesico - cocciniglia.jpg")
        End If
        If InStr(z, "acido gallico - tannini") > 0 Then
            PictureBox1.Load("C:\HPLC\acido gallico - tannini.jpg")
        End If
        If InStr(z, "acido kermesico - cocciniglia") > 0 Then
            PictureBox1.Load("C:\HPLC\acido kermesico - cocciniglia.jpg")
        End If
        If InStr(z, "acido laccaico A+B - kerria lacca") > 0 Then
            PictureBox1.Load("C:\HPLC\acido laccaico A+B - kerria lacca.jpg")
        End If
        If InStr(z, "acido laccaico C - kerria lacca") > 0 Then
            PictureBox1.Load("C:\HPLC\acido laccaico C - kerria lacca.jpg")
        End If
        If InStr(z, "alfa amminorceina - oricello") > 0 Then
            PictureBox1.Load("C:\HPLC\alfa amminorceina - oricello.jpg")
        End If
        If InStr(z, "alfa idrossiorceina - oricello") > 0 Then
            PictureBox1.Load("C:\HPLC\alfa idrossiorceina - oricello.jpg")
        End If
        If InStr(z, "apigenina - erba guada") > 0 Then
            PictureBox1.Load("C:\HPLC\apigenina - erba guada.jpg")
        End If
        If InStr(z, "bergenina - endopleura uchi") > 0 Then
            PictureBox1.Load("C:\HPLC\bergenina - endopleura uchi.jpg")
        End If
        If InStr(z, "beta,gamma amminorceinimmina - oricello") > 0 Then
            PictureBox1.Load("C:\HPLC\beta,gamma amminorceinimmina - oricello.jpg")
        End If
        If InStr(z, "brasileina - sequoia legno rosso 2") > 0 Then
            PictureBox1.Load("C:\HPLC\brasileina - sequoia legno rosso 2.jpg")
        End If
        If InStr(z, "cartamo 1") > 0 Then
            PictureBox1.Load("C:\HPLC\cartamo 1.jpg")
        End If
        If InStr(z, "cartamo 2") > 0 Then
            PictureBox1.Load("C:\HPLC\cartamo 2.jpg")
        End If
        If InStr(z, "cartamo 3") > 0 Then
            PictureBox1.Load("C:\HPLC\cartamo 3.jpg")
        End If
        If InStr(z, "catechina - tannini") > 0 Then
            PictureBox1.Load("C:\HPLC\catechina - tannini.jpg")
        End If
        If InStr(z, "curcuma") > 0 Then
            PictureBox1.Load("C:\HPLC\curcuma.jpg")
        End If
        If InStr(z, "dCII - cocciniglia") > 0 Then
            PictureBox1.Load("C:\HPLC\dCII - cocciniglia.jpg")
        End If
        If InStr(z, "dCIV - cocciniglia") > 0 Then
            PictureBox1.Load("C:\HPLC\dCIV - cocciniglia.jpg")
        End If
        If InStr(z, "deidroematina - campeggio") > 0 Then
            PictureBox1.Load("C:\HPLC\deidroematina - campeggio.jpg")
        End If
        If InStr(z, "emodina - corteccia bacche spinose") > 0 Then
            PictureBox1.Load("C:\HPLC\emodina - corteccia bacche spinose.jpg")
        End If
        If InStr(z, "fisetina - cotino") > 0 Then
            PictureBox1.Load("C:\HPLC\fisetina - cotino.jpg")
        End If
        If InStr(z, "genisteina - ginestra") > 0 Then
            PictureBox1.Load("C:\HPLC\genisteina - ginestra.jpg")
        End If
        If InStr(z, "indigotina") > 0 Then
            PictureBox1.Load("C:\HPLC\indigotina.jpg")
        End If
        If InStr(z, "indirubina") > 0 Then
            PictureBox1.Load("C:\HPLC\indirubina.jpg")
        End If
        If InStr(z, "isatina") > 0 Then
            PictureBox1.Load("C:\HPLC\isatina.jpg")
        End If
        If InStr(z, "juglone - noce nero") > 0 Then
            PictureBox1.Load("C:\HPLC\juglone - noce nero.jpg")
        End If
        If InStr(z, "kaempferolo - legno giallo") > 0 Then
            PictureBox1.Load("C:\HPLC\kaempferolo - legno giallo.jpg")
        End If
        If InStr(z, "lawsone - noce nero") > 0 Then
            PictureBox1.Load("C:\HPLC\lawsone - noce nero.jpg")
        End If
        If InStr(z, "legno del brasile 1") > 0 Then
            PictureBox1.Load("C:\HPLC\legno del brasile 1.jpg")
        End If
        If InStr(z, "legno del brasile 2") > 0 Then
            PictureBox1.Load("C:\HPLC\legno del brasile 2.jpg")
        End If
        If InStr(z, "legno del brasile 3") > 0 Then
            PictureBox1.Load("C:\HPLC\legno del brasile 3.jpg")
        End If
        If InStr(z, "legno del brasile 4") > 0 Then
            PictureBox1.Load("C:\HPLC\legno del brasile 4.jpg")
        End If
        If InStr(z, "legno del brasile 5") > 0 Then
            PictureBox1.Load("C:\HPLC\legno del brasile 5.jpg")
        End If
        If InStr(z, "luteolina - erba guada") > 0 Then
            PictureBox1.Load("C:\HPLC\luteolina - erba guada.jpg")
        End If
        If InStr(z, "mallo di noce 1") > 0 Then
            PictureBox1.Load("C:\HPLC\mallo di noce 1.jpg")
        End If
        If InStr(z, "noce di galla 1") > 0 Then
            PictureBox1.Load("C:\HPLC\noce di galla 1.jpg")
        End If
        If InStr(z, "noce di galla 2") > 0 Then
            PictureBox1.Load("C:\HPLC\noce di galla 2.jpg")
        End If
        If InStr(z, "oricello") > 0 Then
            PictureBox1.Load("C:\HPLC\oricello.jpg")
        End If
        If InStr(z, "quercetina - quercus e rhamnus") > 0 Then
            PictureBox1.Load("C:\HPLC\quercetina - quercus e rhamnus.jpg")
        End If
        If InStr(z, "roccellina") > 0 Then
            PictureBox1.Load("C:\HPLC\roccellina.jpg")
        End If
        If InStr(z, "rosso metile") > 0 Then
            PictureBox1.Load("C:\HPLC\rosso metile.jpg")
        End If
        If InStr(z, "purpurina - robbia") > 0 Then
            PictureBox1.Load("C:\HPLC\purpurina - robbia.jpg")
        End If
        If InStr(z, "sulfuretina - cotino") > 0 Then
            PictureBox1.Load("C:\HPLC\sulfuretina - cotino.jpg")
        End If
        If InStr(z, "type C component - verzino") > 0 Then
            PictureBox1.Load("C:\HPLC\type C component - verzino.jpg")
        End If
        If InStr(z, "zafferano 1") > 0 Then
            PictureBox1.Load("C:\HPLC\zafferano 1.jpg")
        End If
        If InStr(z, "zafferano 2") > 0 Then
            PictureBox1.Load("C:\HPLC\zafferano 2.jpg")
        End If

    End Sub


    
    
End Class










