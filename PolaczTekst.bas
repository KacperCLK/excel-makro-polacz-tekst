Attribute VB_Name = "Module"
Sub PolaczTekstWJednejKomorce()
Attribute PolaczTekstWJednejKomorce.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim komorka As Range
    Dim tekstZlaczony As String
    Dim pierwszyZakres As Range
    Dim zakres As Range

    ' Sprawdzenie, czy co� jest zaznaczone
    If TypeName(Selection) <> "Range" Then
        MsgBox "Zaznacz kom�rki z tekstem do po��czenia."
        Exit Sub
    End If

    Set zakres = Selection
    Set pierwszyZakres = zakres.Cells(1, 1)

    tekstZlaczony = ""

    ' ��czenie tekstu z wszystkich kom�rek
    For Each komorka In zakres
        If komorka.Value <> "" Then
            tekstZlaczony = tekstZlaczony & komorka.Value & " "
        End If
    Next komorka

    tekstZlaczony = Trim(tekstZlaczony)

    ' Wstawienie tekstu do pierwszej kom�rki
    pierwszyZakres.Value = tekstZlaczony

    ' Wyczy�� pozosta�e kom�rki
    For Each komorka In zakres
        If komorka.Address <> pierwszyZakres.Address Then
            komorka.ClearContents
        End If
    Next komorka
End Sub

