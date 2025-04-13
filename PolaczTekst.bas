Attribute VB_Name = "Module1"
Sub PolaczTekstWJednejKomorce()
Attribute PolaczTekstWJednejKomorce.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim komorka As Range
    Dim tekstZlaczony As String
    Dim pierwszyZakres As Range
    Dim zakres As Range

    ' Sprawdzenie, czy coœ jest zaznaczone
    If TypeName(Selection) <> "Range" Then
        MsgBox "Zaznacz komórki z tekstem do po³¹czenia."
        Exit Sub
    End If

    Set zakres = Selection
    Set pierwszyZakres = zakres.Cells(1, 1)

    tekstZlaczony = ""

    ' £¹czenie tekstu z wszystkich komórek
    For Each komorka In zakres
        If komorka.Value <> "" Then
            tekstZlaczony = tekstZlaczony & komorka.Value & " "
        End If
    Next komorka

    tekstZlaczony = Trim(tekstZlaczony)

    ' Wstawienie tekstu do pierwszej komórki
    pierwszyZakres.Value = tekstZlaczony

    ' Wyczyœæ pozosta³e komórki
    For Each komorka In zakres
        If komorka.Address <> pierwszyZakres.Address Then
            komorka.ClearContents
        End If
    Next komorka
End Sub

