VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8220.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4665
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Modulweite Variablen
Private doc As Document
Private ch As String
Private num As Integer
Private schluessel As Integer
Private i As Long
Dim SchluesselListe() As Integer ' Schlüssel-Array

Private Sub UserForm_Initialize()
    Randomize
End Sub

Private Sub btnEnkrypten_Click()
    Set doc = ActiveDocument
    Dim textContent As String
    textContent = doc.Content.text

    Dim totalLen As Long
    totalLen = Len(textContent)

    ReDim SchluesselListe(1 To totalLen)

    Dim resultText As String
    resultText = ""

    For i = 1 To totalLen
        ch = Mid(textContent, i, 1)

        If Asc(ch) >= 32 And Asc(ch) <= 126 Then
            schluessel = Int(95 * Rnd)
            SchluesselListe(i) = schluessel
            num = Asc(ch) - 32
            num = (num + schluessel) Mod 95
            resultText = resultText & Chr(num + 32)
        Else
            SchluesselListe(i) = 0
            resultText = resultText & ch
        End If
    Next i

    ' Text zurückschreiben
    doc.Content.text = resultText

    ' ? Entferne überschüssiges Absatzzeichen, wenn zwei am Ende
    If Len(doc.Content.text) >= 2 Then
        If Right(doc.Content.text, 2) = vbCr & vbCr Then
            doc.Content.Characters(doc.Content.Characters.Count).Delete
        End If
    End If
End Sub

Private Sub btnDekrypten_Click()
    Dim selectedKeyIndex As Long
    selectedKeyIndex = txtSchluessel.ListIndex + 1

    If selectedKeyIndex > 0 And selectedKeyIndex <= Len(doc.Content.text) Then
        Dim ch As String
        ch = Mid(doc.Content.text, selectedKeyIndex, 1)

        If Asc(ch) >= 32 And Asc(ch) <= 126 Then
            Dim decryptedChar As String
            Dim schluessel As Integer
            schluessel = SchluesselListe(selectedKeyIndex)

            decryptedChar = Chr(((Asc(ch) - 32 - schluessel + 95) Mod 95) + 32)

            ' Ersetze das Zeichen im Dokument
            doc.Range(Start:=selectedKeyIndex - 1, End:=selectedKeyIndex).text = decryptedChar

            ' Markiere den Schlüssel als "verbraucht"
            SchluesselListe(selectedKeyIndex) = 0
        End If
    End If
End Sub

Private Sub BtnShowKeys_Click()
    txtSchluessel.Clear
    Set doc = ActiveDocument

    For i = 1 To UBound(SchluesselListe)
        ch = doc.Content.Characters(i).text
        If Asc(ch) = 13 Then
            txtSchluessel.AddItem "¶"
        Else
            txtSchluessel.AddItem SchluesselListe(i)
        End If
    Next i
End Sub

Public Sub DecryptAllEntries()
    Set doc = ActiveDocument

    If Not Not SchluesselListe Then
        For i = UBound(SchluesselListe) To 1 Step -1
            ch = doc.Content.Characters(i).text
            schluessel = SchluesselListe(i)

            If Asc(ch) >= 32 And Asc(ch) <= 126 And schluessel <> 0 Then
                num = Asc(ch) - 32
                num = (num - schluessel + 95) Mod 95
                doc.Content.Characters(i).text = Chr(num + 32)
                SchluesselListe(i) = 0
            End If
        Next i

        Unload UserForm2
        Unload UserForm1
    Else
        MsgBox "Es wurden noch keine Zeichen Verschlüsselt.", vbExclamation
    End If
End Sub

Private Sub btnOpenUserForm2_Click()
    UserForm2.Show
End Sub


