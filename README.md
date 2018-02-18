# DEMO_EXCEL_VBA_DoEvents
## Demo-Code der DoEvents-Anweisung in Excel VBA


````vba
'DoEvents
'--------

'Versuch 1:
'   - DoEvents auskommentieren
'   - Schleife starten (mit dem Button auf dem Tabellenblatt oder F5 im Code)
'   - Die Schleife läuft 100000x durch und schreibt den Wert jeweils in E5
'   - Excel ist bis zum Ende der Schleife "eingefroren"
'   - Wenn der Button "Nachricht" geklickt wird kommt die MessageBox
'     sobald die Schleife fertig ist

'Versuch 2:
'   - DoEvents einkommentieren
'   - Schleife starten (mit dem Button auf dem Tabellenblatt oder F5 im Code)
'   - Die Schleife läuft 100000x durch und schreibt den Wert jeweils in E5
'   - Wenn der Button "Nachricht" geklickt wird kommt die MessageBox sofort
'     und die Schleife macht erst weiter, wenn die Nachricht weggeklickt wird

'Versuch 3:
'   - DoEvents einkommentieren
'   - Schleife starten (mit dem Button auf dem Tabellenblatt oder F5 im Code)
'   - Die Schleife läuft 100000x durch und schreibt den Wert jeweils in E5
'   - Wenn ein Wert in irgendeine Zelle eingetragen wird bricht die Schleife ab

Sub Schleife()
    Dim i As Long
    Do Until i > 100000
        Range("E5").Value = i
        i = i + 1
        DoEvents
    Loop
End Sub


Sub Nachricht()
    MsgBox "Moin"
End Sub
````
