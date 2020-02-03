Attribute VB_Name = "Modul1"
Sub DatenSummierenProLandUndMonat()
   Dim cn As Object
   Dim rs As Object
   Dim strConnection As String
   Dim strSQL As String

  Set cn = CreateObject("ADODB.CONNECTION")

 'Den Treiber bekanntgeben
   strConnection = "DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}; DBQ=" & ThisWorkbook.FullName

   tbl_Ziel.UsedRange.Clear
   tbl_Ziel.Range("A1:D1").Value = Array("Produkt", "Monat", "Land", "Umsatz")

   With cn

   'Datenverbindung öffnen
   .Open strConnection
    
   'Abfragestring zusammenbasteln und Abfrage starten
     strSQL = "SELECT Produkt, MONTH(Datum) as Monat, Land, SUM(Umsatz) as Summe FROM [Quelle$] " & _
              "WHERE Produkt IN (11, 21) GROUP BY Produkt, Land, MONTH(Datum) ORDER BY Produkt, MONTH(Datum)"

     Set rs = CreateObject("ADODB.RECORDSET")
     With rs
      .Source = strSQL
      .ActiveConnection = strConnection
      .Open
       tbl_Ziel.Range("A2").CopyFromRecordset rs
      .Close
     End With
 
   End With

 'ADO-Verbindung kappen
   cn.Close

   Set cn = Nothing
   Set rs = Nothing

End Sub

