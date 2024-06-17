' Programma in BASIC LibreOffice che calcola il TFR netto
' a partire dal TFR lordo. Scaglioni aggiornati alle aliquote
' IRPEF 2024.
' Sono necessarie alcune informazioni:
' 1. TFR Lordo (in cella "TFR_LORDO")
' 2. Data di assunzione (in cella "DATA_ASSUNZIONE")
' 3. Il nome del foglio (attualmente "Foglio 1")
' Il risultato verrà memorizzato nella cella "TFR_NETTO"

Sub CalcolaTfrNetto

	Dim tfrLordo As Double
	Dim dataAssunzione As Date
	Dim mesiPassatiDaAssunzione as Integer
	Dim redditoAnnuoMedio as Double
	Dim aliquotaIrpefMedia as Double
	Dim aliquotaTfrRiferimento as Double
	Dim tfrNetto as Double
	Dim sheet As Object

	sheet = ThisComponent.Sheets.getByName("Foglio 1")

	If Not IsNull(sheet) Then
		tfrLordo = sheet.getCellRangeByName("TFR_LORDO").getValue()
		dataAssunzione = sheet.getCellRangeByName("DATA_ASSUNZIONE").getValue()
		mesiPassatiDaAssunzione = calcolaMesiPassatiDaAssunzione(dataAssunzione)
		redditoAnnuoMedio = tfrLordo * 12 / (mesiPassatiDaAssunzione / 12)
		aliquotaIrpefMedia = calcolaAliquotaIrpefMedia(redditoAnnuoMedio)
		aliquotaTfrRiferimento = aliquotaIrpefMedia / redditoAnnuoMedio
		tfrNetto = tfrLordo * (1 - aliquotaTfrRiferimento)
		sheet.getCellRangeByName("TFR_NETTO").setValue(Format(tfrNetto, "0.00"))
	Else
		MsgBox "Foglio 'Foglio 1' non trovato."
	End If

End Sub

Function calcolaMesiPassatiDaAssunzione(ByVal dataAssunzione as date) As Integer
	Dim today As Date
	today = Date
	calcolaMesiPassatiDaAssunzione = (Year(today) - Year(dataAssunzione)) * 12 + (Month(today) - Month(dataAssunzione))
End Function


Function calcolaAliquotaIrpefMedia(ByVal redditoAnnuoMedio) as double
	Dim scaglione1 as Integer
	Dim scaglione2 as Integer
	Dim scaglione3 as Integer
	Dim totale as Double
	
	scaglione1 = 0
	scaglione2 = 0
	scaglione3 = 0
	
	' Calcolo scaglione 1
	scaglione1 = IIf(redditoAnnuoMedio > 28000, 28000, redditoAnnuoMedio)

	' calcolo scaglione 2
	If redditoAnnuoMedio >= 28001 Then
		scaglione2 = IIf(redditoAnnuoMedio > 50000, 50000, redditoAnnuoMedio)
		scaglione2 = scaglione2 - 28000
	End If

	' calcolo scaglione 3
	If redditoAnnuoMedio >= 50001 Then
		scaglione3 = redditoAnnuoMedio
		scaglione3 = scaglione3 - 50000
	End If

	
	totale = 0.23 * scaglione1 + 0.35 * scaglione2 + 0.43 * scaglione3

	calcolaAliquotaIrpefMedia = totale
End Function
