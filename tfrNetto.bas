' Programma in BASIC LibreOffice che calcola il TFR netto
' a partire dal TFR lordo. Scaglioni aggiornati alle aliquote
' IRPEF 2024.
' Sono necessarie alcune informazioni:
' 1. TFR Lordo (in cella "TFR_LORDO")
' 2. Data di assunzione (in cella "DATA_ASSUNZIONE")
' 3. Il nome del foglio (attualmente "Foglio 1")

Sub CalcolaTfrNetto

	Dim tfrLordo As Variant
	Dim dataAssunzione As Date
	Dim mesiPassatiDaAssunzione as integer
	Dim redditoAnnuoMedio as double	
	Dim aliquotaIrpefMedia as double
	Dim aliquotaTfrRiferimento as double
	Dim tfrNetto as double
    Dim sheet As Object
    
    sheet = ThisComponent.Sheets.getByName("Foglio 1")
    
    If Not IsNull(sheet) Then
        tfrLordo = sheet.getCellRangeByName("TFR_LORDO").getValue()
        dataAssunzione = sheet.getCellRangeByName("DATA_ASSUNZIONE").getValue()
        mesiPassatiDaAssunzione = calcolaMesiPassatiDaAssunzione(dataAssunzione)
        redditoAnnuoMedio = tfrLordo * 12 / (mesiPassatiDaAssunzione / 12)
        aliquotaIrpefMedia = calcolaAliquotaIrpefMedia(redditoAnnuoMedio)
        aliquotaTfrRiferimento =  aliquotaIrpefMedia / redditoAnnuoMedio
        tfrNetto = tfrLordo * (1-aliquotaTfrRiferimento)
        sheet.getCellRangeByName("TFR_NETTO").setValue(tfrNetto)
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
	Dim scaglione1 as integer
	Dim scaglione2 as integer
	Dim scaglione3 as integer
	Dim scaglione4 as integer
	Dim scaglione5 as integer
	Dim sottrazione as integer

	' Calcolo scaglione 1
	scaglione1 = IIf(redditoAnnuoMedio > 28000, 28000, redditoAnnuoMedio)
	sottrazione = 28000
	
	' calcolo scaglione 2
	if redditoAnnuoMedio >= 28001 then
		scaglione2 = IIf(redditoAnnuoMedio > 50000, 50000, redditoAnnuoMedio)
		scaglione2 = scaglione2 - sottrazione
		sottrazione = 50000
	else
		scaglione2 = 0
	end if
	
	' calcolo scaglione 3
	if redditoAnnuoMedio >= 50001 then
		scaglione3 = redditoAnnuoMedio
		scaglione3 = scaglione3 - sottrazione
	else
		scaglione3 = 0
	end if
	
	Dim totale as double
	totale = 0.23 * scaglione1 + 0.35 * scaglione2 + 0.43 * scaglione3

	calcolaAliquotaIrpefMedia = totale
end function


