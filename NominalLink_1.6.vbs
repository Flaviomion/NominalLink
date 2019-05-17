Option Explicit
Const Versione = "1.6"
Const Name = "NominalLink"
Dim Fdebug
Dim sheet, objExcel, objWorkbook, objWorksheet, objRange, objShell, out, report, objStdOut
Dim CDir, indice, indice_S, y, z, vInput, vOutput, Uscita, vTest, file_output, sheet_n, n_fileIn
Dim ris, idx, scritto, modificato, bReadOnly
Dim folderDest, fileDest, file_dest, fileINI, NomeProgramma
Fdebug = false
fileINI = "NominalLink.ini"
NomeProgramma = Name&"_"&Versione
'folderDest = WScript.Arguments.Item(0)
'fileDest = WScript.Arguments.Item(1)
Dim lista(12,100)		 ' 0	nome del Parametro
                         ' 1	directory			(D)
                         ' 2	nome file			(D)
                         ' 3	chiave_file			(D)
                         ' 4 	sheet				(D)
                         ' 5 	colonna n.			(D)
                         ' 6 	colonna lettrere	(D)
                         ' 7 	riga				(D)
                         ' 8 	Directory			(S)
						 ' 9 	file				(S)
						 '10    chiave_file			(S)
						 '11    puntatore a listaS

Dim listaS(9,100)		 ' 0	nome del Parametro	(S)
                         ' 1	directory			(S)
                         ' 2	nome file			(S)
                         ' 3	chiave_file			(S)
                         ' 4 	sheet				(S)
                         ' 5 	colonna n.			(S)
                         ' 6 	colonna lettrere	(S)
                         ' 7 	riga				(S)
						 ' 8	trovato
						 
Dim listaFile(100,3) 	'0	dir
						'1	indice nella lista
						'2	file

Const OK_BUTTON = 0
Const CRITICAL_ICON = 16
Const INFO_ICON_YN = 36
Const INFO_ICON = 64
Const AUTO_DISMISS = 0
Const AttesaMessaggioVV = 1
Const AttesaMessaggioV = 2
Const AttesaMessaggio = 5
Const AttesaMessaggioL = 30
Const AttesaMessaggioLL = 60

Set objStdOut = WScript.StdOut
Set objShell = CreateObject("Wscript.Shell")
Set objExcel = CreateObject("Excel.Application")

leggiINI folderDest, fileDest, report, Fdebug

dim fso: set fso = CreateObject("Scripting.FileSystemObject")
CDir = fso.GetAbsolutePathName(".")
On Error Resume Next
objStdOut.Write "<font color ='blue'>Start "&NomeProgramma&"</font>"&vbCrLf
on error goto 0
objExcel.DisplayAlerts = 0
'objExcel.Visible = False

file_dest = cercaFile(fileDest, folderDest, fso)
ris = objShell.Popup("Elaboro il file:" & folderDest&"\"&file_dest, AttesaMessaggio, "Info", INFO_ICON_YN + 4096)
if (ris = 7) then
	On Error Resume Next
	objStdOut.Write "<font color ='orange'>Programma fermato dall'utente</font>"&vbCrLf
    Wscript.Quit 0
end if
on error resume next
objStdOut.Write "<font color ='Olive'>Elaboro il file:" & folderDest&"\"&file_dest&"</font>"&vbCrLf
on error goto 0
Set objWorkbook = objExcel.Workbooks.Open(folderDest&"\"&file_dest, False, False)
If (Err.Number <> 0) Then
    on error resume next
	objStdOut.Write "<font color='red'>Errore Open file Destinazione:"&folderDest&"\"&file_dest&" Descrizione Err: " & Err.Description&"</font>"&vbCrLf
    on error goto 0
	Wscript.Quit Err.Number
End If

bReadOnly = objWorkbook.ReadOnly
If bReadOnly = True Then
	On Error Resume Next
	objStdOut.Write "<font color ='red'>Errore apertura file "&folderDest&"\"&file_dest&" File OCCUPATO</font><br/>"&vbCrlf
	on error goto 0
    'objShell.popup "Errore apertura file "&folderDest&"\"&file_dest&" File OCCUPATO" , 5, "Errore", CRITICAL_ICON + 4096
	Call objWorkbook.Close
	objExcel.Quit
	On Error Resume Next
	objStdOut.Write  "<br/><font color ='orange'>Programma Terminato a causa di file Occupato</font>"&vbCrLf&vbCrLf
	on error goto 0
	Wscript.quit 1055
End If


sheet_n = objWorkbook.Sheets.Count
'msgBox "Numero di sheets : " & sheet_n
indice = 0 'indice viene incrementato da FindMyComments
'-- 1) del flow --- Legge da Input tutti i Flink# su tutti gli sheet
'Flink#D#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#parametro#Elink
'ricerca dei link nelfile di destinazione 
for sheet = 1 to sheet_n step 1
	Set objWorksheet = objWorkbook.Worksheets(sheet)
	If (Err.Number <> 0) Then
        On Error Resume Next
		objStdOut.Write "<font color='red'>Errore sheet file Destinazione, Descrizione Err: " & Err.Description&"</font>"&vbCrLf
        on error goto 0
		Wscript.quit Err.Number
	End If
	Set objRange = objWorksheet.UsedRange
	On Error Resume Next
	objStdOut.Write "Ricerca Flink su sheet: "&objWorksheet.name&vbCrLf
	on error goto 0
	FindMyComments objWorksheet, sheet, fileDest, folderDest&"\"&file_dest, lista, indice
Next
Uscita = "<html><body><table><tr><td colspan=10><h2><font color='blue'>" & NomeProgramma & "</font></h2></td></tr>"
Uscita = Uscita & "<tr><td colspan=10><font color ='DarkGreen'>Gestisce link a campi Nominali</font></td></tr>"
Uscita = Uscita & "<tr><td  colspan='10' align='center'><font color ='DarkBlue'>" & folderDest&"\"&file_dest &"</font></td></tr>"
Uscita = Uscita & "<tr><td colspan=10><font color ='blue'>Rapporto del "&Date&" ore:"&Time&"</font></td></tr>"
objExcel.ActiveWorkbook.Close False,folderDest&"\"&file_dest 'non mi accetta come Workbook objWorkbook

n_fileIn = popolaListaFile(listaFile, indice)

'-- 2)--- usa listaFile per cercare tutti i link Nominali che hanno un file di riferimento
indice_S = 0 'indice_S viene incrementato da FindMyComments
On Error Resume Next
objStdOut.Write "Ricerca Parametri sui file Sorgenti"&vbCrLf
on error goto 0
for z = 0 to n_fileIn-1 step 1
	if Not (listaFile(z,2 ) = "") then
		'ho il file
		loopSheet listaFile(z,0),listaFile(z,2)
	else
		'ricerca su tutti gli excel della sourceDir
		listaDir listaFile(z,0)
	end if
	objExcel.ActiveWorkbook.Close False,lista(z,0)
Next ' (utilizzando la listaFile come indice per andare a prendere il giusto file sulla lista lista)

'-- 3)--- Cerca su listaS i link corrispondenti a quelli riscontrati su lista 
On Error Resume Next
objStdOut.Write "Crea la corrispondenza fra Destinazione e Sorgente"&vbCrLf
on error goto 0
idx = 0
for y = 0 to indice-1 step 1
	ris = cerca(lista,y,listaS, indice_S, idx) 'cerca la corrispondena di Parametro D e S
	if (ris) then
		'WScript.Echo "TROVATO link " & lista(0,y) & " " & lista(1,y) & " " & lista(2,y) & " " & lista(3,y) & " " & lista(6,y) & " " & lista(7,y) & " " & lista(8,y) & " " & lista(9,y)
		lista(11,y) = idx   'metto su lista il link a listaS
		listaS(8,idx) = y   'metto su listaS il link a lista
	else
		On Error Resume Next
		objStdOut.Write "<font color='red'>Parametro non trovato nei Sorgenti:" & lista(0,y)&"</font>"&vbCrLf
		on error goto 0
		'objShell.Popup "Parametro non trovato nei Sorgenti:" & lista(0,y), AttesaMessaggio, "Parametro", INFO_ICON + 4096
		Uscita = Uscita & "<tr><td colspan='10'><font color='Coral'>Parametro non riscontrato nei file sorgenti " &lista(0,y)& "</font></td></tr>"
        lista(11,y) = -1  'metto su lista il link a listaS
		'listaS(8,idx) = -1 il link negativo non ha senso sulla lista dei Sorgenti
	end if
Next

esportaListe("S")'scrive
On Error Resume Next
objStdOut.Write "Ordina la lista per ottimizzare l'accesso ai files"&vbCrLf
on error goto 0
riordinaLista()

esportaListe("A")'appende
On Error Resume Next
objStdOut.Write "<font color='blue'>Esegue la copia dei dati Sorgenti su Destinatari</font>"&vbCrLf
on error goto 0
'------ Legge tutti i link da lista ed esegue la migrazione dati 5) del flow
Dim i_l, precedente, lettura, scrittura, obJWK, objWorkbWrite
precedente = "primo"
scrittura = lista(1,0)&"\"&lista(2,0)
Set  objWorkbWrite = objExcel.Workbooks.Open(scrittura,0,False)
If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore Apertura Destinazione Description " & Err.Description&"</font>"&vbCrLf
	Wscript.Quit = Err.Number
End If
for y = 0 to indice_S-1 step 1
	i_l = listaS(8,y)
	if (i_l <> -1) then
		lettura = listaS(1,y)&"\"&listaS(2,y)
		if Not (precedente = lettura) then
		   ' chiudo il vecchio file di lettura
		   if Not (precedente = "primo") then
			   'Set obJWK = ActiveWorkbook
			   Call obJWK.Close(False,precedente)
		   end if
		   'apro un nuovo file di lettura
		   set obJWK = objExcel.Workbooks.Open(lettura, False, True)
		   If (Err.Number <> 0) Then
				On Error Resume Next
				objStdOut.Write "<font color='red'>Errore Apertura Sorgente Description " & Err.Description&"</font>"&vbCrLf
				Wscript.Quit Err.Number
		   End If
		end if
		vInput = ReadExcelFile(obJWK,listaS(4,y),listaS(7,y),listaS(5,y)) 'Workbook,sInput,rInput,cInput,objExcel
		Uscita = Uscita & "<tr><td><font color='green'>Parametro: "&listaS(0,y) &"</font></td><td><font color='green'>file sorgente:"&lettura& " s:"&  listaS(4,y) & " r:"&  listaS(7,y) & " c:"&  listaS(6,y) & "</font></td><td> Dato: " & vInput & "</td></tr>" 
		If (Err.Number <> 0) Then
			On Error Resume Next
			objStdOut.Write "<font color='red'>Errore nel Dato:" &listaS(0,y) &" Errore:"& Err.Number & " Description " & Err.Descriptio&"</font>"&vbCrLf
			on error goto 0
			'objShell.popup "Errore nel Dato:" &listaS(0,y) &" Errore:"& Err.Number & " Description " & Err.Description, AttesaMessaggio, "errore", CRITICAL_ICON + 4096
            Uscita = Uscita & "<tr><td colspan='10'><font color='red'>Errore nel Dato:" &listaS(0,y) &" "&lettura& "</td></tr>"
			Err.Clear
		End If
        vOutput = WEF(objWorkbWrite,lista(4,i_l),lista(7,i_l),lista(5,i_l),vInput) 'file,sh,rOutput,cOutput,valore
		vTest = ReadExcelFile(objWorkbWrite,lista(4,i_l),lista(7,i_l),lista(5,i_l))
		if (vOutput <> vTest) then
			On Error Resume Next
			objStdOut.Write "<font color='red'>Errore Dato Non Scritto:" & scrittura & " s:"&  lista(4,i_l) & " r:"&  lista(7,i_l) & " c:"& lista(5,i_l) & " Dato: " & vOutput&"</font>"&vbCrLf
			on error goto 0
			'objShell.popup "Errore Dato Non Scritto:" & scrittura & " s:"&  lista(4,i_l) & " r:"&  lista(7,i_l) & " c:"& lista(5,i_l) & " Dato: " & vOutput, AttesaMessaggio, "errore di scrittura", CRITICAL_ICON + 4096
			Uscita = Uscita & "<tr><td colspan='10'><font color='red'>Output IN ERRORE " &scrittura& " s:"&  lista(4,i_l) & " r:"& lista(7,i_l) & " c:"& lista(6,i_l) & "</font></td><td>  Dato: " & vOutput & "</td></tr>"
		else
			Uscita = Uscita & "<tr><td></td><td><font color='brown'>file destinazione:"&scrittura& " s:"&  lista(4,i_l) & " r:"&  lista(7,i_l) & " c:"& lista(6,i_l) & "</font></td><td>  Dato: " & vOutput & "</td></tr>"
		end if   
		precedente = lettura
	else
		On Error Resume Next
		objStdOut.Write "<font color='red'>Parametro Riscontrato solo nei Sorgenti:" & listaS(0,y)&"</font>"&vbCrLf
		on error goto 0
		'objShell.popup "Parametro Riscontrato solo nei Sorgenti:" & listaS(0,y), AttesaMessaggio,"Attenzione", INFO_ICON + 4096
		Uscita = Uscita & "<tr><td colspan='10'><font color='violet'>Parametro nei sorgenti non usato "  & listaS(0,y) & " sorgente:"&listaS(1,y)&"\"&listaS(2,y)&"</font></td>/tr>"
	end if
Next
Call objWorkbWrite.Save 'salva il file
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore save Destinazione:" & Err.Number & " Description " & Err.Description&"</font>"&vbCrLf
    Wscript.Quit Err.Number
  End If
  Call objWorkbWrite.Close 'chiude il file
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore close Destinazione Description " & Err.Description&"</font>"&vbCrLf
    Wscript.Quit Err.Number
  End If
Uscita = Uscita & "</table></body></html>"
'WScript.Echo "dir:"&CDIR
scriviSu report, Uscita
objShell.run(report) 'Lancia l'eseguibile definito per il tipo di file da leggere.
objExcel.Quit
On Error Resume Next
objStdOut.Write "<font color='blue'>"&NomeProgramma&" terminato</font>"&vbCrLf
Wscript.Quit 0
' ------------------------Funzioni ----------------------------------------------------------------------------------------------

Function riordinaLista()
On Error Resume Next
   Dim i, inv
   inv = true
   while (inv)
        inv = false
        for i = 0 to indice_S-2 step 1
            if ((listaS(1,i)&"\"&listaS(2,i)) > (listaS(1,i+1)&"\"&listaS(2,i+1))) then
                inverti(i)
                inv = true
            end if
        Next
   Wend
If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore in riordinaLista ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
	on error goto 0
	Err.Clear
End If
on error goto 0
End Function

Function inverti(ByVal idx)
On Error Resume Next
Dim temp, i
    for i = 0 to 8
        temp = listaS(i,idx)
        listaS(i,idx) = listaS(i,idx+1)
        listaS(i,idx+1) = temp
    Next
If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore in inverti ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
	on error goto 0
	Err.Clear
End If
on error goto 0
End Function

Function loopSheet(ByVal dir, ByVal file)
On Error Resume Next
	Set objWorkbook = objExcel.Workbooks.Open(dir&"\"&file, False, False)
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore Open Sorgente:"&dir&"\"&file&" Descrizione Err: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
		on error goto 0
	End If
	sheet_n = objWorkbook.Sheets.Count
	'msgBox "File : "&lista(0,z)&" Numero di sheets : " & sheet_n
	for sheet = 1 to sheet_n step 1
		Set objWorksheet = objWorkbook.Worksheets(sheet)
		If (Err.Number <> 0) Then
			On Error Resume Next
			objStdOut.Write "<font color='red'>Errore sheet Sorgente:" & Err.Number & " Descrizione Err: " & Err.Description&"</font>"&vbCrLf
			Err.Clear
			on error goto 0
		End If
		Set objRange = objWorksheet.UsedRange
		If (Err.Number <> 0) Then
			On Error Resume Next
			objStdOut.Write "<font color='red'>Errore set Range:" & Err.Number & " Descrizione Err: " & Err.Description&"</font>"&vbCrLf
			Err.Clear
			on error goto 0
		End If
		FindNamed objWorksheet, sheet, dir&"\"&file
				  'sheet = oggetto, sheet_n numero, file_ori =parziale, f_get = dir+file
	Next 'Loop su tutti gli sheet di un file Input
	on error goto 0
end function

Function leggiINI(ByRef folderOutput, ByRef fileOutput, ByRef repOutput, ByRef debug)
Dim objFileToRead, linea, x, debug_st
On Error Resume Next
debug = false
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(fileINI,1)
	for x = 1 to 4
		linea = objFileToRead.ReadLine()
		if (Instr(1,linea,"cartella=",1) > 0)then
			folderOutput = Mid(linea,10,Len(linea)-9)
		else
			if (Instr(1,linea,"file=",1) > 0)then
				fileOutput = Mid(linea,6,Len(linea)-5)
			else
				if (Instr(1,linea,"rapporto=",1) > 0)then
					repOutput = Mid(linea,10,Len(linea)-9)
				else
					if (Instr(1,linea,"debug=",1) > 0)then
						debug_st = Mid(linea,7,Len(linea)-6)
						if ((Instr(1,Lcase(debug_st),"si",1) > 0) or  (Instr(1,Lcase(debug_st),"yes",1))) then
							debug = true
						else
							debug = false
						end if
					end if
				end if
			end if
		end if
	Next
	objFileToRead.Close
	Set objFileToRead = Nothing
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in leggiINI ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
		on error goto 0
	End If
	on error goto 0
End Function

Function calcolaColonna(ByVal nome_colonna)
	Dim tot, ci, c, i
	On Error Resume Next
	tot = 0
	ci = 0
	if (Len(nome_colonna) > 1) then
		if (Len(nome_colonna) > 2) then
			if (Len(nome_colonna) > 3) then
				objShell.Popup "Considero le colonne solo fino alla terza lettera", AttesaMessaggio, "errore", CRITICAL_ICON + 4096
				calcolaColonna = "+ZZZ"
				Exit Function
			end if
			'WScript.Echo "siamo a 3"
			c = Mid(nome_colonna,1,1)
			ci = CInt(Asc(UCase(c))-64)
			ci = ci * 676
			tot = tot + ci
			'WScript.Echo "1tot:"&tot
			'---------------------------
			c = Mid(nome_colonna,2,1)
			ci = CInt(Asc(UCase(c))-64)
			ci = ci * 26
			tot = tot + ci
			'WScript.Echo "2tot:"&tot
			'---------------------------
			c = Mid(nome_colonna,3,1)
			ci = CInt(Asc(UCase(c))-64)
			tot = tot + ci
			'WScript.Echo "3tot:"&tot
		else
			c = Mid(nome_colonna,1,1)
			ci = CInt(Asc(UCase(c))-64)
			ci = ci * 26
			tot = tot + ci
			'---------------------------
			c = Mid(nome_colonna,2,1)
			ci = CInt(Asc(UCase(c))-64)
			tot = tot + ci
		end If
	Else
		if (nome_colonna = "") then
			'Wscript.Echo "Errore Nome colonna vuoto"
			On Error Resume Next
			objStdOut.Write "<font color='red'>Errore Apertura Nome colonna vuoto"&vbCrLf
			'objShell.popup "Errore Apertura Nome colonna vuoto", AttesaMessaggioL, "errore", CRITICAL_ICON + 4096
			WScript.Quit 1005
		end if
		tot = CInt(Asc(UCase(nome_colonna))-64)
		end if
		calcolaColonna = tot
		If (Err.Number <> 0) Then
			On Error Resume Next
			objStdOut.Write "<font color='red'>Errore in calcolaColonna ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
			Err.Clear
			on error goto 0
		End If
		on error goto 0
End Function

Function cerca(lis,ii,li, ix, ByRef idx_listaS)
Dim i, trovato, ris
On Error Resume Next
'lista(0,y),lista(5,y),lista(6,y),lista(7,y)
	trovato = false
	for i = 0 to ix step 1
		if ((lis(0,ii) = li(0,i)) and (lis(8,ii) = li(1,i))) then
			cerca = true
			idx_listaS = i
			exit function
		end if
	Next
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in cerca ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
		on error goto 0
	End If
	on error goto 0
end function

Sub FindMyComments(ByVal sheet, ByVal sheet_n, ByVal file_ori, ByVal f_get, ByRef lis, ByRef index)
					'sheet = oggetto, sheet_n = numero di sheet, file_ori = file di lettura dei link parte iniziale
					'f_get = file nome completo da cui leggo i link, lis lista su cui registrare
					'index indice nella lista
Dim cmt, f_s, f_c_s, f_link, c_n_s,  r_s, c_s_s
Dim PosI, PosF, PosC, PosFi, PosPar, intermedio, nomePar, cartellaS, cartellaD, temp
'Flink#D#cartellaC:\xxxxxx\yyyyy#fileNomeFile#NomeParametro#Elink tipo FROM = 2 destinazione
'12345678901234567890
On Error Resume Next
For Each cmt In sheet.Comments
	'WScript.Echo "Loop:" & index
	f_link = ""
	PosI = InStr(1,cmt.text,"Flink#",1)
	PosF = InStr(1,cmt.text,"#Elink",1)
	if (PosI > 0) then 'è un link
		SeparaRigheColonne cmt.Parent.Address(0, 0), r_s, c_s_s
		c_n_s = calcolaColonna(c_s_s)
        temp = Mid(cmt.text,PosI+6,1) 'Flink# = 6 caratteri
		if (Mid(cmt.text,PosI+6,1) = "D") then
			PosC = InStr(1,cmt.text,"#cartella=",1)
			PosFi = InStr(1,cmt.text,"#file=",1)
			cartellaD = GetCartella(f_get, f_link) 'estrae il nome della cartella e il nome del file come f_link
			if (PosC > 0) then
				intermedio = Mid(cmt.text,PosC+10,Len(cmt.text))
				cartellaS = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link il nome della cartella S
                ' se si tratta di un riferimento relativo lo risolvo
                if (InStr(1,cartellaS,"..") = 1) then
                    cartellaS = risolviRelativo(cartellaS)
                end if
				if (fso.FolderExists(cartellaS)) then
					if (PosFi > 0) then
						intermedio = Mid(cmt.text,PosFi+6,Len(cmt.text))
                        PosPar = Instr(1,intermedio,"#",1)
                        PosF = Instr(1,intermedio,"#Elink",1)
                        nomePar = Mid(intermedio,PosPar+1,PosF-PosPar-1) 'estrae dal link il nome dell file
						f_s = Mid(intermedio,1,PosPar-1) 'estrae dal link il nome dell file
						f_c_s = cercaFile(f_s, cartellaS , fso)
						if (f_c_s <> "NULLA") then
							lis(9,index) =  f_c_s ' file S
							lis(10,index) = f_s 'chiave_file S
							'WScript.Echo "con file Nome Parametro:"&nomePar
						else
							lis(9,index) =  "" ' file S
							lis(10,index) = "" 'chiave_file S
							'objShell.Popup "Errore " & "Il file " &f_s&" non esiste", AttesaMessaggio, "errore", CRITICAL_ICON + 4096
							On Error Resume Next
							objStdOut.Write "<font color='red'>Errore Il file " &cartellaS&"\"&f_s&" non esiste</font>"&vbCrLf
							on error goto 0
							Uscita = Uscita & "<tr><td colspan=10><font color=red> Errore il file "&cartellaS&"\"&f_s&" non esiste</font></td></tr>"
						end if
                    else
                        lis(9,index) =  "" ' file S
						lis(10,index) = "" 'chiave_file S
                        PosPar = Instr(1,intermedio,"#",1)
					    PosF = Instr(1,intermedio,"#Elink",1)
					    nomePar = Mid(intermedio,PosPar+1,PosF-PosPar-1) 'estrae dal link il nome dell file
						'WScript.Echo "Senza file Nome Parametro:"&nomePar
					end if
					lis(0,index) = nomePar
					lis(1,index) = cartellaD
					lis(2,index) = f_link
					lis(3,index) = file_ori
					lis(4,index) =  sheet_n
					lis(5,index) =  c_n_s
					lis(6,index) =  c_s_s
					lis(7,index) =  r_s
					lis(8,index) =  cartellaS 'cartella S
					index = index+1
				else
					On Error Resume Next
					objStdOut.Write "<font color='red'>Errore DIR-NON-VALIDA: " & cartellaS&"</font>"&vbCrLf
					on error goto 0
					'objShell.Popup "Errore " & "DIR-NON-VALIDA: " & cartellaS, AttesaMessaggio, "errore", CRITICAL_ICON + 4096
					Uscita = Uscita & "<tr><td colspan=10><font color=red> Errore la cartella "&cartellaS&" non esiste</font></td></tr>"
				End If
			else
				'problema non posso cercare in tutto lo spazio files
				On Error Resume Next
				objStdOut.Write "<font color='red'>Non è stata definita la cartella dei file Sorgenti" & vbCrLf&" Solo il nome del file è facoltativo</font>"&vbCrLf
				on error goto 0
				objShell.Popup "Non è stata definita la cartella dei file Sorgenti" & vbCrLf&"Solo il nome del file è facoltativo", AttesaMessaggio, "errore", CRITICAL_ICON + 4096
			end if
		end if
	end if
Next
If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore in FindMyComments ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
	on error goto 0
	Wscript.Quit Err.Number
End If
on error goto 0
End Sub

function risolviRelativo(add)
Dim temp, temp2, Pos, sotto_dir, i
    if (Instr(1,add,"..") > 0) then
        Pos = Instr(1,add,"\")
        if (Pos = 0) then
           Pos = Instr(1,add,"/")
        end if
        temp = Mid(add,Pos+1,Len(add)-Pos)
        sotto_dir = 0
        while (Instr(1,temp,"..") > 0)
            Pos = Instr(1,temp,"\")
            if (Pos = 0) then
                Pos = Instr(1,temp,"/")
            end if
            temp = Mid(temp,Pos+1,Len(temp)-Pos)
            sotto_dir = sotto_dir+1
        Wend
        temp2 = folderDest
        on error goto 0
        for i=0 to sotto_dir
            Pos = InStrRev(temp2,"\")
            if (Pos = 0) then
                Pos = InStrRev(temp2,"/")
            end if
            temp2 = Mid(temp2,1,Pos-1)
        Next
        risolviRelativo = temp2&"\"&temp
    else
        if Not (Instr(1,add,"\") > 0) then
        'contiene solo il nome file la dir è quella del Master
            estraiFolderDaAddress = folderMaster
        end if
    end if

end function

Sub FindNamed(ByVal sheet, ByVal sheet_n, ByVal f_get)
				'sheet = oggetto, sheet_n numero, f_get dir+file, lis = listaS, index indice in listaS
	Dim cmt, f_link, c_n_s, c_s_s, r_s
	Dim PosI, PosF, nomePar
	'Flink#S#nomeParametro#Elink tipo Sorgente
	'12345678901234567890	'lis lista su cui registrare
	On Error Resume Next
	For Each cmt In sheet.Comments	'index indice nella lista
		'WScript.Echo "Loop:" & index	'tipo cerco Link IN = 1 o OUT = 2 o S = 3 o D = 4
    	f_link = ""
    	PosI = InStr(1,cmt.text,"Flink#",1)
    	if (PosI > 0) then
    	    PosF = InStr(1,cmt.text,"#Elink",1)
    		SeparaRigheColonne cmt.Parent.Address(0, 0), r_s, c_s_s
    		c_n_s = calcolaColonna(c_s_s)
    		if (Mid(cmt.text,PosI+6,1) = "S") then
                nomePar = Mid(cmt.text,PosI+8,PosF-(PosI+8))
    			if Not (parametriTrovati(nomePar)) then
				    listaS(0,indice_S) = nomePar
				    listaS(1,indice_S) = GetCartella(f_get, f_link) 'estrae il nome della cartella e il nome del file come f_link
				    listaS(2,indice_S) = f_link
				    listaS(3,indice_S) = ""
				    listaS(4,indice_S) = sheet_n	'sheet S
				    listaS(5,indice_S) = c_n_s 	'colonna numerico S
				    listaS(6,indice_S) = c_s_s 	'colonna lettere S
				    listaS(7,indice_S) = r_s 		'riga S
                    listaS(8,indice_S) = -1
				    indice_S = indice_S+1
                end if
			end if 'se non è un sorgente non mi interessa
		end if
	Next
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in FindNamed ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Sub


Function parametriTrovati(ByVal nomePar)
Dim x
On Error Resume Next
    for x = 0 to indice_S step 1
        if (nomePar = listaS(0,x)) then
            parametriTrovati = true
			on error goto 0
            Exit Function
        end if
    Next
    parametriTrovati = false
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in parametriTrovati ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Function

Function popolaListaFile(ByRef lf, ByVal ind) 'Crea la lista dei file di input per evitare di leggerli due volte
Dim x, y, ce
On Error Resume Next
y = 0
	for x = 0 to ind-1 step 1
		ce = thereis(lista(8,x),lista(9,x),lf,y)
		if Not (ce) then
			lf(y,0) = lista(8,x)
			lf(y,1) = x 
			lf(y,2) = lista(9,x) 'ci mette il nome completo del file anche se è vuoto, serve ad avere un record per ogni file della dir più nel caso uno vuoto
			y = y+1
		end if
	Next
	popolaListaFile = y ' esporto il livello al quale è arrivata la listaFile
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in popolaListaFile ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
end Function

Function thereis(ByVal dir, ByVal file, ByRef lis, ByVal upto)
dim i, ce_dir
On Error Resume Next
	for i = 0 to upto step 1
		if (dir = lis(i,0)) then 'se la dir c'è
			if ( file = lis(i,2)) then 'controllo se anche il file corrisponde
				thereis = true
				on error goto 0
				exit function
			end if
		end if
	Next
	thereis = false
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in thereis ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
end function

Function SeparaRigheColonne(ByRef indirizzo, ByRef riga, ByRef colonna)
Dim c, i
On Error Resume Next
	For i=1 To Len(indirizzo)
		c = Mid(indirizzo,i,1)
		if (IsNumeric(c)) then
			Exit For
		End If
	Next 
	'WScript.Echo "Numerico da " & i
	colonna = Mid(indirizzo, 1, i-1)
	riga = Mid(indirizzo, i, Len(indirizzo))
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in SeparaRigheColonne ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Function

Function GetPuntatore(ByVal nome, ByRef file, ByRef sh, ByRef col, ByRef rig)
	 'C:\Sviluppo_secondaria\VbScript\in#fileAction#foglio2#colonnaAB#riga10
	Dim ss, xx, ff, n_parts, i
	On Error Resume Next
	ss = Split(nome, "\")
	n_parts = UBound(ss)
	'in#fileAction#foglio2#colonnaAB#riga10
	xx = Split(ss(n_parts), "#")
	file = mid(xx(1),5,Len(xx(1))) 'xx(1) = fileNOMEFILE
	sh = mid(xx(2),7,Len(xx(2))) 'xx(2) = foglio2
	col = mid(xx(3),8,Len(xx(3))) 'xx(3) = colonnaAB
	rig = mid(xx(4),5,Len(xx(4))) 'xx(4) = riga10
	For i = 0 to n_parts -1 Step 1
		if (ff = "") then
			ff = ff & ss(i)
		else
			ff = ff & "\" & ss(i)
		end if
	Next
	ff = ff & "\" & xx(0)
	GetPuntatore = ff
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in GetPuntatore ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Function

Function GetCartella(ByVal nome, ByRef file)
	Dim ss, xx, ff, n_parts, i, pre
	On Error Resume Next
	pre = ""
	if Not (InStr(nome,":") > 0) Then
		pre = "\\"
	end if
	ss = Split(nome, "\")
	n_parts = UBound(ss)
	For i = 0 to n_parts -1 Step 1
		if (ff = "") then
			ff = ff & ss(i)
		else
			ff = ff & "\" & ss(i)
		end if
	Next
    ff = pre & ff
	file = ss(n_parts)
	GetCartella = ff
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in GetCartella ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Function

Function ReadExcelFile(obJWK,ByVal sheet, ByVal Row, ByVal Col)
  ' Local variable declarations
  Dim objSheet, objCells
  Dim cellCont
  Dim cc, rr
  On Error Resume Next
  cellCont = "nullo"
  ' Default return value
  ReadExcelFile = Null
  'On Error Resume Next 
  'seleziona lo sheet da usare
  Set objsheet = obJWK.Worksheets(sheet)
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore sel Sheet:" & Err.Number & " Description " & Err.Description&"</font>"&vbCrLf
    WBscript.Quit Err.Number
  End If
  ' Get the used cells
  Set objCells = objSheet.Cells
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore set Cells:" & Err.Number & " Description " & Err.Description&"</font>"&vbCrLf
	WBscript.Quit Err.Number
  End If
  'MsgBox objCells(10, 18).Value
  cellCont = objCells(Row, Col).Value
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore read:" & Err.Number & " Description " & Err.Description&"</font>"&vbCrLf
	WBscript.Quit Err.Number
  End If
  ReadExcelFile = cellCont
End Function

Function WEF(objWorkb,ByVal sheet, ByVal Row, ByVal Col,ByVal valore)
  Dim objSheet, objCells
  WEF = Null
  On Error Resume Next
  'seleziona lo sheet da usare
  Set objSheet = objWorkb.Worksheets(sheet)
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore WEF sheet select:" & Err.Number & " Description " & Err.Description&"</font>"&vbCrLf
    WBscript.Quit Err.Number
  End If
  Set objCells = objSheet.Cells
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore WEF set cell:" & Err.Number & " Description " & Err.Description&"</font>"&vbCrLf
    WBscript.Quit Err.Number
  End If
  objCells(Row,Col).Value = valore ' scrive il valore nella cella
  If (Err.Number <> 0) Then
	On Error Resume Next
	objStdOut.Write "<font color='red'>Errore WEF write cell:" & Err.Number & " Description " & Err.Description&"</font>"&vbCrLf
    WBscript.Quit Err.Number
  End If
  WEF = valore
End Function


Function cercaFile(ByVal patt, ByVal folder, ByVal fso) 'pattern da cercare e directory
Dim filenamecompleto
Dim f, parte
Dim objFolder
On Error Resume Next
	filenamecompleto = "NULLA"
	'WScript.Echo "Folder in cerca:" & folder
	Set objFolder  = fso.GetFolder(folder)
	patt = LCase(patt)
	For Each f In objFolder.Files
		parte = Left(LCase(f.Name),Len(patt)) 'preleva i primi caratteri
		'WScript.Echo patt & " " & parte & " " & LCase(f.Name)
		if InStr(parte,patt) = 0 Then
			'WScript.Echo "---------------- " & LCase(f.Name)
		else
			'WScript.Echo "Trovato " & patt & " " & f.Name
			filenamecompleto = LCase(f.Name)
		End If
    Next
	cercaFile = filenamecompleto
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in cercaFile ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Function

Function scriviSu(ByVal nome, ByVal dato)
Dim objFileToWrite
On Error Resume Next
	'WScript.Echo "File di scrittura:"&nome
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(nome,2,true)
	objFileToWrite.WriteLine(dato)
	objFileToWrite.Close
	Set objFileToWrite = Nothing
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in scriviSu ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Function

Function appendiA(ByVal nome, ByVal dato)
Dim objFileToWrite
On Error Resume Next
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(nome,8,true)
	objFileToWrite.WriteLine(dato)
	objFileToWrite.Close
	Set objFileToWrite = Nothing
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in appendiA ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
End Function

Function listaDir(dir)
Dim objDir, objFile, colFiles, estens
On Error Resume Next
	Set objDir = fso.GetFolder(dir)
	'Wscript.Echo objDir.Path
	Set colFiles = objDir.Files
	For Each objFile in colFiles
		'WBscript.Echo "estensione:" & objFSO.GetExtensionName(objFile.name)
        estens = fso.GetExtensionName(objFile.name)
		If (UCase(Mid(estens,1,3)) = "XLS") Then
            if (Instr(1,objFile.name,"~$") = 0) then
			    'Wscript.Echo "Cerco Parametri nel file:"& objFile.Name
			    loopSheet dir, objFile.Name
            end if
		End If
	Next
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in listaDir ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
end function

Function esportaListe(ByVal modo)
On Error Resume Next
    if (Fdebug) then
	    out = "Lista" & vbCrLf
	    for y = 0 to indice-1 step 1
		    out = out & "Parametro			   :" & lista(0,y) & vbCrLf
		    out = out & "directory			(D):" & lista(1,y) & vbCrLf
		    out = out & "nome file			(D):" & lista(2,y) & vbCrLf
		    out = out & "chiave_file		(D):" & lista(3,y) & vbCrLf
		    out = out & "sheet				(D):" & lista(4,y) & vbCrLf
		    out = out & "colonna n.			(D):" & lista(5,y) & vbCrLf
		    out = out & "colonna lettrere	(D):" & lista(6,y) & vbCrLf
		    out = out & "riga				(D):" & lista(7,y) & vbCrLf
		    out = out & "Directory			(S):" & lista(8,y) & vbCrLf
		    out = out & "file				(S):" & lista(9,y) & vbCrLf
		    out = out & "chiave_file		(S):" & lista(10,y)& vbCrLf
		    out = out & "trovato               :" & lista(11,y)& vbCrLf
	    Next
	    out = out & "Lista_I" & vbCrLf
	    for y = 0 to indice_S-1 step 1
		    out = out & "Parametro			   :" &listaS(0,y)& vbCrLf
		    out = out & "directory			(D):" &listaS(1,y)& vbCrLf
		    out = out & "nome file			(D):" &listaS(2,y)& vbCrLf
		    out = out & "chiave_file		(D):" &listaS(3,y)& vbCrLf
		    out = out & "sheet				(D):" &listaS(4,y)& vbCrLf
		    out = out & "colonna n.			(D):" &listaS(5,y)& vbCrLf
		    out = out & "colonna lettrere	(D):" &listaS(6,y)& vbCrLf
		    out = out & "riga				(D):" &listaS(7,y)& vbCrLf
		    out = out & "trovato			   :" &listaS(8,y)& vbCrLf
	    Next
        if (modo = "A") then
            appendiA CDir & "\liste.txt", out
        else
	        scriviSu CDir & "\liste.txt", out
        end if
    end if
	If (Err.Number <> 0) Then
		On Error Resume Next
		objStdOut.Write "<font color='red'>Errore in esportaListe ErrN." & Err.Number & " Descrizione: " & Err.Description&"</font>"&vbCrLf
		Err.Clear
	End If
	on error goto 0
end function
