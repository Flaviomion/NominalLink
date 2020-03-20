Option Explicit On

Imports System
Imports System.IO

Module NominalLink

    Sub Main()
        Const Versione = "1.6.1"
        Dim dati_versione, myPass

        Const Name = "NominalLink"
        Dim Fdebug
        Dim sheet, objExcel, objWorkbook, objWorksheet, objRange, objShell, out, report, objStdOut
        Dim CDir, indice, indice_S, y, z, vInput, vOutput, Uscita, vTest, file_output, sheet_n, n_fileIn
        Dim ris, idx, scritto, modificato, bReadOnly
        Dim folderDest, fileDest, file_dest, fileINI, NomeProgramma
        Dim errore
        Dim descrizione
        Fdebug = False
        fileINI = "NominalLink.ini"
        NomeProgramma = Name & "_" & Versione

        dati_versione = "Programma di valorizzazione parametri attraverso un link Nominale con variabili su altri file Excel"
        dati_versione = dati_versione & Chr(10) & Chr(13) & " 1.6.1 Inserimento gestione Password per Fogli Bloccati"

        Dim ofile As TextWriter = File.CreateText("Output.txt")
        Dim oOut As TextWriter = Console.Out
        Dim width = Console.WindowWidth
        Dim heigth = Console.WindowHeight
        Console.SetWindowSize(width * 2, heigth * 2)
        'Console.Read()
        WriteMia(ConsoleColor.White, "Partenza " & NomeProgramma, oOut, ofile)

        Dim objArgs() As String = Environment.GetCommandLineArgs()
        Console.BackgroundColor = ConsoleColor.Black
        If (objArgs.Count > 1) Then
            If (StrComp(objArgs(1), "versione") = 0) Then
                versioneDisp(ofile, oOut, Versione, dati_versione)
            Else
                myPass = objArgs(1)
            End If
        Else
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.Write("Dammi la Password:")
            Console.ForegroundColor = ConsoleColor.Black
            myPass = Console.ReadLine()
            If (myPass = "versione") Then
                versioneDisp(ofile, oOut, Versione, dati_versione)
            End If
        End If

        'solo per test
        'myPass = "segreto"


        'folderDest = WScript.Arguments.Item(0)
        'fileDest = WScript.Arguments.Item(1)
        Dim lista(12, 100)       ' 0	nome del Parametro
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

        Dim listaS(9, 100)       ' 0	nome del Parametro	(S)
        ' 1	directory			(S)
        ' 2	nome file			(S)
        ' 3	chiave_file			(S)
        ' 4 	sheet				(S)
        ' 5 	colonna n.			(S)
        ' 6 	colonna lettrere	(S)
        ' 7 	riga				(S)
        ' 8	trovato

        Dim listaFile(100, 3)   '0	dir
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

        Dim Started
        objShell = CreateObject("Wscript.Shell")
        objExcel = CreateObject("Excel.Application")
        Dim fso : fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo 0
        leggiINI(folderDest, fileDest, report, Fdebug, fileINI, ofile, oOut)
        Dim folderMaster = folderDest

        CDir = fso.GetAbsolutePathName(".")
        WriteMia(ConsoleColor.Cyan, "Start " & NomeProgramma, oOut, ofile)

        objExcel.DisplayAlerts = 0
        'objExcel.Visible = False

        file_dest = cercaFile(fileDest, folderDest, fso, ofile, oOut)
        WriteMia(ConsoleColor.Cyan, "Elaboro il file:" & folderDest & "\" & file_dest, oOut, ofile)

        On Error Resume Next
        objWorkbook = objExcel.Workbooks.Open(folderDest & "\" & file_dest, False, False)
        If (Err.Number <> 0) Then
            errore = Err.Number
            descrizione = Err.Description
            WriteMia(ConsoleColor.Red, "Errore Open file Destinazione:" & folderDest & "\" & file_dest & " Descrizione Err: " & descrizione, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If

        bReadOnly = objWorkbook.ReadOnly
        If bReadOnly = True Then
            WriteMia(ConsoleColor.Red, "Errore apertura file " & folderDest & "\" & file_dest & " File OCCUPATO", oOut, ofile)
            Call objWorkbook.Close
            objExcel.Quit
            WriteMia(ConsoleColor.Magenta, "Programma Terminato a causa di file Occupato", oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If

        settaPassword(objWorkbook, myPass, ofile, oOut)

        sheet_n = objWorkbook.Sheets.Count
        'msgBox "Numero di sheets : " & sheet_n
        indice = 0 'indice viene incrementato da FindMyComments
        '-- 1) del flow --- Legge da Input tutti i Flink# su tutti gli sheet
        'Flink#D#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#parametro#Elink
        'ricerca dei link nelfile di destinazione 
        For sheet = 1 To sheet_n Step 1
            On Error Resume Next
            objWorksheet = objWorkbook.Worksheets(sheet)
            If (Err.Number <> 0) Then
                errore = Err.Number
                descrizione = Err.Description
                WriteMia(ConsoleColor.Red, "Errore sheet file Destinazione, Descrizione: " & descrizione, oOut, ofile)
                ofile.Close()
                Console.Read()
                End
            End If
            On Error GoTo 0
            objRange = objWorksheet.UsedRange
            WriteMia(ConsoleColor.Yellow, "Ricerca Flink su sheet: " & objWorksheet.name, oOut, ofile)
            FindMyComments(objWorksheet, sheet, fileDest, folderDest & "\" & file_dest, lista, indice, fso, folderDest, folderMaster, Uscita, ofile, oOut)
        Next
        Uscita = "<html><body><table><tr><td colspan=10><h2><font color='blue'>" & NomeProgramma & "</font></h2></td></tr>"
        Uscita = Uscita & "<tr><td colspan=10><font color ='DarkGreen'>Gestisce link a campi Nominali</font></td></tr>"
        Uscita = Uscita & "<tr><td  colspan='10' align='center'><font color ='DarkBlue'>" & folderDest & "\" & file_dest & "</font></td></tr>"
        Uscita = Uscita & "<tr><td colspan=10><font color ='blue'>Rapporto del " & Day(Now()) & " ore:" & Now() & "</font></td></tr>"
        objExcel.ActiveWorkbook.Close(False, folderDest & "\" & file_dest) 'non mi accetta come Workbook objWorkbook

        n_fileIn = popolaListaFile(listaFile, indice, lista, ofile, oOut)

        '-- 2)--- usa listaFile per cercare tutti i link Nominali che hanno un file di riferimento
        indice_S = 0 'indice_S viene incrementato da FindMyComments
        WriteMia(ConsoleColor.Yellow, "Ricerca Parametri sui file Sorgenti", oOut, ofile)
        For z = 0 To n_fileIn - 1 Step 1
            If Not (listaFile(z, 2) = "") Then
                'ho il file
                loopSheet(listaFile(z, 0), listaFile(z, 2), sheet_n, objWorkbook, objWorksheet, objRange, objExcel, myPass, listaS, indice_S, ofile, oOut)
            Else
                'ricerca su tutti gli excel della sourceDir
                listaDir(listaFile(z, 0), sheet_n, objWorkbook, objWorksheet, objRange, objExcel, myPass, fso, listaS, indice_S, ofile, oOut)
            End If
            objExcel.ActiveWorkbook.Close(False, lista(z, 0))
        Next ' (utilizzando la listaFile come indice per andare a prendere il giusto file sulla lista lista)

        '-- 3)--- Cerca su listaS i link corrispondenti a quelli riscontrati su lista 
        WriteMia(ConsoleColor.Yellow, "Crea la corrispondenza fra Destinazione e Sorgente", oOut, ofile)
        idx = 0
        For y = 0 To indice - 1 Step 1
            ris = cerca(lista, y, listaS, indice_S, idx, ofile, oOut) 'cerca la corrispondena di Parametro D e S
            If (ris) Then
                'WScript.Echo "TROVATO link " & lista(0,y) & " " & lista(1,y) & " " & lista(2,y) & " " & lista(3,y) & " " & lista(6,y) & " " & lista(7,y) & " " & lista(8,y) & " " & lista(9,y)
                lista(11, y) = idx   'metto su lista il link a listaS
                listaS(8, idx) = y   'metto su listaS il link a lista
            Else
                WriteMia(ConsoleColor.Cyan, "Parametro non trovato nei Sorgenti:" & lista(0, y), oOut, ofile)
                Uscita = Uscita & "<tr><td colspan='10'><font color='Coral'>Parametro non riscontrato nei file sorgenti </font><font color='blue'>" & lista(0, y) & "</font></td></tr>"
                lista(11, y) = -1  'metto su lista il link a listaS
                'listaS(8,idx) = -1 il link negativo non ha senso sulla lista dei Sorgenti
            End If
        Next

        esportaListe("S", out, lista, indice, indice_S, listaS, CDir, Fdebug, ofile, oOut) 'scrive
        WriteMia(ConsoleColor.Yellow, "Ordina la lista per ottimizzare l'accesso ai files", oOut, ofile)
        riordinaLista(indice_S, listaS, ofile, oOut)

        esportaListe("A", out, lista, indice, indice_S, listaS, CDir, Fdebug, ofile, oOut) 'appende
        WriteMia(ConsoleColor.Yellow, "Esegue la copia dei dati Sorgenti su Destinatari", oOut, ofile)



        '------ Legge tutti i link da lista ed esegue la migrazione dati 5) del flow
        Dim i_l, precedente, lettura, scrittura, obJWK, objWorkbWrite
        precedente = "primo"
        scrittura = lista(1, 0) & "\" & lista(2, 0)
        On Error Resume Next
        objWorkbWrite = objExcel.Workbooks.Open(scrittura, 0, False)
        If (Err.Number <> 0) Then
            descrizione = Err.Description
            WriteMia(ConsoleColor.Red, "Errore Apertura Destinazione Description " & descrizione, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If
        On Error GoTo 0
        settaPassword(objWorkbWrite, myPass, ofile, oOut)
        For y = 0 To indice_S - 1 Step 1
            i_l = listaS(8, y)
            If (i_l <> -1) Then
                lettura = listaS(1, y) & "\" & listaS(2, y)
                If Not (precedente = lettura) Then
                    ' chiudo il vecchio file di lettura
                    If Not (precedente = "primo") Then
                        'Set obJWK = ActiveWorkbook
                        Call obJWK.Close(False, precedente)
                    End If
                    'apro un nuovo file di lettura
                    On Error Resume Next
                    obJWK = objExcel.Workbooks.Open(lettura, False, True)
                    If (Err.Number <> 0) Then
                        descrizione = Err.Description
                        WriteMia(ConsoleColor.Red, "Errore Apertura Sorgente Description " & descrizione, oOut, ofile)
                        ofile.Close()
                        Console.Read()
                        End
                    End If
                    On Error GoTo 0
                    settaPassword(obJWK, myPass, ofile, oOut)
                End If
                vInput = ReadExcelFile(obJWK, listaS(4, y), listaS(7, y), listaS(5, y), ofile, oOut) 'Workbook,sInput,rInput,cInput,objExcel
                WriteMia(ConsoleColor.Cyan, "Leggo: chiave:" & listaS(0, y) & " file:" & lettura & " foglio:" & listaS(4, y) & " righa:" & listaS(7, y) & " colonna:" & listaS(5, y), oOut, ofile)
                On Error Resume Next
                Uscita = Uscita & "<tr><td><font color='green'>Parametro: " & listaS(0, y) & "</font></td><td><font color='green'>file sorgente:" & lettura & " s:" & listaS(4, y) & " r:" & listaS(7, y) & " c:" & listaS(6, y) & "</font></td><td> Dato: " & vInput & "</td></tr>"
                If (Err.Number <> 0) Then
                    errore = Err.Number
                    descrizione = Err.Description
                    WriteMia(ConsoleColor.Red, "Errore nel Dato:" & listaS(0, y) & " Errore:" & errore & " Description " & descrizione, oOut, ofile)
                    Uscita = Uscita & "<tr><td colspan='10'><font color='red'>Errore nel Dato:" & listaS(0, y) & " " & lettura & " " & " foglio:" & listaS(4, y) & " righa:" & listaS(7, y) & " colonna:" & listaS(5, y) & "</td></tr>"
                    Err.Clear()
                Else
                    On Error GoTo 0
                    vOutput = WEF(objWorkbWrite, lista(4, i_l), lista(7, i_l), lista(5, i_l), vInput, ofile, oOut) 'file,sh,rOutput,cOutput,valore
                    vTest = ReadExcelFile(objWorkbWrite, lista(4, i_l), lista(7, i_l), lista(5, i_l), ofile, oOut)
                    If (vOutput <> vTest) Then
                        WriteMia(ConsoleColor.Red, "Errore Dato Non Scritto:" & scrittura & " s:" & lista(4, i_l) & " r:" & lista(7, i_l) & " c:" & lista(5, i_l) & " Dato: " & vOutput, oOut, ofile)
                        Uscita = Uscita & "<tr><td colspan='10'><font color='red'>Output IN ERRORE " & scrittura & " s:" & lista(4, i_l) & " r:" & lista(7, i_l) & " c:" & lista(6, i_l) & "</font></td><td>  Dato: " & vOutput & "</td></tr>"
                    Else
                        Uscita = Uscita & "<tr><td></td><td><font color='brown'>file destinazione:" & scrittura & " s:" & lista(4, i_l) & " r:" & lista(7, i_l) & " c:" & lista(6, i_l) & "</font></td><td>  Dato: " & vOutput & "</td></tr>"
                    End If
                    precedente = lettura
                End If
            Else
                WriteMia(ConsoleColor.Yellow, "Parametro Riscontrato solo nei Sorgenti:" & listaS(0, y), oOut, ofile)
                Uscita = Uscita & "<tr><td colspan='10'><font color='violet'>Parametro Riscontrato solo nei Sorgenti: </font><font color='blue'>" & listaS(0, y) & " sorgente:" & listaS(1, y) & "\" & listaS(2, y) & "</font></td>/tr>"
            End If
        Next
        On Error Resume Next
        Call objWorkbWrite.Save 'salva il file
        If (Err.Number <> 0) Then
            errore = Err.Number
            descrizione = Err.Description
            WriteMia(ConsoleColor.White, "Errore save Destinazione:" & errore & " Description " & descrizione, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If
        Call objWorkbWrite.Close 'chiude il file
        If (Err.Number <> 0) Then
            errore = Err.Number
            descrizione = Err.Description
            WriteMia(ConsoleColor.Red, "Errore close Destinazione Description " & descrizione, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If
        On Error GoTo 0
        Uscita = Uscita & "</table></body></html>"
        'WScript.Echo "dir:"&CDIR
        scriviSu(report, Uscita, ofile, oOut)
        objShell.run(report) 'Lancia l'eseguibile definito per il tipo di file da leggere.
        objExcel.Quit
        WriteMia(ConsoleColor.Yellow, NomeProgramma & " terminato", oOut, ofile)
        ofile.Close()
        Console.Read()
        End

    End Sub
    ' ------------------------Funzioni ----------------------------------------------------------------------------------------------

    Function WriteMia(colore, messaggio, ByRef oOut, ByRef ofile)
        Console.ForegroundColor = colore
        Console.SetOut(oOut)
        Console.WriteLine(messaggio)
        Console.SetOut(ofile)
        Console.WriteLine(messaggio)
        Console.SetOut(oOut)
        WriteMia = 0
    End Function

    Function versioneDisp(ByRef ofile, ByRef oOut, Versione, dati_versione)
        WriteMia(ConsoleColor.Cyan, Versione, oOut, ofile)
        WriteMia(ConsoleColor.Cyan, dati_versione, oOut, ofile)
        Dim dummy = Console.ReadLine()
        ofile.Close()
        End
    End Function

    Function settaPassword(WK, myPass, ByRef ofile, ByRef oOut)
        Dim wsp
        If (StrComp(myPass, "") <> 0) Then
            For Each wsp In WK.Worksheets
                On Error Resume Next
                If (wsp.ProtectContents) Then
                    'objStdOut.Write "<font color ='Purple'>"&WK.Name&" Foglio:"&wsp.Name&" è Protetto per contenuto</font>"&vbCrLf
                    'objStdOut.Write "<font color ='Purple'>Setto la Password pass:"&myPass&" su "&WK.Name&" Foglio:"&wsp.Name&"</font>"&vbCrLf
                    wsp.Protect(myPass, "True", "True", "True", "True")
                    If (Err.Number = 1004) Then
                        WriteMia(ConsoleColor.Magenta, "Password differente su " & WK.Name & " Foglio:" & wsp.Name, oOut, ofile)
                        Err.Clear()
                    End If
                Else
                    'objStdOut.Write "<font color ='Purple'>"&WK.Name&" Foglio:"&wsp.Name&" Non Protetto</font>"&vbCrLf
                    Err.Clear()
                    'if ( wsp.ProtectDrawingObjects ) then
                    '	objStdOut.Write "<font color ='Purple'>"&WK.Name&" Foglio:"&wsp.Name&" è Protetto per Drawing</font>"&vbCrLf
                    'end if
                End If
                'if ( wsp.ProtectScenarios ) then
                'objStdOut.Write "<font color ='Purple'>"&WK.Name&" Foglio:"&wsp.Name&" è Protetto per Scenario</font>"&vbCrLf
                'end if
                If (wsp.ProtectionMode) Then
                    WriteMia(ConsoleColor.Magenta, WK.Name & " Foglio:" & wsp.Name & " Accesso consentito allo script", oOut, ofile)
                    Err.Clear()
                End If
                On Error GoTo 0
            Next
        Else
            On Error Resume Next
            WriteMia(ConsoleColor.Green, "Senza password", oOut, ofile)
            On Error GoTo 0
        End If
        'The three protection properties of a worksheet are the following:
        '  Sheets(1).ProtectContents
        '  Sheets(1).ProtectDrawingObjects
        '  Sheets(1).ProtectScenarios
        'You can check whether both 3 are False. If this is the case, it is not protected.
        '.ProtectionMode
    End Function

    Function riordinaLista(indice_S, listaS, ByRef ofile, ByRef oOut)
        On Error Resume Next
        Dim i, inv
        inv = True
        While (inv)
            inv = False
            For i = 0 To indice_S - 2 Step 1
                If ((listaS(1, i) & "\" & listaS(2, i)) > (listaS(1, i + 1) & "\" & listaS(2, i + 1))) Then
                    inverti(i, listaS, ofile, oOut)
                    inv = True
                End If
            Next
        End While
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in riordinaLista ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function inverti(ByVal idx, listaS, ByRef ofile, ByRef oOut)
        On Error Resume Next
        Dim temp, i
        For i = 0 To 8
            temp = listaS(i, idx)
            listaS(i, idx) = listaS(i, idx + 1)
            listaS(i, idx + 1) = temp
        Next
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in inverti ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function loopSheet(ByVal dir, ByVal file, sheet_n, ByRef objWorkbook, ByRef objWorksheet, ByRef objRange, objExcel, myPass, ByRef listaS, ByRef indice_S, ByRef ofile, ByRef oOut)
        On Error Resume Next
        objWorkbook = objExcel.Workbooks.Open(dir & "\" & file, False, False)
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore Open Sorgente:" & dir & "\" & file & " Descrizione Err: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        settaPassword(objWorkbook, myPass, oOut, ofile)
        sheet_n = objWorkbook.Sheets.Count
        'msgBox "File : "&lista(0,z)&" Numero di sheets : " & sheet_n
        For sheet = 1 To sheet_n Step 1
            objWorksheet = objWorkbook.Worksheets(sheet)
            If (Err.Number <> 0) Then
                WriteMia(ConsoleColor.Red, "Errore sheet Sorgente:" & Err.Number & " Descrizione Err: " & Err.Description, oOut, ofile)
                Err.Clear()
            End If
            objRange = objWorksheet.UsedRange
            If (Err.Number <> 0) Then
                WriteMia(ConsoleColor.Red, "Errore set Range:" & Err.Number & " Descrizione Err: " & Err.Description, oOut, ofile)
                Err.Clear()
            End If
            FindNamed(objWorksheet, sheet, dir & "\" & file, listaS, indice_S, ofile, oOut)
            'sheet = oggetto, sheet_n numero, file_ori =parziale, f_get = dir+file
        Next 'Loop su tutti gli sheet di un file Input
        On Error GoTo 0
    End Function

    Function leggiINI(ByRef folderOutput, ByRef fileOutput, ByRef repOutput, ByRef debug, ByRef fileINI, ByRef ofile, ByRef oOut)
        Dim objFileToRead, linea, x, debug_st
        On Error Resume Next
        debug = False
        objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(fileINI, 1)
        For x = 1 To 4
            linea = objFileToRead.ReadLine()
            If (InStr(1, linea, "cartella=", 1) > 0) Then
                folderOutput = Mid(linea, 10, Len(linea) - 9)
            Else
                If (InStr(1, linea, "file=", 1) > 0) Then
                    fileOutput = Mid(linea, 6, Len(linea) - 5)
                Else
                    If (InStr(1, linea, "rapporto=", 1) > 0) Then
                        repOutput = Mid(linea, 10, Len(linea) - 9)
                    Else
                        If (InStr(1, linea, "debug=", 1) > 0) Then
                            debug_st = Mid(linea, 7, Len(linea) - 6)
                            If ((InStr(1, LCase(debug_st), "si", 1) > 0) Or (InStr(1, LCase(debug_st), "yes", 1))) Then
                                debug = True
                            Else
                                debug = False
                            End If
                        End If
                    End If
                End If
            End If
        Next
        objFileToRead.Close
        objFileToRead = Nothing
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in leggiINI ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function calcolaColonna(ByVal nome_colonna, ByRef ofile, ByRef oOut)
        Dim tot, ci, c, i
        On Error Resume Next
        tot = 0
        ci = 0
        If (Len(nome_colonna) > 1) Then
            If (Len(nome_colonna) > 2) Then
                If (Len(nome_colonna) > 3) Then
                    WriteMia(ConsoleColor.Magenta, "Considero le colonne solo fino alla terza lettera", oOut, ofile)
                    calcolaColonna = "+ZZZ"
                    Exit Function
                End If
                'WScript.Echo "siamo a 3"
                c = Mid(nome_colonna, 1, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                ci = ci * 676
                tot = tot + ci
                'WScript.Echo "1tot:"&tot
                '---------------------------
                c = Mid(nome_colonna, 2, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                ci = ci * 26
                tot = tot + ci
                'WScript.Echo "2tot:"&tot
                '---------------------------
                c = Mid(nome_colonna, 3, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                tot = tot + ci
                'WScript.Echo "3tot:"&tot
            Else
                c = Mid(nome_colonna, 1, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                ci = ci * 26
                tot = tot + ci
                '---------------------------
                c = Mid(nome_colonna, 2, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                tot = tot + ci
            End If
        Else
            If (nome_colonna = "") Then
                'Wscript.Echo "Errore Nome colonna vuoto"
                WriteMia(ConsoleColor.Red, "Errore Apertura Nome colonna vuoto", oOut, ofile)
                ofile.Close()
                Console.Read()
                End
            End If
            tot = CInt(Asc(UCase(nome_colonna)) - 64)
        End If
        calcolaColonna = tot
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in calcolaColonna ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function cerca(lis, ii, li, ix, ByRef idx_listaS, ByRef ofile, ByRef oOut)
        Dim i, trovato, ris
        On Error Resume Next
        'lista(0,y),lista(5,y),lista(6,y),lista(7,y)
        trovato = False
        For i = 0 To ix Step 1
            If ((lis(0, ii) = li(0, i)) And (lis(8, ii) = li(1, i))) Then
                cerca = True
                idx_listaS = i
                Exit Function
            End If
        Next
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in cerca ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Sub FindMyComments(ByVal sheet, ByVal sheet_n, ByVal file_ori, ByVal f_get, ByRef lis, ByRef index, ByRef fso, ByRef folderDest, ByRef folderMaster, ByRef Uscita, ByRef ofile, ByRef oOut)
        'sheet = oggetto, sheet_n = numero di sheet, file_ori = file di lettura dei link parte iniziale
        'f_get = file nome completo da cui leggo i link, lis lista su cui registrare
        'index indice nella lista
        Dim cmt, f_s, f_c_s, f_link, c_n_s, r_s, c_s_s
        Dim PosI, PosF, PosC, PosFi, PosPar, intermedio, nomePar, cartellaS, cartellaD, temp
        'Flink#D#cartellaC:\xxxxxx\yyyyy#fileNomeFile#NomeParametro#Elink tipo FROM = 2 destinazione
        '12345678901234567890
        On Error Resume Next
        For Each cmt In sheet.Comments
            'WScript.Echo "Loop:" & index
            f_link = ""
            PosI = InStr(1, cmt.text, "Flink#", 1)
            PosF = InStr(1, cmt.text, "#Elink", 1)
            If (PosI > 0) Then 'è un link
                SeparaRigheColonne(cmt.Parent.Address(0, 0), r_s, c_s_s, ofile, oOut)
                c_n_s = calcolaColonna(c_s_s, ofile, oOut)
                temp = Mid(cmt.text, PosI + 6, 1) 'Flink# = 6 caratteri
                If (Mid(cmt.text, PosI + 6, 1) = "D") Then
                    PosC = InStr(1, cmt.text, "#cartella=", 1)
                    PosFi = InStr(1, cmt.text, "#file=", 1)
                    cartellaD = GetCartella(f_get, f_link, ofile, oOut) 'estrae il nome della cartella e il nome del file come f_link
                    If (PosC > 0) Then
                        intermedio = Mid(cmt.text, PosC + 10, Len(cmt.text))
                        cartellaS = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link il nome della cartella S
                        ' se si tratta di un riferimento relativo lo risolvo
                        If (InStr(1, cartellaS, "..") = 1) Then
                            cartellaS = risolviRelativo(cartellaS, folderDest, folderMaster)
                        End If
                        If (fso.FolderExists(cartellaS)) Then
                            If (PosFi > 0) Then
                                intermedio = Mid(cmt.text, PosFi + 6, Len(cmt.text))
                                PosPar = InStr(1, intermedio, "#", 1)
                                PosF = InStr(1, intermedio, "#Elink", 1)
                                nomePar = Mid(intermedio, PosPar + 1, PosF - PosPar - 1) 'estrae dal link il nome dell file
                                f_s = Mid(intermedio, 1, PosPar - 1) 'estrae dal link il nome dell file
                                f_c_s = cercaFile(f_s, cartellaS, fso, ofile, oOut)
                                If (f_c_s <> "NULLA") Then
                                    lis(9, index) = f_c_s ' file S
                                    lis(10, index) = f_s 'chiave_file S
                                    'WScript.Echo "con file Nome Parametro:"&nomePar
                                Else
                                    lis(9, index) = "" ' file S
                                    lis(10, index) = "" 'chiave_file S
                                    WriteMia(ConsoleColor.Red, "Errore Il file " & cartellaS & "\" & f_s & " non esiste</font>", oOut, ofile)
                                    Uscita = Uscita & "<tr><td colspan=10><font color=red> Errore il file " & cartellaS & "\" & f_s & " non esiste</font></td></tr>"
                                End If
                            Else
                                lis(9, index) = "" ' file S
                                lis(10, index) = "" 'chiave_file S
                                PosPar = InStr(1, intermedio, "#", 1)
                                PosF = InStr(1, intermedio, "#Elink", 1)
                                nomePar = Mid(intermedio, PosPar + 1, PosF - PosPar - 1) 'estrae dal link il nome dell file
                                'WScript.Echo "Senza file Nome Parametro:"&nomePar
                            End If
                            lis(0, index) = nomePar
                            lis(1, index) = cartellaD
                            lis(2, index) = f_link
                            lis(3, index) = file_ori
                            lis(4, index) = sheet_n
                            lis(5, index) = c_n_s
                            lis(6, index) = c_s_s
                            lis(7, index) = r_s
                            lis(8, index) = cartellaS 'cartella S
                            index = index + 1
                        Else
                            WriteMia(ConsoleColor.Red, "Errore DIR-NON-VALIDA: " & cartellaS, oOut, ofile)
                            Uscita = Uscita & "<tr><td colspan=10><font color=red> Errore la cartella " & cartellaS & " non esiste</font></td></tr>"
                        End If
                    Else
                        'problema non posso cercare in tutto lo spazio files
                        WriteMia(ConsoleColor.Red, "Non è stata definita la cartella dei file Sorgenti" & vbCrLf & " Solo il nome del file è facoltativo", oOut, ofile)
                        WriteMia(ConsoleColor.Red, "Non è stata definita la cartella dei file Sorgenti" & vbCrLf & "Solo il nome del file è facoltativo", oOut, ofile)
                    End If
                End If
            End If
        Next
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in FindMyComments ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If
        On Error GoTo 0
    End Sub

    Function risolviRelativo(add, ByRef folderDest, ByRef folderMaster)
        Dim temp, temp2, Pos, sotto_dir, i
        If (InStr(1, add, "..") > 0) Then
            Pos = InStr(1, add, "\")
            If (Pos = 0) Then
                Pos = InStr(1, add, "/")
            End If
            temp = Mid(add, Pos + 1, Len(add) - Pos)
            sotto_dir = 0
            While (InStr(1, temp, "..") > 0)
                Pos = InStr(1, temp, "\")
                If (Pos = 0) Then
                    Pos = InStr(1, temp, "/")
                End If
                temp = Mid(temp, Pos + 1, Len(temp) - Pos)
                sotto_dir = sotto_dir + 1
            End While
            temp2 = folderDest
            On Error GoTo 0
            For i = 0 To sotto_dir
                Pos = InStrRev(temp2, "\")
                If (Pos = 0) Then
                    Pos = InStrRev(temp2, "/")
                End If
                temp2 = Mid(temp2, 1, Pos - 1)
            Next
            risolviRelativo = temp2 & "\" & temp
        Else
            If Not (InStr(1, add, "\") > 0) Then
                'contiene solo il nome file la dir è quella del Master
                risolviRelativo = folderMaster
            End If
        End If

    End Function

    Sub FindNamed(ByVal sheet, ByVal sheet_n, ByVal f_get, ByRef listaS, ByRef indice_S, ByRef ofile, ByRef oOut)
        'sheet = oggetto, sheet_n numero, f_get dir+file, lis = listaS, index indice in listaS
        Dim cmt, f_link, c_n_s, c_s_s, r_s
        Dim PosI, PosF, nomePar
        'Flink#S#nomeParametro#Elink tipo Sorgente
        '12345678901234567890	'lis lista su cui registrare
        On Error Resume Next
        For Each cmt In sheet.Comments  'index indice nella lista
            'WScript.Echo "Loop:" & index	'tipo cerco Link IN = 1 o OUT = 2 o S = 3 o D = 4
            f_link = ""
            PosI = InStr(1, cmt.text, "Flink#", 1)
            If (PosI > 0) Then
                PosF = InStr(1, cmt.text, "#Elink", 1)
                SeparaRigheColonne(cmt.Parent.Address(0, 0), r_s, c_s_s, ofile, oOut)
                c_n_s = calcolaColonna(c_s_s, ofile, oOut)
                If (Mid(cmt.text, PosI + 6, 1) = "S") Then
                    nomePar = Mid(cmt.text, PosI + 8, PosF - (PosI + 8))
                    If Not (parametriTrovati(nomePar, indice_S, listaS, ofile, oOut)) Then
                        listaS(0, indice_S) = nomePar
                        listaS(1, indice_S) = GetCartella(f_get, f_link, ofile, oOut) 'estrae il nome della cartella e il nome del file come f_link
                        listaS(2, indice_S) = f_link
                        listaS(3, indice_S) = ""
                        listaS(4, indice_S) = sheet_n   'sheet S
                        listaS(5, indice_S) = c_n_s     'colonna numerico S
                        listaS(6, indice_S) = c_s_s     'colonna lettere S
                        listaS(7, indice_S) = r_s       'riga S
                        listaS(8, indice_S) = -1
                        indice_S = indice_S + 1
                    End If
                End If 'se non è un sorgente non mi interessa
            End If
        Next
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in FindNamed ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Sub


    Function parametriTrovati(ByVal nomePar, ByRef indice_S, ByRef listaS, ByRef ofile, ByRef oOut)
        Dim x
        On Error Resume Next
        For x = 0 To indice_S Step 1
            If (nomePar = listaS(0, x)) Then
                parametriTrovati = True
                On Error GoTo 0
                Exit Function
            End If
        Next
        parametriTrovati = False
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in parametriTrovati ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function popolaListaFile(ByRef lf, ByVal ind, ByRef lista, ByRef ofile, ByRef oOut) 'Crea la lista dei file di input per evitare di leggerli due volte
        Dim x, y, ce
        On Error Resume Next
        y = 0
        For x = 0 To ind - 1 Step 1
            ce = thereis(lista(8, x), lista(9, x), lf, y, ofile, oOut)
            If Not (ce) Then
                lf(y, 0) = lista(8, x)
                lf(y, 1) = x
                lf(y, 2) = lista(9, x) 'ci mette il nome completo del file anche se è vuoto, serve ad avere un record per ogni file della dir più nel caso uno vuoto
                y = y + 1
            End If
        Next
        popolaListaFile = y ' esporto il livello al quale è arrivata la listaFile
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in popolaListaFile ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function thereis(ByVal dir, ByVal file, ByRef lis, ByVal upto, ByRef ofile, ByRef oOut)
        Dim i, ce_dir
        On Error Resume Next
        For i = 0 To upto Step 1
            If (dir = lis(i, 0)) Then 'se la dir c'è
                If (file = lis(i, 2)) Then 'controllo se anche il file corrisponde
                    thereis = True
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        Next
        thereis = False
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in thereis ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function SeparaRigheColonne(ByRef indirizzo, ByRef riga, ByRef colonna, ByRef ofile, ByRef oOut)
        Dim c, i
        On Error Resume Next
        For i = 1 To Len(indirizzo)
            c = Mid(indirizzo, i, 1)
            If (IsNumeric(c)) Then
                Exit For
            End If
        Next
        'WScript.Echo "Numerico da " & i
        colonna = Mid(indirizzo, 1, i - 1)
        riga = Mid(indirizzo, i, Len(indirizzo))
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in SeparaRigheColonne ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function GetPuntatore(ByVal nome, ByRef file, ByRef sh, ByRef col, ByRef rig, ByRef ofile, ByRef oOut)
        'C:\Sviluppo_secondaria\VbScript\in#fileAction#foglio2#colonnaAB#riga10
        Dim ss, xx, ff, n_parts, i
        On Error Resume Next
        ss = Split(nome, "\")
        n_parts = UBound(ss)
        'in#fileAction#foglio2#colonnaAB#riga10
        xx = Split(ss(n_parts), "#")
        file = Mid(xx(1), 5, Len(xx(1))) 'xx(1) = fileNOMEFILE
        sh = Mid(xx(2), 7, Len(xx(2))) 'xx(2) = foglio2
        col = Mid(xx(3), 8, Len(xx(3))) 'xx(3) = colonnaAB
        rig = Mid(xx(4), 5, Len(xx(4))) 'xx(4) = riga10
        For i = 0 To n_parts - 1 Step 1
            If (ff = "") Then
                ff = ff & ss(i)
            Else
                ff = ff & "\" & ss(i)
            End If
        Next
        ff = ff & "\" & xx(0)
        GetPuntatore = ff
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in GetPuntatore ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function GetCartella(ByVal nome, ByRef file, ByRef ofile, ByRef oOut)
        Dim ss, xx, ff, n_parts, i, pre
        On Error Resume Next
        pre = ""
        If Not (InStr(nome, ":") > 0) Then
            pre = "\\"
        End If
        ss = Split(nome, "\")
        n_parts = UBound(ss)
        For i = 0 To n_parts - 1 Step 1
            If (ff = "") Then
                ff = ff & ss(i)
            Else
                ff = ff & "\" & ss(i)
            End If
        Next
        ff = pre & ff
        file = ss(n_parts)
        GetCartella = ff
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in GetCartella ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function ReadExcelFile(obJWK, ByVal sheet, ByVal Row, ByVal Col, ByRef ofile, ByRef oOut)
        ' Local variable declarations
        Dim objSheet, objCells
        Dim cellCont
        Dim cc, rr
        On Error Resume Next
        cellCont = "nullo"
        ' Default return value
        ReadExcelFile = 0
        'seleziona lo sheet da usare
        objSheet = obJWK.Worksheets(sheet)
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore sel Sheet:" & Err.Number & " Description " & Err.Description, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If
        ' Get the used cells
        objCells = objSheet.Cells
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore set Cells:" & Err.Number & " Description " & Err.Description, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If
        'MsgBox objCells(10, 18).Value
        cellCont = objCells(Row, Col).Value
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore read:" & Err.Number & " Description " & Err.Description, oOut, ofile)
            ofile.Close()
            Console.Read()
            End
        End If
        ReadExcelFile = cellCont
        On Error GoTo 0
    End Function

    Function WEF(objWorkb, ByVal sheet, ByVal Row, ByVal Col, ByVal valore, ByRef ofile, ByRef oOut)
        Dim objSheet, objCells
        WEF = 0
        On Error Resume Next
        'seleziona lo sheet da usare
        objSheet = objWorkb.Worksheets(sheet)
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore WEF sheet select:" & Err.Number & " Description " & Err.Description, oOut, ofile)
            ofile.close()
            Console.Read()
            End
        End If
        objCells = objSheet.Cells
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore WEF set cell:" & Err.Number & " Description " & Err.Description, oOut, ofile)
            ofile.close()
            Console.Read()
            End
        End If
        objCells(Row, Col).Value = valore ' scrive il valore nella cella
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore WEF write cell:" & Err.Number & " Description " & Err.Description, oOut, ofile)
            ofile.close()
            Console.Read()
            End
        End If
        WEF = valore
        On Error GoTo 0
    End Function


    Function cercaFile(ByVal patt, ByVal folder, ByVal fso, ByRef ofile, ByRef oOut) 'pattern da cercare e directory
        Dim filenamecompleto
        Dim f, parte
        Dim objFolder
        On Error Resume Next
        filenamecompleto = "NULLA"
        'WScript.Echo "Folder in cerca:" & folder
        objFolder = fso.GetFolder(folder)
        patt = LCase(patt)
        For Each f In objFolder.Files
            parte = Left(LCase(f.Name), Len(patt)) 'preleva i primi caratteri
            'WScript.Echo patt & " " & parte & " " & LCase(f.Name)
            If InStr(parte, patt) = 0 Then
                'WScript.Echo "---------------- " & LCase(f.Name)
            Else
                'WScript.Echo "Trovato " & patt & " " & f.Name
                filenamecompleto = LCase(f.Name)
            End If
        Next
        cercaFile = filenamecompleto
        If (Err.Number <> 0) Then
            On Error Resume Next
            WriteMia(ConsoleColor.Red, "Errore in cercaFile ErrN." & Err.Number & " Descrizione: " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function scriviSu(ByVal nome, ByVal dato, ByRef ofile, ByRef oOut)
        Dim objFileToWrite
        On Error Resume Next
        'WScript.Echo "File di scrittura:"&nome
        objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(nome, 2, True)
        objFileToWrite.WriteLine(dato)
        objFileToWrite.Close
        objFileToWrite = Nothing
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in scriviSu ErrN." & Err.Number & " Description " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function appendiA(ByVal nome, ByVal dato, ByRef ofile, ByRef oOut)
        Dim objFileToWrite
        On Error Resume Next
        objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(nome, 8, True)
        objFileToWrite.WriteLine(dato)
        objFileToWrite.Close
        objFileToWrite = Nothing
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in appendiA ErrN." & Err.Number & " Description " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function listaDir(dir, ByRef sheet_n, ByRef objWorkbook, ByRef objWorksheet, ByRef objRange, ByRef objExcel, ByRef myPass, ByRef fso, ByRef listaS, ByRef indice_S, ByRef ofile, ByRef oOut)
        Dim objDir, objFile, colFiles, estens
        On Error Resume Next
        objDir = fso.GetFolder(dir)
        'Wscript.Echo objDir.Path
        colFiles = objDir.Files
        For Each objFile In colFiles
            'WBscript.Echo "estensione:" & objFSO.GetExtensionName(objFile.name)
            estens = Mid(objFile.name, InStrRev(objFile.name, ".") + 1, 3)
            'estens = fso.GetExtensionName(objFile.name)
            If (UCase(estens) = "XLS") Then
                If (InStr(1, objFile.name, "~$") = 0) Then
                    'Wscript.Echo "Cerco Parametri nel file:"& objFile.Name
                    loopSheet(dir, objFile.Name, sheet_n, objWorkbook, objWorksheet, objRange, objExcel, myPass, listaS, indice_S, ofile, oOut)
                End If
            End If
        Next
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in listaDir ErrN." & Err.Number & " Description " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

    Function esportaListe(ByVal modo, ByRef out, ByRef lista, ByRef indice, ByRef indice_S, ByRef listaS, ByRef CDir, ByRef Fdebug, ByRef ofile, ByRef oOut)
        On Error Resume Next
        If (Fdebug) Then
            out = "Lista" & vbCrLf
            For y = 0 To indice - 1 Step 1
                out = out & "Parametro			   :" & lista(0, y) & vbCrLf
                out = out & "directory			(D):" & lista(1, y) & vbCrLf
                out = out & "nome file			(D):" & lista(2, y) & vbCrLf
                out = out & "chiave_file		(D):" & lista(3, y) & vbCrLf
                out = out & "sheet				(D):" & lista(4, y) & vbCrLf
                out = out & "colonna n.			(D):" & lista(5, y) & vbCrLf
                out = out & "colonna lettrere	(D):" & lista(6, y) & vbCrLf
                out = out & "riga				(D):" & lista(7, y) & vbCrLf
                out = out & "Directory			(S):" & lista(8, y) & vbCrLf
                out = out & "file				(S):" & lista(9, y) & vbCrLf
                out = out & "chiave_file		(S):" & lista(10, y) & vbCrLf
                out = out & "trovato               :" & lista(11, y) & vbCrLf
            Next
            out = out & "Lista_I" & vbCrLf
            For y = 0 To indice_S - 1 Step 1
                out = out & "Parametro			   :" & listaS(0, y) & vbCrLf
                out = out & "directory			(D):" & listaS(1, y) & vbCrLf
                out = out & "nome file			(D):" & listaS(2, y) & vbCrLf
                out = out & "chiave_file		(D):" & listaS(3, y) & vbCrLf
                out = out & "sheet				(D):" & listaS(4, y) & vbCrLf
                out = out & "colonna n.			(D):" & listaS(5, y) & vbCrLf
                out = out & "colonna lettrere	(D):" & listaS(6, y) & vbCrLf
                out = out & "riga				(D):" & listaS(7, y) & vbCrLf
                out = out & "trovato			   :" & listaS(8, y) & vbCrLf
            Next
            If (modo = "A") Then
                appendiA(CDir & "\liste.txt", out, ofile, oOut)
            Else
                scriviSu(CDir & "\liste.txt", out, ofile, oOut)
            End If
        End If
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore in esportaListe ErrN." & Err.Number & " Description " & Err.Description, oOut, ofile)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function

End Module
