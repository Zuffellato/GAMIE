Imports System.Data
Imports NTSInformatica.CLN__STD
Imports Microsoft.Office.Interop

Public Class CLEZCIMAN
    Inherits CLE__BASE

    Public oCldIMAN As CLDZCIMAN
    Public bAnnullato As Boolean
    Public bData As Date
    Public bCentro As Integer
    Public bCommeca As Integer
    Public bFam As String
    Public CentroTestata As Integer = 0
    Public Anno As Integer = 0
    Public ComAntColt As String = "0"

    Public bNomeFile As String = ""
    Public bNomeFoglio As String = ""
    Public bProcedi As Boolean = False



    Public strDbTransito As String = "PROVA"
    Public strConnString As String = ""
    Public seriedocVend As String = ""
    Public ContoClie As Integer = 0
    Public TipoBFVend As Integer = 0

    Public ContoRisorse As Integer = 0
    Public ContoMezzi As Integer = 0

    Public strMeta As String = ""


    Public Overrides Function Init(ByRef App As CLE__APP, ByRef oScriptEngine As INT__SCRIPT, ByRef oCleLbmenu As Object, ByVal strTabella As String, ByVal bFiller1 As Boolean, ByVal strFiller1 As String, ByVal strFiller2 As String) As Boolean
        Try

            If MyBase.strNomeDal = "BD__BASE" Then MyBase.strNomeDal = "BDZCIMAN"
            MyBase.Init(App, oScriptEngine, oCleLbmenu, strTabella, False, "", "")
            oCldIMAN = CType(MyBase.ocldBase, CLDZCIMAN)
            oCldIMAN.Init(oApp)

            Return True
        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Function

#Region "Import Dati"


    Public Overridable Function LeggiRegistro() As Boolean
        Try
            strDbTransito = NTSCStr(ocldBase.GetSettingBus("BSZCIMAN", "OPZIONI", ".", "DbTransito", " ", " ", " "))
            strConnString = NTSCStr(ocldBase.GetSettingBus("BSZCIMAN", "OPZIONI", ".", "ConnString", "", " ", ""))
            seriedocVend = NTSCStr(ocldBase.GetSettingBus("BSZCIMAN", "OPZIONI", ".", "seriedocVend", "", " ", ""))
            ContoClie = NTSCInt(ocldBase.GetSettingBus("BSZCIMAN", "OPZIONI", ".", "ContoClie", "", " ", 0))
            TipoBFVend = NTSCInt(ocldBase.GetSettingBus("BSZCIMAN", "OPZIONI", ".", "TipoBFVend", "", " ", 0))

            ContoRisorse = NTSCInt(ocldBase.GetSettingBus("BSZCIMAN", "OPZIONI", ".", "ContoRisorse", "", " ", 0))
            ContoMezzi = NTSCInt(ocldBase.GetSettingBus("BSZCIMAN", "OPZIONI", ".", "ContoMezzi", "", " ", 0))
            Return True
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
        Return False
    End Function


    Public Overridable Function ApriConnCrm() As Boolean
        Try
            Return oCldIMAN.ApriConnCrm(strConnString)
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
    End Function

    Public Overridable Function ChiudiConnCrm() As Boolean
        Try
            Return oCldIMAN.ChiudiConnMes()
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
    End Function

    ''---------------------------------------- CANCELLA ----------------------------------------
    'Public Overridable Function ElaboraFile(ByRef RigheImp As Integer) As Boolean
    '    Dim dtDati As New DataTable
    '    Dim i As Integer = 0
    '    Dim drDati As DataRow
    '    Dim dttTmp As New DataTable
    '    Dim strTmp As String = ""
    '    Dim j As Integer = 0
    '    Dim Conto As Integer = 0


    '    Try


    '        'crea datatable con campi del file
    '        CreaStruttureTabelle(dtDati)




    '        'leggo i dati dal file
    '        If Not LeggiFile(dtDati) Then
    '            'MsgBox("Impossibile leggere i dati dal file")
    '            Return False
    '        End If

    '        'verifico la correttezza dei dati nel file
    '        If Not VerificaDati(dtDati) Then Return False

    '        ''Avvio la transazione
    '        'oCldIMAN.dbConn = oCldIMAN.ApriDB(CLE__APP.DBTIPO.DBAZI)
    '        'oCldIMAN.ApriTrans(oCldIMAN.dbConn)

    '        i = 0
    '        'scandisco il datatable con tutti i dati ed inserisco in PRIANA
    '        For Each drDati In dtDati.Rows
    '            If Not oCldIMAN.InsRigaPriana(strDittaCorrente, drDati) Then
    '                ''se sono in transazione la annullo
    '                'If oCldIMAN.IsInTrans Then oCldIMAN.AnnullaTrans()
    '                Return False
    '            End If
    '            i = i + 1
    '        Next


    '        ''Chiudo la transazione ed il database
    '        'oCldIMAN.ChiudiTrans()
    '        'oCldIMAN.dbConn.Close()


    '        RigheImp = i

    '        Return True

    '    Catch ex As Exception
    '        '--------------------------------------------------------------
    '        ''se sono in transazione la annullo
    '        'If oCldIMAN.IsInTrans Then oCldIMAN.AnnullaTrans()


    '        CLN__STD.GestErr(ex, Me, "")

    '        '--------------------------------------------------------------
    '    Finally
    '        ThrowRemoteEvent(New NTSEventArgs("AGGZCSTA", "Pronto."))
    '    End Try
    'End Function
    ''------------------------------------------------------------------------------------------

    Public Overridable Function Elabora(ByVal brisorse As Boolean, ByVal bmezzi As Boolean, ByVal bmateriali As Boolean, ByRef RigheImp As Integer, ByRef oCleBoll As CLEVEBOLL) As Boolean
        Dim dtDati As New DataTable
        Dim i As Integer = 0
        Dim drDati As DataRow
        Dim dttTmp As New DataTable
        Dim strTmp As String = ""
        Dim j As Integer = 0


        Dim dtMaterialiCtrl As New DataTable
        Dim dtRisorseCtrl As New DataTable

        Dim dtMateriali As New DataTable
        Dim dtMaterialiRighe As New DataTable
        Dim dtRisorse As New DataTable
        Dim dtMezzi As New DataTable

        Dim rapportino As String = ""
        Dim rapportinoOld As String = ""
        Dim commessa As Integer = 0

        Dim dtRigheCancMezzi As New DataTable

        Dim dtModMezzi As New DataTable

        Dim dttArtico As New DataTable
        Dim dttCope As New DataTable


        Try

            LeggiRegistro()



            i = 0

            ''MATERIALI
            'If bmateriali Then
            '    'leggo i nuovi inserimenti raggruppati per rapportino per poi ciclare e inserire un ddt per rapportino
            '    oCldIMAN.LeggiMaterialiTestata(dtMateriali, strDbTransito)
            '    For Each drMateriali As DataRow In dtMateriali.Rows
            '        'cerco le righe da inserire per ogni documento e le metto nel datable
            '        oCldIMAN.LeggiMaterialiRighe(0, drMateriali!rapportino.ToString, dtMaterialiRighe, strDbTransito)
            '        If Not CreaDocumentoVendita(oCleBoll, dtMaterialiRighe) Then Return False
            '        'cancello la riga dal db di frontiera
            '    Next

            '    'leggo le modifiche e faccio l'update della quantità
            '    For Each drMateriali As DataRow In dtMateriali.Rows
            '        'cerco le righe da inserire per ogni documento e le metto nel datable
            '        oCldIMAN.LeggiMaterialiRighe(0, drMateriali!rapportino.ToString, dtMaterialiRighe, strDbTransito)
            '        If Not CreaDocumentoVendita(oCleBoll, dtMaterialiRighe) Then Return False
            '    Next
            'End If

            'MATERIALI
            If bmateriali Then
                ' MsgBox("1")
                'apro la connessione sul db di scambio
                If Not oCldIMAN.ApriConnCrm(strConnString) Then
                    oApp.MsgBoxInfo("ATTENZIONE!!! Impossibile accedere al database di scambio (" & strConnString & ").")
                    Return False
                End If

                'leggo quanto presente nella tabella dei materiali per poi creare un file con gli articoli che mancano
                oCldIMAN.LeggiMateriali(dtMaterialiCtrl, strDbTransito)
                ' LogStart("BSZCIMAN_" & Today.Year & "-" & Today.Month & "-" & Today.Day & "_" & Now.Hour & "h" & Now.Minute & "m" & Now.Second & "s", "BSZCIMAN")

                strMeta = ""

                For Each drMaterialictrl As DataRow In dtMaterialiCtrl.Rows
                    'controllo che esista il codice dell'articolo, se non esistye cotruisco il log
                    If drMaterialictrl!codice.ToString <> "" Then
                        'esiste?
                        oCldIMAN.ValCodiceDb(NTSCStr(drMaterialictrl!codice), strDittaCorrente, "artico", "", "", dttArtico)
                        If dttArtico.Rows.Count <= 0 Then
                            strMeta = "Elabora Articoli..."
                            LogWrite("Articolo " & NTSCStr(drMaterialictrl!codice) & " inesistente in Business.", False)
                            'Return False
                        Else
                        End If
                    End If

                Next

                If strMeta <> "" Then
                    'chiudo la connessione con il db di scambio
                    oCldIMAN.ChiudiConnMes()

                    Return False
                End If

                'MsgBox("ctrl")
                'Return False


                'leggo i record e li metto in due datatable, uno di testata e uno di righe
                oCldIMAN.LeggiMateriali(dtMateriali, strDbTransito)
                oCldIMAN.LeggiMateriali(dtMaterialiRighe, strDbTransito)

                'INSERIMENTI: prendo tutti i nuovi inserimenti (stato 0) ordinati per rapportino e li scorro
                For Each drMateriali As DataRow In dtMateriali.Select("stato = 0", "rapportino")
                    'mi salvo il rapportino attuale per creare un documento per rapportino
                    rapportino = drMateriali!rapportino.ToString

                    'se il numero di rapportino cambia devo fare un altro documento
                    If rapportino <> rapportinoOld Then
                        commessa = drMateriali!commessa.ToString
                        'creo il documento passando le righe materiali che poi filtrerò per rapportino corrente nella funzione
                        If Not CreaDocumentoVendita(oCleBoll, dtMaterialiRighe, rapportino, commessa) Then Return False
                    End If
                    'setto il rapportino precedente per fare il confronto
                    rapportinoOld = drMateriali!rapportino.ToString
                Next

                For Each drMateriali As DataRow In dtMateriali.Select("stato = 1", "rapportino")
                    'cerco il documento per quel rapportino: 
                    rapportino = drMateriali!rapportino.ToString

                    oCldIMAN.TrovaDDTRigaCancellataMateriale(drMateriali!codice.ToString, drMateriali!rapportino.ToString, drMateriali!commessa, strDittaCorrente, dtModMezzi)

                    If Not oCleBoll.ApriDoc(strDittaCorrente, False, dtModMezzi.Rows(0)!tm_tipork.ToString, dtModMezzi.Rows(0)!tm_anno, dtModMezzi.Rows(0)!tm_serie.ToString, dtModMezzi.Rows(0)!tm_numdoc.ToString, oCleBoll.dsShared) Then Return False

                    For Each dr As DataRow In oCleBoll.dsShared.Tables("CORPO").Select("ec_codart= " & CStrSQL(drMateriali!codice.ToString))

                        dr!ec_colli = drMateriali!quantita
                        dr!ec_quant = drMateriali!quantita

                        oCleBoll.dsShared.Tables("CORPO").AcceptChanges()

                    Next

                    oCleBoll.SalvaDocumento("U")

                    'cancello le righe dal db di frontiera che ho importato
                    '  For Each drdati2 As DataRow In dtDati.Select("stato=1 and rapportino=" & CStrSQL(rapportino))
                    oCldIMAN.CancRigaImportataMaterialeDBFrontiera(drMateriali!id, strDbTransito)
                    ' Next


                Next

                'CANCELLAZIONI: scorro il datatable e cerco tutte le righe da cancellare
                For Each drMateriali As DataRow In dtMateriali.Select("stato = 2")
                    'per ogni riga da cancellare intercetto il suo ddt
                    oCldIMAN.TrovaDDTRigaCancellataMateriale(drMateriali!codice.ToString, drMateriali!rapportino.ToString, drMateriali!commessa, strDittaCorrente, dtRigheCancMezzi)
                    Dim totrighe As Integer = 0
                    If dtRigheCancMezzi.Rows.Count >= 0 Then
                        'verifico se il ddt ha una sola riga
                        oCldIMAN.TornaNumeroRigheDDT(dtRigheCancMezzi.Rows(0)!mm_tipork.ToString, dtRigheCancMezzi.Rows(0)!mm_anno, dtRigheCancMezzi.Rows(0)!mm_serie.ToString, dtRigheCancMezzi.Rows(0)!mm_numdoc, strDittaCorrente, totrighe)
                        If totrighe = 1 Then
                            'il ddt ha una sola riga quindi cancello tutto il ddt
                            If Not oCleBoll.ApriDoc(strDittaCorrente, False, dtRigheCancMezzi.Rows(0)!tm_tipork.ToString, dtRigheCancMezzi.Rows(0)!tm_anno, dtRigheCancMezzi.Rows(0)!tm_serie.ToString, dtRigheCancMezzi.Rows(0)!tm_numdoc.ToString, oCleBoll.dsShared) Then Return False

                            oCleBoll.SalvaDocumento("D")
                        Else
                            'il ddt ha più righe cancello solo quella riga
                            'apro il documento
                            If Not oCleBoll.ApriDoc(strDittaCorrente, False, dtRigheCancMezzi.Rows(0)!tm_tipork.ToString, dtRigheCancMezzi.Rows(0)!tm_anno, dtRigheCancMezzi.Rows(0)!tm_serie.ToString, dtRigheCancMezzi.Rows(0)!tm_numdoc.ToString, oCleBoll.dsShared) Then Return False


                            For Each dr As DataRow In oCleBoll.dsShared.Tables("CORPO").Select("ec_riga= " & dtRigheCancMezzi.Rows(0)!mm_riga)
                                'oCleBoll.RecordSalva(1, True, dr)

                                'oCleBoll.RecordSalva(dr, True, dr)
                                oCleBoll.dsShared.Tables("CORPO").Rows.Remove(dr)
                                oCleBoll.dsShared.Tables("CORPO").AcceptChanges()
                            Next

                            oCleBoll.SalvaDocumento("U")


                        End If


                        '    For Each drMaterialiCanc As DataRow In dtRigheCancMezzi.Select("stato = 2", "rapportino")
                        '        rapportino = drMaterialiCanc!rapportino.ToString
                        '        oCldIMAN.TornaNumeroRigheDDT(dtRigheCancMezzi.Rows(0)!mm_tipork.ToString, dtRigheCancMezzi.Rows(0)!mm_anno, dtRigheCancMezzi.Rows(0)!mm_serie.ToString, dtRigheCancMezzi.Rows(0)!mm_numdoc, strDittaCorrente, totrighe)


                        '        If rapportino <> rapportinoOld Then

                        '        End If
                        '        rapportinoOld = drMateriali!rapportino.ToString
                        '    Next
                        '        oCldIMAN.TornaNumeroRigheDDT(dtRigheCancMezzi.Rows(0)!mm_tipork.ToString, dtRigheCancMezzi.Rows(0)!mm_anno, dtRigheCancMezzi.Rows(0)!mm_serie.ToString, dtRigheCancMezzi.Rows(0)!mm_numdoc, strDittaCorrente, totrighe)

                        '        'se il ddt ha una sola riga cancello tutto il ddt
                        '        If totrighe = 1 Then
                        '            If Not oCleBoll.ApriDoc(strDittaCorrente, True, dtRigheCancMezzi.Rows(0)!tm_tipork.ToString, dtRigheCancMezzi.Rows(0)!tm_anno, dtRigheCancMezzi.Rows(0)!tm_serie.ToString, dtRigheCancMezzi.Rows(0)!tm_numdoc.ToString, oCleBoll.dsShared) Then Return False
                        '            Return oCleBoll.SalvaDocumento("D")
                        '        Else
                        '            'se il ddt ha più righe cancello solo quella
                        '            If Not oCleBoll.ApriDoc(strDittaCorrente, True, dtRigheCancMezzi.Rows(0)!tm_tipork.ToString, dtRigheCancMezzi.Rows(0)!tm_anno, dtRigheCancMezzi.Rows(0)!tm_serie.ToString, dtRigheCancMezzi.Rows(0)!tm_numdoc.ToString, oCleBoll.dsShared) Then Return False
                        '            oCleBoll.dttEC.Select("mm_codart=" & CStrSQL(dtRigheCancMezzi.Rows(0)!mm_codart.ToString) & " and mm_hhrapp= " & CStrSQL(dtRigheCancMezzi.Rows(0)!mm_codart.ToString) & " and mm_quant=" & CDblSQL(dtRigheCancMezzi.Rows(0)!mm_quant))

                        '            ' dtRigheCancMezzi.Rows.Count > 1
                        '        End If

                    End If
                Next

                ''MODIFICHE: scorro il datatable e cerco tutte le righe da cancellare
                'For Each drMateriali As DataRow In dtMateriali.Select("stato = 2")

                'Next

                'chiudo la connessione con il db di scambio
                oCldIMAN.ChiudiConnMes()



            End If
            'RISORSE
            If brisorse Then
                'MsgBox("2")
                'apro la connessione sul db di scambio
                If Not oCldIMAN.ApriConnCrm(strConnString) Then
                    oApp.MsgBoxInfo("ATTENZIONE!!! Impossibile accedere al database di scambio (" & strConnString & ").")
                    Return False
                End If

                'leggo quanto presente nella tabella delle risorse per poi creare un file con gli articoli che mancano
                oCldIMAN.LeggiRisorse(dtRisorseCtrl, strDbTransito)
                ' LogStart("BSZCIMAN_" & Today.Year & "-" & Today.Month & "-" & Today.Day & "_" & Now.Hour & "h" & Now.Minute & "m" & Now.Second & "s", "BSZCIMAN")

                strMeta = ""

                For Each drRisorsectrl As DataRow In dtRisorseCtrl.Rows
                    'controllo che esista il codice dell'articolo, se non esistye cotruisco il log
                    If drRisorsectrl!codice.ToString <> "" Then
                        'esiste?
                        oCldIMAN.ValCodiceDb(NTSCStr(drRisorsectrl!codice), strDittaCorrente, "tabcope", "", "", dttCope)
                        If dttCope.Rows.Count <= 0 Then
                            strMeta = "Elabora Risorse..."
                            LogWrite("Risorsa " & NTSCStr(drRisorsectrl!codice) & " inesistente in Business.", False)
                            'Return False
                        Else
                        End If
                    End If

                Next

                If strMeta <> "" Then
                    'chiudo la connessione con il db di scambio
                    oCldIMAN.ChiudiConnMes()

                    Return False
                End If

                'MsgBox("ctrl")
                'Return False


                ''apro la connessione sul db di scambio
                'If Not oCldIMAN.ApriConnCrm(strConnString) Then
                '    oApp.MsgBoxInfo("ATTENZIONE!!! Impossibile accedere al database di scambio (" & strConnString & ").")
                '    Return False
                'End If

                'Nuovi inserimenti: faccio un giro dove inserisco tutte le righe nuove in priana e poi cancello dal db di frontiera
                oCldIMAN.LeggiRisorse(dtRisorse, strDbTransito)
                For Each drRisorse As DataRow In dtRisorse.Rows
                    Select Case drRisorse!stato
                        Case 0
                            'inserisco una nuova riga in priana
                            oCldIMAN.InsRigaPrianaRisorse(strDittaCorrente, ContoRisorse, drRisorse)
                            'cancello la riga dalla frontiera
                            oCldIMAN.CancRigaImportataRisorseDaDBFrontiera(drRisorse!id, strDbTransito)
                    End Select
                Next

                'Cancellazioni: faccio un giro dove cancello tutte le righe da priana esposte nel db e poi cancello dal db di frontiera
                For Each drRisorse As DataRow In dtRisorse.Rows
                    Select Case drRisorse!stato
                        Case 2
                            'cancella la riga
                            oCldIMAN.CancRigaPrianaRisorse(strDittaCorrente, drRisorse!codice, drRisorse!rapportino, CDate(drRisorse!data), drRisorse!id, drRisorse!dalle.ToString, drRisorse!alle.ToString, drRisorse!commessa, CInt(drRisorse!tipologia))

                            'cancello la riga dalla frontiera
                            oCldIMAN.CancRigaImportataRisorseDaDBFrontiera(drRisorse!id, strDbTransito)
                    End Select
                Next

                'Modifiche: faccio un primo giro dove cancello da priana le righe da modificare poi un secondo giro deove le reinserisco
                For Each drRisorse As DataRow In dtRisorse.Rows
                    Select Case drRisorse!stato
                        Case 1
                            'cancello la riga da modificare da priana
                            oCldIMAN.CancRigaPrianaRisorsedaMod(strDittaCorrente, drRisorse!codice, drRisorse!rapportino, CDate(drRisorse!data), drRisorse!id, drRisorse!dalle.ToString, drRisorse!alle.ToString, drRisorse!commessa, CInt(drRisorse!tipologia))
                    End Select
                Next

                For Each drRisorse As DataRow In dtRisorse.Rows
                    Select Case drRisorse!stato
                        Case 1
                            'inserisco una nuova riga in priana
                            oCldIMAN.InsRigaPrianaRisorse(strDittaCorrente, ContoRisorse, drRisorse)
                            'cancello la riga dalla frontiera
                            oCldIMAN.CancRigaImportataRisorseDaDBFrontiera(drRisorse!id, strDbTransito)
                    End Select
                Next

                'chiudo la connessione con il db di scambio
                oCldIMAN.ChiudiConnMes()


            End If

            'MEZZI
            If bmezzi Then
                '   MsgBox("3")
                'apro la connessione sul db di scambio
                If Not oCldIMAN.ApriConnCrm(strConnString) Then
                    oApp.MsgBoxInfo("ATTENZIONE!!! Impossibile accedere al database di scambio (" & strConnString & ").")
                    Return False
                End If

                'Nuovi inserimenti: faccio un giro dove inserisco tutte le righe nuove in priana e poi cancello dal db di frontiera
                oCldIMAN.LeggiMezzi(dtMezzi, strDbTransito)
                For Each drMezzi As DataRow In dtMezzi.Rows
                    Select Case drMezzi!stato
                        Case 0
                            'inserisco una nuova riga in priana
                            oCldIMAN.InsRigaPrianaMezzi(strDittaCorrente, ContoRisorse, drMezzi)
                            'cancello la riga dalla frontiera
                            oCldIMAN.CancRigaImportataMezziDaDBFrontiera(drMezzi!id, strDbTransito)
                    End Select
                Next

                'Cancellazioni: faccio un giro dove cancello tutte le righe da priana esposte nel db e poi cancello dal db di frontiera
                For Each drMezzi As DataRow In dtMezzi.Rows
                    Select Case drMezzi!stato
                        Case 2
                            'cancella la riga
                            oCldIMAN.CancRigaPrianaMezzi(strDittaCorrente, drMezzi!codice, drMezzi!rapportino, drMezzi!id, drMezzi!ore)
                            'cancello la riga dalla frontiera
                            oCldIMAN.CancRigaImportataMezziDaDBFrontiera(drMezzi!id, strDbTransito)
                    End Select
                Next

                'Modifiche: faccio un primo giro dove cancello da priana le righe da modificare poi un secondo giro deove le reinserisco
                For Each drMezzi As DataRow In dtMezzi.Rows
                    Select Case drMezzi!stato
                        Case 1
                            'cancello la riga da modificare da priana
                            oCldIMAN.CancRigaPrianaMezzi(strDittaCorrente, drMezzi!codice, drMezzi!rapportino, drMezzi!id, drMezzi!prevalue)
                    End Select
                Next

                For Each drMezzi As DataRow In dtMezzi.Rows
                    Select Case drMezzi!stato
                        Case 1
                            'inserisco una nuova riga in priana
                            oCldIMAN.InsRigaPrianaMezzi(strDittaCorrente, ContoRisorse, drMezzi)
                            'cancello la riga dalla frontiera
                            oCldIMAN.CancRigaImportataMezziDaDBFrontiera(drMezzi!id, strDbTransito)
                    End Select
                Next

                'chiudo la connessione con il db di scambio
                oCldIMAN.ChiudiConnMes()

            End If


            ''verifico la correttezza dei dati nelle tabelle...da capire 
            'If Not VerificaDati(dtDati) Then Return False

            'i = 0
            ''scandisco il datatable delle risorse ed inserisco in PRIANA
            'For Each drDati In dtDati.Rows
            '    If Not oCldIMAN.InsRigaPriana(strDittaCorrente, drDati) Then
            '        ''se sono in transazione la annullo
            '        'If oCldIMAN.IsInTrans Then oCldIMAN.AnnullaTrans()
            '        Return False
            '    End If
            '    i = i + 1
            'Next

            'RigheImp = i

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            ''se sono in transazione la annullo
            'If oCldIMAN.IsInTrans Then oCldIMAN.AnnullaTrans()


            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        Finally
            ThrowRemoteEvent(New NTSEventArgs("AGGZCSTA", "Pronto."))
        End Try
    End Function


    Public Overrides Function LogWrite(ByVal strMsg As String, ByVal bIncrementaCount As Boolean) As Boolean
        Try
            'If strMeta.Trim.Length > 0 Then
            '    If strMsg.IndexOf(":"c) > -1 AndAlso strMsg.IndexOf(":"c) < 20 Then
            '        strMsg = strMsg.Substring(0, strMsg.IndexOf(":"c)).Trim & ": " & strMeta.Trim & " " & strMsg.Substring(strMsg.IndexOf(":"c) + 1).Trim
            '    Else
            '        strMsg = strMeta & " " & strMsg
            '    End If
            'End If
            'MyBase.LogWrite(strMsg, bIncrementaCount)
            'Return True

            Return MyBase.LogWrite("[" & strMeta & "] " & strMsg, bIncrementaCount)

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
        Return False
    End Function

    'Public Overridable Function CreaDocumentoVendita(ByRef oCleBoll As CLEVEBOLL, ByVal dtDati As DataTable) As Boolean
    '    Dim dataDocumento As Date
    '    Dim nAnno As Integer = 0
    '    Dim lnumdocTmp As Integer = 0
    '    Dim dsGsbo As New DataSet
    '    Dim Conto As Integer = 0
    '    Dim dsOrdine As New DataSet

    '    Try

    '        'data documento
    '        dataDocumento = NTSCDate(Now.Date)

    '        'ricavo l'anno
    '        nAnno = Year(dataDocumento)

    '        'creo un nuovo documento
    '        lnumdocTmp = oCleBoll.LegNuma("B", seriedocVend, nAnno)
    '        If lnumdocTmp = 0 Then
    '            Return False
    '        End If

    '        If Not oCleBoll.ApriDoc(strDittaCorrente, True, "B", 0, " ", 0, dsGsbo) Then Return False

    '        If oCleBoll.NuovoDocumento(strDittaCorrente, "B", nAnno, seriedocVend, lnumdocTmp) = False Then Return False

    '        dsGsbo.Tables("TESTA").AcceptChanges()
    '        oCleBoll.dttEC.Clear()

    '        'setto i dati di testata
    '        oCleBoll.dttET.Rows(0)!et_tipobf = TipoBFVend
    '        oCleBoll.dttET.Rows(0)!et_listino = 0
    '        oCleBoll.dttET.Rows(0)!et_speacc = 0
    '        oCleBoll.dttET.Rows(0)!et_scorpo = "N"
    '        oCleBoll.dttET.Rows(0)!et_datdoc = NTSCDate(dataDocumento)
    '        oCleBoll.dttET.Rows(0)!et_riferim = NTSCStr(dtDati.Rows(0)!rapportino)
    '        'da sostituire probabilmente con un dato che ci passeranno in tabella altrimenti metti come opz di reg
    '        oCleBoll.dttET.Rows(0)!et_magaz = 1
    '        oCleBoll.dttET.Rows(0)!et_conto = ContoClie

    '        For Each drdati As DataRow In dtDati.Rows
    '            With oCleBoll.dttEC.Rows.Add()
    '                !codditt = "."
    '                !codditt = strDittaCorrente
    '                'da sostituire probabilmente con un dato che ci passeranno in tabella altrimenti metti come opz di reg
    '                !ec_magaz = 1
    '                !ec_codart = drdati!codice
    '                !ec_colli = drdati!quantita
    '                !ec_quant = drdati!quantita
    '                !ec_note = drdati!commessa

    '            End With
    '        Next

    '        'salvo il documento
    '        If Not oCleBoll.SalvaDocumento("N") Then Return False

    '        'setto lo stato come importato
    '        oCldIMAN.SettaMaterialeImportato(0, NTSCStr(dtDati.Rows(0)!rapportino), strDbTransito)
    '        Return True


    '    Catch ex As Exception
    '        Dim strErr As String = CLN__STD.GestError(ex, Nothing, "", oApp.InfoError, oApp.ErrorLogFile, True)
    '    End Try

    'End Function

    Public Overridable Function CreaDocumentoVendita(ByRef oCleBoll As CLEVEBOLL, ByVal dtDati As DataTable, ByVal rapportino As String, ByVal commessa As Integer) As Boolean
        Dim dataDocumento As Date
        Dim nAnno As Integer = 0
        Dim lnumdocTmp As Integer = 0
        Dim dsGsbo As New DataSet
        Dim Conto As Integer = 0
        Dim dsOrdine As New DataSet

        Try

            'data documento
            dataDocumento = NTSCDate(Now.Date)

            'ricavo l'anno
            nAnno = Year(dataDocumento)

            'creo un nuovo documento
            lnumdocTmp = oCleBoll.LegNuma("B", seriedocVend, nAnno)
            If lnumdocTmp = 0 Then
                Return False
            End If

            If Not oCleBoll.ApriDoc(strDittaCorrente, True, "B", 0, " ", 0, dsGsbo) Then Return False

            If oCleBoll.NuovoDocumento(strDittaCorrente, "B", nAnno, seriedocVend, lnumdocTmp) = False Then Return False

            dsGsbo.Tables("TESTA").AcceptChanges()
            oCleBoll.dttEC.Clear()

            'setto i dati di testata
            oCleBoll.dttET.Rows(0)!et_tipobf = TipoBFVend
            oCleBoll.dttET.Rows(0)!et_listino = 0
            oCleBoll.dttET.Rows(0)!et_speacc = 0
            oCleBoll.dttET.Rows(0)!et_scorpo = "N"
            oCleBoll.dttET.Rows(0)!et_datdoc = NTSCDate(dataDocumento)
            oCleBoll.dttET.Rows(0)!et_riferim = NTSCStr(rapportino)

            'da sostituire probabilmente con un dato che ci passeranno in tabella altrimenti metti come opz di reg
            oCleBoll.dttET.Rows(0)!et_magaz = 1
            oCleBoll.dttET.Rows(0)!et_conto = ContoClie
            oCleBoll.dttET.Rows(0)!et_commeca = NTSCStr(commessa)

            For Each drdati As DataRow In dtDati.Select("stato=0 and rapportino=" & CStrSQL(rapportino))
                With oCleBoll.dttEC.Rows.Add()
                    !codditt = "."
                    !codditt = strDittaCorrente
                    'da sostituire probabilmente con un dato che ci passeranno in tabella altrimenti metti come opz di reg
                    !ec_magaz = drdati!codmag
                    !ec_codart = drdati!codice
                    !ec_colli = drdati!quantita
                    !ec_quant = drdati!quantita
                    !ec_note = drdati!commessa
                    !ec_hhrapp = drdati!rapportino
                    !ec_commeca = drdati!commessa

                End With
            Next


            'salvo il documento
            If Not oCleBoll.SalvaDocumento("N") Then Return False

            'cancello le righe dal db di frontiera che ho importato
            For Each drdati As DataRow In dtDati.Select("stato=0 and rapportino=" & CStrSQL(rapportino))
                oCldIMAN.CancRigaImportataMaterialeDBFrontiera(drdati!id, strDbTransito)
            Next

            Return True


        Catch ex As Exception
            Dim strErr As String = CLN__STD.GestError(ex, Nothing, "", oApp.InfoError, oApp.ErrorLogFile, True)
        End Try

    End Function



    '------------------------------------------- CANCELLA -----------------------------------------
    Public Overridable Sub CreaStruttureTabelle(ByRef dtDati As DataTable)
        Dim i As Integer = 0

        Try
            If dtDati.Columns.Count = 0 Then
                dtDati.Columns.Add("pa_integr", GetType(String))
                dtDati.Columns.Add("pa_datreg", GetType(Date))
                dtDati.Columns.Add("pa_numreg", GetType(Integer))
                dtDati.Columns.Add("pa_datdoc", GetType(Date))
                dtDati.Columns.Add("pa_numdoc", GetType(Integer))
                dtDati.Columns.Add("pa_alfdoc", GetType(String))
                dtDati.Columns.Add("pa_codart", GetType(String))
                dtDati.Columns.Add("pa_desart", GetType(String))
                dtDati.Columns.Add("pa_note", GetType(String))
                dtDati.Columns.Add("pa_codlavo", GetType(Integer))
                dtDati.Columns.Add("pa_codcena", GetType(Integer))
                dtDati.Columns.Add("pa_unmis", GetType(String))
                dtDati.Columns.Add("pa_quant", GetType(Decimal))
                dtDati.Columns.Add("pa_darave", GetType(String))
                dtDati.Columns.Add("pa_importo", GetType(Decimal))
                dtDati.Columns.Add("pa_importoda", GetType(Decimal))
                dtDati.Columns.Add("pa_codcfam", GetType(String))
                dtDati.Columns.Add("pa_commeca", GetType(Integer))
                dtDati.Columns.Add("pa_conto", GetType(Integer))
                dtDati.Columns.Add("pa_controp", GetType(Integer))
                dtDati.Columns.Add("pa_codcope", GetType(Integer))
                dtDati.Columns.Add("pa_subcommeca", GetType(String))
                dtDati.Columns.Add("pa_tipore", GetType(String))

                'questo potrebbe servire per il rapportino?
                'dtDati.Columns.Add("pa_hhprogr", GetType(Integer))
                dtDati.Columns.Add("pa_hhcodrapp", GetType(String))
            Else
                dtDati.Clear()
                dtDati.AcceptChanges()
            End If

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Sub
    '------------------------------------------------------------------------


    Public Overridable Sub CreaStrutturaTabellaAnalitica(ByRef dtDati As DataTable)
        Dim i As Integer = 0

        Try
            If dtDati.Columns.Count = 0 Then
                'codice
                dtDati.Columns.Add("pa_codcope", GetType(Integer))
                dtDati.Columns.Add("pa_integr", GetType(String))
                dtDati.Columns.Add("pa_datreg", GetType(Date))
                dtDati.Columns.Add("pa_numreg", GetType(Integer))
                dtDati.Columns.Add("pa_datdoc", GetType(Date))
                dtDati.Columns.Add("pa_numdoc", GetType(Integer))
                dtDati.Columns.Add("pa_alfdoc", GetType(String))
                dtDati.Columns.Add("pa_codart", GetType(String))
                dtDati.Columns.Add("pa_desart", GetType(String))
                dtDati.Columns.Add("pa_note", GetType(String))
                dtDati.Columns.Add("pa_codlavo", GetType(Integer))
                dtDati.Columns.Add("pa_codcena", GetType(Integer))
                dtDati.Columns.Add("pa_unmis", GetType(String))
                dtDati.Columns.Add("pa_quant", GetType(Decimal))
                dtDati.Columns.Add("pa_darave", GetType(String))
                dtDati.Columns.Add("pa_importo", GetType(Decimal))
                dtDati.Columns.Add("pa_importoda", GetType(Decimal))
                dtDati.Columns.Add("pa_codcfam", GetType(String))
                dtDati.Columns.Add("pa_commeca", GetType(Integer))
                dtDati.Columns.Add("pa_conto", GetType(Integer))
                dtDati.Columns.Add("pa_controp", GetType(Integer))
                dtDati.Columns.Add("pa_subcommeca", GetType(String))
                dtDati.Columns.Add("pa_tipore", GetType(String))

                'questo potrebbe servire per il rapportino?
                dtDati.Columns.Add("pa_hhcodrapp", GetType(String))

            Else
                dtDati.Clear()
                dtDati.AcceptChanges()
            End If

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Sub

    Public Overridable Function LeggiFile(ByRef dtDati As DataTable) As Boolean
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlRange As Excel.Range
        Dim nRiga As Integer = 0
        Dim dtrNuovaRiga As DataRow
        Dim NumMese As Integer = 0
        Dim dtrT() As DataRow = Nothing
        Dim sheet As Excel.Worksheet
        Dim FoglioEsiste As Boolean = False
        Dim dtTmp As New DataTable
        Dim Gruppo As Integer = 0
        Dim Sottogruppo As Integer = 0
        Dim Classifica1 As String = ""
        Dim Classifica2 As String = ""
        Dim DescrArt As String = ""
        Dim UnmisArt As String = ""
        Dim TipoOre As String = "N"
        Dim MinOpe As Integer = 0
        Dim MinTrat As Integer = 0
        Dim Tenuta As String = ""


        Try


            'avvia l'applicazione Excel
            xlApp = New Excel.Application
            'apre il file Excel
            xlWorkBook = xlApp.Workbooks.Open(bNomeFile, True)
            'verifico se il foglio di lavoro esiste
            FoglioEsiste = False
            For Each sheet In xlWorkBook.Sheets
                If sheet.Name = bNomeFoglio Then
                    FoglioEsiste = True
                End If
            Next
            If FoglioEsiste = False Then
                'se il foglio non esiste non posso andare avanti
                MsgBox("Il foglio indicato non esiste nel file selezionato. Impossibile continuare.")
                Return False
            End If

            'apre il foglio di lavoro
            xlWorkSheet = CType(xlWorkBook.Sheets(bNomeFoglio), Excel.Worksheet)


            xlRange = xlWorkSheet.UsedRange


            For nRiga = 2 To xlRange.Rows.Count
                'controllo che siano compilati il codice operaio e il numero ore
                If NTSCStr(CType(xlRange.Cells(nRiga, 1), Excel.Range).Value) <> "" Then
                    'riga nuova
                    dtrNuovaRiga = dtDati.NewRow()
                    dtrNuovaRiga!pa_integr = NTSCStr(CType(xlRange.Cells(nRiga, 1), Excel.Range).Value)
                    dtrNuovaRiga!pa_datreg = NTSCDate(CType(xlRange.Cells(nRiga, 3), Excel.Range).Value)
                    dtrNuovaRiga!pa_numreg = NTSCInt(CType(xlRange.Cells(nRiga, 4), Excel.Range).Value)
                    dtrNuovaRiga!pa_datdoc = NTSCDate(CType(xlRange.Cells(nRiga, 5), Excel.Range).Value)
                    dtrNuovaRiga!pa_numdoc = NTSCInt(CType(xlRange.Cells(nRiga, 6), Excel.Range).Value)
                    dtrNuovaRiga!pa_alfdoc = NTSCStr(CType(xlRange.Cells(nRiga, 7), Excel.Range).Value)
                    dtrNuovaRiga!pa_codart = NTSCStr(CType(xlRange.Cells(nRiga, 8), Excel.Range).Value)
                    dtrNuovaRiga!pa_desart = NTSCStr(CType(xlRange.Cells(nRiga, 9), Excel.Range).Value)
                    dtrNuovaRiga!pa_note = NTSCStr(CType(xlRange.Cells(nRiga, 10), Excel.Range).Value)
                    dtrNuovaRiga!pa_codlavo = NTSCInt(CType(xlRange.Cells(nRiga, 11), Excel.Range).Value)
                    dtrNuovaRiga!pa_codcena = NTSCInt(CType(xlRange.Cells(nRiga, 13), Excel.Range).Value)
                    dtrNuovaRiga!pa_unmis = NTSCStr(CType(xlRange.Cells(nRiga, 15), Excel.Range).Value)
                    dtrNuovaRiga!pa_quant = NTSCDec(CType(xlRange.Cells(nRiga, 16), Excel.Range).Value)
                    'compilato il dare
                    If NTSCStr(CType(xlRange.Cells(nRiga, 17), Excel.Range).Value).Trim <> "" Then
                        dtrNuovaRiga!pa_darave = "D"
                        dtrNuovaRiga!pa_importo = NTSCDec(CType(xlRange.Cells(nRiga, 17), Excel.Range).Value)
                        dtrNuovaRiga!pa_importoda = NTSCDec(CType(xlRange.Cells(nRiga, 17), Excel.Range).Value)
                    End If
                    'compilato l'avere
                    If NTSCStr(CType(xlRange.Cells(nRiga, 18), Excel.Range).Value).Trim <> "" Then
                        dtrNuovaRiga!pa_darave = "A"
                        dtrNuovaRiga!pa_importo = NTSCDec(CType(xlRange.Cells(nRiga, 18), Excel.Range).Value)
                        dtrNuovaRiga!pa_importoda = NTSCDec(CType(xlRange.Cells(nRiga, 18), Excel.Range).Value)
                    End If
                    dtrNuovaRiga!pa_codcfam = Right("0000" & NTSCStr(CType(xlRange.Cells(nRiga, 19), Excel.Range).Value), 4)
                    dtrNuovaRiga!pa_commeca = NTSCInt(CType(xlRange.Cells(nRiga, 21), Excel.Range).Value)
                    dtrNuovaRiga!pa_conto = NTSCInt(CType(xlRange.Cells(nRiga, 23), Excel.Range).Value)
                    dtrNuovaRiga!pa_controp = NTSCInt(CType(xlRange.Cells(nRiga, 25), Excel.Range).Value)
                    dtrNuovaRiga!pa_codcope = NTSCInt(CType(xlRange.Cells(nRiga, 27), Excel.Range).Value)
                    dtrNuovaRiga!pa_subcommeca = NTSCStr(CType(xlRange.Cells(nRiga, 29), Excel.Range).Value)
                    dtrNuovaRiga!pa_tipore = NTSCStr(CType(xlRange.Cells(nRiga, 30), Excel.Range).Value)
                    dtrNuovaRiga!pa_hhprogr = NTSCInt(CType(xlRange.Cells(nRiga, 2), Excel.Range).Value)

                    dtDati.Rows.Add(dtrNuovaRiga)
                    dtDati.AcceptChanges()

                End If
            Next



            xlWorkBook.Close(False)
            xlApp.Quit()


            Return True

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        Finally
            ScaricaOggettoExcel(xlApp)
            ScaricaOggettoExcel(xlWorkBook)
            ScaricaOggettoExcel(xlWorkSheet)
        End Try
    End Function


    Public Overridable Function LeggiRisorse(ByRef dtDati As DataTable) As Boolean
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlRange As Excel.Range
        Dim nRiga As Integer = 0
        Dim dtrNuovaRiga As DataRow
        Dim NumMese As Integer = 0
        Dim dtrT() As DataRow = Nothing
        Dim sheet As Excel.Worksheet
        Dim FoglioEsiste As Boolean = False
        Dim dtTmp As New DataTable
        Dim Gruppo As Integer = 0
        Dim Sottogruppo As Integer = 0
        Dim Classifica1 As String = ""
        Dim Classifica2 As String = ""
        Dim DescrArt As String = ""
        Dim UnmisArt As String = ""
        Dim TipoOre As String = "N"
        Dim MinOpe As Integer = 0
        Dim MinTrat As Integer = 0
        Dim Tenuta As String = ""


        Try


            'avvia l'applicazione Excel
            xlApp = New Excel.Application
            'apre il file Excel
            xlWorkBook = xlApp.Workbooks.Open(bNomeFile, True)
            'verifico se il foglio di lavoro esiste
            FoglioEsiste = False
            For Each sheet In xlWorkBook.Sheets
                If sheet.Name = bNomeFoglio Then
                    FoglioEsiste = True
                End If
            Next
            If FoglioEsiste = False Then
                'se il foglio non esiste non posso andare avanti
                MsgBox("Il foglio indicato non esiste nel file selezionato. Impossibile continuare.")
                Return False
            End If

            'apre il foglio di lavoro
            xlWorkSheet = CType(xlWorkBook.Sheets(bNomeFoglio), Excel.Worksheet)


            xlRange = xlWorkSheet.UsedRange


            For nRiga = 2 To xlRange.Rows.Count
                'controllo che siano compilati il codice operaio e il numero ore
                If NTSCStr(CType(xlRange.Cells(nRiga, 1), Excel.Range).Value) <> "" Then
                    'riga nuova
                    dtrNuovaRiga = dtDati.NewRow()
                    dtrNuovaRiga!pa_integr = NTSCStr(CType(xlRange.Cells(nRiga, 1), Excel.Range).Value)
                    dtrNuovaRiga!pa_datreg = NTSCDate(CType(xlRange.Cells(nRiga, 3), Excel.Range).Value)
                    dtrNuovaRiga!pa_numreg = NTSCInt(CType(xlRange.Cells(nRiga, 4), Excel.Range).Value)
                    dtrNuovaRiga!pa_datdoc = NTSCDate(CType(xlRange.Cells(nRiga, 5), Excel.Range).Value)
                    dtrNuovaRiga!pa_numdoc = NTSCInt(CType(xlRange.Cells(nRiga, 6), Excel.Range).Value)
                    dtrNuovaRiga!pa_alfdoc = NTSCStr(CType(xlRange.Cells(nRiga, 7), Excel.Range).Value)
                    dtrNuovaRiga!pa_codart = NTSCStr(CType(xlRange.Cells(nRiga, 8), Excel.Range).Value)
                    dtrNuovaRiga!pa_desart = NTSCStr(CType(xlRange.Cells(nRiga, 9), Excel.Range).Value)
                    dtrNuovaRiga!pa_note = NTSCStr(CType(xlRange.Cells(nRiga, 10), Excel.Range).Value)
                    dtrNuovaRiga!pa_codlavo = NTSCInt(CType(xlRange.Cells(nRiga, 11), Excel.Range).Value)
                    dtrNuovaRiga!pa_codcena = NTSCInt(CType(xlRange.Cells(nRiga, 13), Excel.Range).Value)
                    dtrNuovaRiga!pa_unmis = NTSCStr(CType(xlRange.Cells(nRiga, 15), Excel.Range).Value)
                    dtrNuovaRiga!pa_quant = NTSCDec(CType(xlRange.Cells(nRiga, 16), Excel.Range).Value)
                    'compilato il dare
                    If NTSCStr(CType(xlRange.Cells(nRiga, 17), Excel.Range).Value).Trim <> "" Then
                        dtrNuovaRiga!pa_darave = "D"
                        dtrNuovaRiga!pa_importo = NTSCDec(CType(xlRange.Cells(nRiga, 17), Excel.Range).Value)
                        dtrNuovaRiga!pa_importoda = NTSCDec(CType(xlRange.Cells(nRiga, 17), Excel.Range).Value)
                    End If
                    'compilato l'avere
                    If NTSCStr(CType(xlRange.Cells(nRiga, 18), Excel.Range).Value).Trim <> "" Then
                        dtrNuovaRiga!pa_darave = "A"
                        dtrNuovaRiga!pa_importo = NTSCDec(CType(xlRange.Cells(nRiga, 18), Excel.Range).Value)
                        dtrNuovaRiga!pa_importoda = NTSCDec(CType(xlRange.Cells(nRiga, 18), Excel.Range).Value)
                    End If
                    dtrNuovaRiga!pa_codcfam = Right("0000" & NTSCStr(CType(xlRange.Cells(nRiga, 19), Excel.Range).Value), 4)
                    dtrNuovaRiga!pa_commeca = NTSCInt(CType(xlRange.Cells(nRiga, 21), Excel.Range).Value)
                    dtrNuovaRiga!pa_conto = NTSCInt(CType(xlRange.Cells(nRiga, 23), Excel.Range).Value)
                    dtrNuovaRiga!pa_controp = NTSCInt(CType(xlRange.Cells(nRiga, 25), Excel.Range).Value)
                    dtrNuovaRiga!pa_codcope = NTSCInt(CType(xlRange.Cells(nRiga, 27), Excel.Range).Value)
                    dtrNuovaRiga!pa_subcommeca = NTSCStr(CType(xlRange.Cells(nRiga, 29), Excel.Range).Value)
                    dtrNuovaRiga!pa_tipore = NTSCStr(CType(xlRange.Cells(nRiga, 30), Excel.Range).Value)
                    dtrNuovaRiga!pa_hhprogr = NTSCInt(CType(xlRange.Cells(nRiga, 2), Excel.Range).Value)

                    dtDati.Rows.Add(dtrNuovaRiga)
                    dtDati.AcceptChanges()

                End If
            Next



            xlWorkBook.Close(False)
            xlApp.Quit()


            Return True

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        Finally
            ScaricaOggettoExcel(xlApp)
            ScaricaOggettoExcel(xlWorkBook)
            ScaricaOggettoExcel(xlWorkSheet)
        End Try
    End Function


    Private Sub ScaricaOggettoExcel(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Overridable Function VerificaDati(ByVal dtFile As DataTable) As Boolean
        Dim dr As DataRow
        Dim dtTmp As New DataTable


        Try

            For Each dr In dtFile.Rows
                'verifico che l'articolo sia corretto
                If NTSCStr(dr!pa_codart).Trim <> "" Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_codart), strDittaCorrente, "artico", "", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Articolo " & NTSCStr(dr!pa_codart) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If
                'verifico che la lavorazione sia corretta
                If NTSCInt(dr!pa_codlavo) <> 0 Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_codlavo), strDittaCorrente, "TABLAVO", "", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Lavorazione " & NTSCStr(dr!pa_codlavo) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If
                'verifico che il centro sia corretto
                If NTSCInt(dr!pa_codcena) <> 0 Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_codcena), strDittaCorrente, "TABCENA", "", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Centro " & NTSCStr(dr!pa_codcena) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If
                'verifico che la famiglia sia corretta
                If NTSCStr(dr!pa_codcfam).Trim <> "" Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_codcfam), strDittaCorrente, "TABCFAM", "", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Famiglia " & NTSCStr(dr!pa_codcfam) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If
                'verifico che la commessa sia corretta
                If NTSCInt(dr!pa_commeca) <> 0 Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_commeca), strDittaCorrente, "commess", "", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Commessa " & NTSCStr(dr!pa_commeca) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If
                'verifico che il conto sia corretto
                If NTSCInt(dr!pa_conto) <> 0 Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_conto), strDittaCorrente, "anagca", "S", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Conto " & NTSCStr(dr!pa_conto) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If
                'verifico che la contropartita sia corretta
                If NTSCInt(dr!pa_controp) <> 0 Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_controp), strDittaCorrente, "anagra", "", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Contropartita " & NTSCStr(dr!pa_controp) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If
                'verifico che ll'operaio sia corretto
                If NTSCInt(dr!pa_codcope) <> 0 Then
                    oCldIMAN.ValCodiceDb(NTSCStr(dr!pa_codcope), strDittaCorrente, "TABCOPE", "", "", dtTmp)
                    If dtTmp.Rows.Count <= 0 Then
                        MsgBox("Operaio " & NTSCStr(dr!pa_codcope) & " inesistente in Business. Impossibile Continuare")
                        Return False
                    Else
                    End If
                End If

            Next


            Return True

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        Finally
        End Try
    End Function



#End Region


End Class
