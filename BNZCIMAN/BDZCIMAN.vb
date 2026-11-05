Imports System.Data.Common
Imports NTSInformatica.CLN__STD
Imports System.IO
Imports System.Drawing

Public Class CLDZCIMAN
  Inherits CLD__BASE

    Public dbConn As DbConnection = Nothing


    'connessione con l'altra istanza
    Public dbConnCrm As New SqlClient.SqlConnection
    Public bCnOpen As Boolean
    Public strCollate As String = " COLLATE SQL_Latin1_General_CP1_CI_AS "


    Public Overridable Function ApriConnCrm(ByVal strConnString As String) As Boolean
        Try
            'MsgBox(strConnString)
            dbConnCrm.ConnectionString = strConnString
            dbConnCrm.Open()
            bCnOpen = True

            Return True
        Catch ex As Exception
            oApp.MsgBoxInfo("ATTENZIONE!!! Impossibile accedere al database di scambio (" & strConnString & ").")
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------

        End Try
    End Function

    Public Overridable Function ChiudiConnMes() As Boolean
        Try
            If bCnOpen Then dbConnCrm.Close()
            bCnOpen = False

            Return True
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
    End Function



#Region "Materiali"
    'Public Overridable Function LeggiMaterialiTestata(ByVal stato As Integer, ByRef dt As DataTable, ByVal nomedb As String) As Boolean
    '    Dim strSql As String = ""
    '    Dim nProg As Integer = 0
    '    Dim dQta As Decimal = 0
    '    Dim dImporto As Decimal = 0
    '    'Dim strMsg As String = "BDZCIMAN"
    '    Dim strErr As String = ""

    '    Try

    '        strSql = "SELECT rapportino FROM [" & nomedb & "].[DBO].[materiali]  WHERE stato=" & stato & " GROUP BY rapportino"
    '        dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
    '        Return True


    '    Catch ex As Exception
    '        '--------------------------------------------------------------
    '        CLN__STD.GestErr(ex, Me, "")
    '        '--------------------------------------------------------------
    '    End Try

    'End Function

    'Public Overridable Function LeggiMaterialiTestata(ByRef dt As DataTable, ByVal nomedb As String) As Boolean
    '    Dim strSql As String = ""
    '    Dim nProg As Integer = 0
    '    Dim dQta As Decimal = 0
    '    Dim dImporto As Decimal = 0
    '    'Dim strMsg As String = "BDZCIMAN"
    '    Dim strErr As String = ""

    '    Try

    '        strSql = "SELECT rapportino, stato FROM [" & nomedb & "].[DBO].[materiali] GROUP BY rapportino, stato"
    '        dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
    '        Return True


    '    Catch ex As Exception
    '        '--------------------------------------------------------------
    '        CLN__STD.GestErr(ex, Me, "")
    '        '--------------------------------------------------------------
    '    End Try

    'End Function

    'Public Overridable Function LeggiMaterialiRighe(ByVal stato As Integer, ByVal rapportino As String, ByRef dt As DataTable, ByVal nomedb As String) As Boolean
    '    Dim strSql As String = ""
    '    Dim nProg As Integer = 0
    '    Dim dQta As Decimal = 0
    '    Dim dImporto As Decimal = 0
    '    'Dim strMsg As String = "BDZCIMAN"
    '    Dim strErr As String = ""

    '    Try

    '        strSql = "SELECT * FROM [" & nomedb & "].[DBO].[materiali]  WHERE stato=" & stato & " and rapportino= " & NTSCStr(rapportino)
    '        dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
    '        Return True


    '    Catch ex As Exception
    '        '--------------------------------------------------------------
    '        CLN__STD.GestErr(ex, Me, "")
    '        '--------------------------------------------------------------
    '    End Try

    'End Function

    'Public Overridable Function SettaMaterialeImportato(ByVal stato As Integer, ByVal rapportino As String, ByVal nomedb As String) As Boolean
    '    Dim strSql As String = ""
    '    Dim nProg As Integer = 0
    '    Dim dQta As Decimal = 0
    '    Dim dImporto As Decimal = 0
    '    'Dim strMsg As String = "BDZCIMAN"
    '    Dim strErr As String = ""

    '    Try

    '        strSql = "UPDATE [" & nomedb & "].[DBO].[materiali]  SET stato=5 WHERE stato=" & stato & " and rapportino= " & NTSCStr(rapportino)
    '        Execute(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
    '        Return True


    '    Catch ex As Exception
    '        '--------------------------------------------------------------
    '        CLN__STD.GestErr(ex, Me, "")
    '        '--------------------------------------------------------------
    '    End Try

    'End Function

    Public Overridable Function LeggiMateriali(ByRef dt As DataTable, ByVal nomedb As String) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            strSql = "SELECT * FROM [" & nomedb & "].[DBO].[materiali] "
            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
            Return True


        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    Public Overridable Function TrovaDDTRigaCancellataMateriale(ByVal codart As String, ByVal rapportino As String, ByVal commeca As Integer, ByVal ditta As String, ByRef dt As DataTable) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            strSql = "SELECT * FROM MOVMAG 
                        LEFT JOIN testmag 
                        ON movmag.codditt=testmag.codditt AND mm_tipork=tm_tipork AND mm_anno=tm_anno AND mm_serie=tm_serie AND mm_numdoc=tm_numdoc 
                        WHERE mm_codart=" & CStrSQL(codart) & " AND mm_hhrapp=" & CStrSQL(rapportino) & " and mm_commeca=" & commeca & " and movmag.codditt=" & CStrSQL(ditta)

            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI)
            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    Public Overridable Function TornaNumeroRigheDDT(ByVal tipork As String, ByVal anno As Integer, ByVal serie As String, ByVal numdoc As Integer, ByVal ditta As String, ByRef totrighe As Integer) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""
        Dim dt As New DataTable

        Try

            strSql = "SELECT count(*) as totrighe FROM MOVMAG 
                        LEFT JOIN testmag 
                        ON movmag.codditt=testmag.codditt AND mm_tipork=tm_tipork AND mm_anno=tm_anno AND mm_serie=tm_serie AND mm_numdoc=tm_numdoc 
                        WHERE mm_tipork=" & CStrSQL(tipork) & " AND mm_anno=" & anno & " AND mm_serie=" & CStrSQL(serie) & " and mm_numdoc=" & numdoc & " and movmag.codditt=" & CStrSQL(ditta)

            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI)

            totrighe = dt.Rows(0)!totrighe

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function



    Public Overridable Function CancRigaImportataMaterialeDBFrontiera(ByVal id As Integer, ByVal nomedb As String) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            strSql = "DELETE FROM [" & nomedb & "].[DBO].[materiali]  WHERE id= " & id
            Execute(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
            Return True


        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function


#End Region


#Region "Risorse"

    'Public Overridable Function LeggiRisorse(ByVal stato As Integer, ByRef dt As DataTable, ByVal nomedb As String) As Boolean
    '    Dim strSql As String = ""
    '    Dim nProg As Integer = 0
    '    Dim dQta As Decimal = 0
    '    Dim dImporto As Decimal = 0
    '    'Dim strMsg As String = "BDZCIMAN"
    '    Dim strErr As String = ""

    '    Try

    '        strSql = "Select * FROM [" & nomedb & "].[DBO].[risorse]  WHERE stato=" & stato
    '        dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
    '        Return True


    '    Catch ex As Exception
    '        '--------------------------------------------------------------
    '        CLN__STD.GestErr(ex, Me, "")
    '        '--------------------------------------------------------------
    '    End Try

    'End Function

    Public Overridable Function LeggiRisorse(ByRef dt As DataTable, ByVal nomedb As String) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            'strSql = "Select * FROM [" & nomedb & "].[DBO].[risorse]  WHERE importato='N' order by id"
            strSql = "SELECT * FROM [" & nomedb & "].[DBO].[risorse] order by id"
            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
            Return True


        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    Public Overridable Function InsRigaPrianaRisorse(ByVal strDitta As String, ByVal conto As Integer, ByVal dr As DataRow) As Boolean

        Dim strlog = ""

        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Dim dalle As String
        Dim alle As String

        Dim nDalle As Decimal
        Dim nAlle As Decimal
        Dim qtaOre As Decimal
        Dim costoUnitario As Decimal
        Dim Importo As Decimal


        Dim codcope As Integer
        Dim dttCope As New DataTable
        Dim codcena As Integer
        Dim codconto As Integer
        Dim mastro As Integer
        Dim tipoore As String


        Dim dttDatiDaOperaio As New DataTable


        Dim codcomme As Integer
        Dim dttComme As New DataTable


        Try
            'Ottengo un progressivo per priana
            nProg = LegNuma(strDitta, "PA", " ", 0, True)
            'Lo salvo e se torna un'errore lancio un'eccezione
            nProg = AggNuma(strDitta, "PA", " ", 0, nProg, True, False, strErr)
            If strErr <> "" Then Throw New NTSException(strErr)

            'converto dalle alle prima in decimal poi lo passo nelle funzioni per metterlo da base 60 a base 10
            dalle = Left(NTSCStr(dr!dalle), 2) & "," & Mid(NTSCStr(dr!dalle), 4, 2)
            alle = Left(NTSCStr(dr!alle), 2) & "," & Mid(NTSCStr(dr!alle), 4, 2)

            nDalle = NTSCDec(dalle)
            nAlle = NTSCDec(alle)

            nDalle = ConvOra60Ora100(nDalle, 2)
            nAlle = ConvOra60Ora100(nAlle, 2)

            'calcolo la qantità di ore in base 100 poi lo risputo fuori in base 60???
            qtaOre = nAlle - nDalle
            'qtaOre = ConvOra100Ora60(qtaOre)

            '14/09/2026 - ci passano la quantità delle ore pulita da eventuali pause
            qtaOre = NTSCDec(dr!totore)

            'perlevo l'operatore
            codcope = dr!codice
            ValCodiceDb(codcope, strDitta, "TABCOPE", "", "N", dttCope)

            If dttCope.Rows.Count >= 0 Then
                'prelevo i dati necessari dall'operatore
                TrovaDatiDaOperaioMezzo(strDitta, codcope, dttDatiDaOperaio)
                If dttDatiDaOperaio.Rows.Count >= 0 Then
                    'prelevo il centro
                    codcena = dttDatiDaOperaio.Rows(0)!codcena
                    'prelevo il conto (se non c'è metto quello da opzione di registro
                    If dttDatiDaOperaio.Rows(0)!codconto = 0 Then
                        codconto = conto
                    Else
                        codconto = dttDatiDaOperaio.Rows(0)!codconto
                    End If

                    'verifico che tipo di ore sono per moltiplicare le ore per il prezzo corretto presente nell'anagrafica dell'operaio 
                    Select Case dr!tipologia
                        Case 0
                            costoUnitario = dttDatiDaOperaio.Rows(0)!Cordinario
                            tipoore = "N"
                        Case 1
                            costoUnitario = dttDatiDaOperaio.Rows(0)!Cstraodinario
                            tipoore = "S"
                        Case 2
                            costoUnitario = dttDatiDaOperaio.Rows(0)!Cnotturno
                            tipoore = "O"
                        Case 3
                            costoUnitario = dttDatiDaOperaio.Rows(0)!Cfestivo
                            tipoore = "Y"
                        Case 4
                            costoUnitario = dttDatiDaOperaio.Rows(0)!Cfestivonotturno
                            tipoore = "G"
                    End Select

                    Importo = qtaOre * costoUnitario

                    'prelevo il mastro
                    mastro = dttDatiDaOperaio.Rows(0)!ac_codmaca

                End If
            Else
                strlog = "Operaio " & codcope & " non presente in business"
                Return True
            End If

            'prelevo la commessa e verifico che esista in business
            codcomme = dr!commessa
            'verifico se la commessa esiste in bus
            ValCodiceDb(codcomme, strDitta, "COMMESS", "N", "", dttComme)
            If dttComme.Rows.Count < 0 Then
                strlog = "Commessa " & codcomme & " non presente in business"
                Return True
            End If

            'Inserisco in priana se non esiste gia a parità di rapportino???? da verificare 
            strSql = "if not exists (SELECT * FROM priana WHERE codditt=" & CStrSQL(strDitta) & " AND pa_note=" & CStrSQL(NTSCInt(dr!rapportino)) & ") 
                        begin 
                     INSERT INTO priana 
                        (codditt,pa_progr,pa_escomp,pa_conto,pa_codcena,pa_commeca, pa_unmis, pa_quant, pa_darave, pa_importo, pa_importoda, pa_origine,pa_datreg,pa_note, pa_codcope, pa_daore, pa_aore,pa_mastro,
						pa_datini, pa_datfin, pa_quantda,pa_tipore)
                     VALUES(" & CStrSQL(strDitta) & "," & nProg & "," & Year(NTSCDate(dr!data)) & " , " & codconto & " , " & codcena & "," & codcomme & ",'H'," & CDblSQL(qtaOre) & ",
                     'D'," & CDblSQL(Importo) & " , " & CDblSQL(Importo) & " , 'O', " & CDataSQL(NTSCDate(dr!data)) & " , " & CStrSQL(NTSCStr(dr!rapportino)) & ", " & codcope & " , " & CDblSQL(nDalle) & "," & CDblSQL(nAlle) & "," & mastro & ",
                     " & CDataSQL(NTSCDate(dr!data)) & "," & CDataSQL(NTSCDate(dr!data)) & "," & CDblSQL(qtaOre) & "," & CStrSQL(tipoore) & ") end"
            Execute(strSql, CLE__APP.DBTIPO.DBAZI)

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function


    'Public Overridable Function CancRigaPrianaRisorse(ByVal strDitta As String, ByVal codcope As Integer, ByVal rapportino As String, ByVal data As Date, ByVal id As Integer) As Boolean
    Public Overridable Function CancRigaPrianaRisorse(ByVal strDitta As String, ByVal codcope As Integer, ByVal rapportino As String, ByVal data As Date, ByVal id As Integer, ByVal dalle As String, ByVal alle As String, ByVal commeca As Integer, ByVal tipologia As Integer) As Boolean


        Dim strlog = ""
        Dim strSql As String = ""
        Dim dt As New DataTable

        Dim TipoOre As String
        Dim ndalle As Double
        Dim nalle As Double


        Try

            Select Case tipologia
                Case 0
                    TipoOre = "N"
                Case 1
                    TipoOre = "S"
                Case 2
                    TipoOre = "O"
                Case 3
                    TipoOre = "Y"
                Case 4
                    TipoOre = "G"
            End Select

            dalle = Left(NTSCStr(dalle), 2) & "," & Mid(NTSCStr(dalle), 4, 2)
            alle = Left(NTSCStr(alle), 2) & "," & Mid(NTSCStr(alle), 4, 2)

            ndalle = NTSCDec(dalle)
            nalle = NTSCDec(alle)

            ndalle = ConvOra60Ora100(ndalle, 2)
            nalle = ConvOra60Ora100(nalle, 2)

            strSql = "SELECT * FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino) & " and pa_datreg=" & CDataOraSQL(data)
            strSql &= " and pa_daore =" & CDblSQL(ndalle) & " and pa_aore= " & CDblSQL(nalle) & " and pa_commeca=" & commeca & " and pa_tipore=" & CStrSQL(TipoOre)

            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI)

            'se trovo solo una riga allora cancello altrimenti no
            Select Case dt.Rows.Count
                Case 0
                    strlog += "ATTENZIONE!Non sono state trovate righe in prima nota corrispondenti a operaio " & codcope & ", rapportino " & rapportino & " data " & data & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
                Case 1
                    strSql = "DELETE FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino) & " and pa_datreg=" & CDataOraSQL(data)
                    strSql &= " and pa_daore =" & CDblSQL(ndalle) & " and pa_aore= " & CDblSQL(nalle) & " and pa_commeca=" & commeca & " and pa_tipore=" & CStrSQL(TipoOre)

                    Execute(strSql, CLE__APP.DBTIPO.DBAZI)
                Case > 1
                    strlog += "ATTENZIONE!Sono state trovate più righe in prima nota corrispondenti a operaio " & codcope & ", rapportino " & rapportino & " data " & data & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
            End Select

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    Public Overridable Function CancRigaPrianaRisorsedaMod(ByVal strDitta As String, ByVal codcope As Integer, ByVal rapportino As String, ByVal data As Date, ByVal id As Integer, ByVal dalle As String, ByVal alle As String, ByVal commeca As Integer, ByVal tipologia As Integer) As Boolean


        Dim strlog = ""
        Dim strSql As String = ""
        Dim dt As New DataTable

        Dim TipoOre As String
        Dim ndalle As Double
        Dim nalle As Double


        Try

            Select Case tipologia
                Case 0
                    TipoOre = "N"
                Case 1
                    TipoOre = "S"
                Case 2
                    TipoOre = "O"
                Case 3
                    TipoOre = "Y"
                Case 4
                    TipoOre = "G"
            End Select

            dalle = Left(NTSCStr(dalle), 2) & "," & Mid(NTSCStr(dalle), 4, 2)
            alle = Left(NTSCStr(alle), 2) & "," & Mid(NTSCStr(alle), 4, 2)

            ndalle = NTSCDec(dalle)
            nalle = NTSCDec(alle)

            ndalle = ConvOra60Ora100(ndalle, 2)
            nalle = ConvOra60Ora100(nalle, 2)

            strSql = "SELECT * FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino) & " and pa_datreg=" & CDataOraSQL(data)
            'strSql &= " and pa_daore =" & CDblSQL(ndalle) & " and pa_aore= " & CDblSQL(nalle) & " and pa_commeca=" & commeca & " and pa_tipore=" & CStrSQL(TipoOre)

            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI)

            strSql = "DELETE FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino) & " and pa_datreg=" & CDataOraSQL(data)
            'strSql &= " and pa_daore =" & CDblSQL(ndalle) & " and pa_aore= " & CDblSQL(nalle) & " and pa_commeca=" & commeca & " and pa_tipore=" & CStrSQL(TipoOre)
            Execute(strSql, CLE__APP.DBTIPO.DBAZI)

            ''se trovo solo una riga allora cancello altrimenti no
            'Select Case dt.Rows.Count
            '    Case 0
            '        strlog += "ATTENZIONE!Non sono state trovate righe in prima nota corrispondenti a operaio " & codcope & ", rapportino " & rapportino & " data " & data & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
            '    Case 1
            '        strSql = "DELETE FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino) & " and pa_datreg=" & CDataOraSQL(data)
            '        strSql &= " and pa_daore =" & CDblSQL(ndalle) & " and pa_aore= " & CDblSQL(nalle) & " and pa_commeca=" & commeca & " and pa_tipore=" & CStrSQL(TipoOre)


            '    Case > 1
            '        strlog += "ATTENZIONE!Sono state trovate più righe in prima nota corrispondenti a operaio " & codcope & ", rapportino " & rapportino & " data " & data & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
            'End Select

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function


    Public Overridable Function SettaRisorseImportato(ByVal id As Integer, ByVal nomedb As String) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            strSql = "UPDATE [" & nomedb & "].[DBO].[risorse]  SET importato='S' WHERE id=" & id
            Execute(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    'cancella riga importata da db di frontiera
    Public Overridable Function CancRigaImportataRisorseDaDBFrontiera(ByVal id As Integer, ByVal nomedb As String) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            strSql = "DELETE FROM [" & nomedb & "].[DBO].[risorse]  WHERE id=" & id
            Execute(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

#End Region


#Region "Mezzi"
    Public Overridable Function LeggiMezzi(ByRef dt As DataTable, ByVal nomedb As String) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            strSql = "SELECT * FROM [" & nomedb & "].[DBO].[mezzi] "
            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
            Return True


        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    Public Overridable Function InsRigaPrianaMezzi(ByVal strDitta As String, ByVal conto As Integer, ByVal dr As DataRow) As Boolean

        Dim strlog = ""

        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Dim qtaOre As Decimal
        Dim costoUnitario As Decimal
        Dim Importo As Decimal


        Dim codcope As Integer
        Dim dttCope As New DataTable
        Dim codcena As Integer
        Dim codconto As Integer
        Dim mastro As Integer
        Dim tipoore As String


        Dim dttDatiDaOperaio As New DataTable


        Dim codcomme As Integer
        Dim dttComme As New DataTable

        Dim stato As Integer

        'per test..poi va alla data odierna
        'Dim dataReg As String = "2025-10-31"
        'per reale
        Dim dataReg As String = Today.Date


        Try
            'Ottengo un progressivo per priana
            nProg = LegNuma(strDitta, "PA", " ", 0, True)
            'Lo salvo e se torna un'errore lancio un'eccezione
            nProg = AggNuma(strDitta, "PA", " ", 0, nProg, True, False, strErr)
            If strErr <> "" Then Throw New NTSException(strErr)

            'se è in inserimento o modifica prende un campo diverso
            'stato = dr!stato
            'Select Case stato
            '    Case 0
            '        qtaOre = ConvOra100Ora60(NTSCDec(dr!ore))
            '    Case 1
            '        qtaOre = ConvOra100Ora60(NTSCDec(dr!prevalue))
            'End Select

            ' qtaOre = ConvOra100Ora60(NTSCDec(dr!ore))
            qtaOre = (NTSCDec(dr!ore))

            'perlevo il mezzo
            codcope = dr!codice
            ValCodiceDb(codcope, strDitta, "TABCOPE", "", "N", dttCope)

            If dttCope.Rows.Count >= 0 Then
                'prelevo i dati necessari dall'operatore
                TrovaDatiDaOperaioMezzo(strDitta, codcope, dttDatiDaOperaio)
                If dttDatiDaOperaio.Rows.Count >= 0 Then
                    'prelevo il centro
                    codcena = dttDatiDaOperaio.Rows(0)!codcena
                    'prelevo il conto (se non c'è metto quello da opzione di registro
                    If dttDatiDaOperaio.Rows(0)!codconto = 0 Then
                        codconto = conto
                    Else
                        codconto = dttDatiDaOperaio.Rows(0)!codconto
                    End If

                    'setto il costo, il tipo ore, e  l'importo
                    costoUnitario = dttDatiDaOperaio.Rows(0)!Cordinario
                    tipoore = "N"
                    Importo = qtaOre * costoUnitario

                    'prelevo il mastro
                    mastro = dttDatiDaOperaio.Rows(0)!ac_codmaca

                End If
            Else
                strlog = "Operaio " & codcope & " non presente in business"
                Return True
            End If

            'prelevo la commessa e verifico che esista in business
            codcomme = dr!commessa
            'verifico se la commessa esiste in bus
            ValCodiceDb(codcomme, strDitta, "COMMESS", "N", "", dttComme)
            If dttComme.Rows.Count < 0 Then
                strlog = "Commessa " & codcomme & " non presente in business"
                Return True
            End If

            'Inserisco in priana se non esiste gia a parità di rapportino???? da verificare 
            strSql = "if not exists (SELECT * FROM priana WHERE codditt=" & CStrSQL(strDitta) & " AND pa_note=" & CStrSQL(NTSCInt(dr!rapportino)) & ") 
                        begin 
                     INSERT INTO priana 
                        (codditt,pa_progr,pa_escomp,pa_conto,pa_codcena,pa_commeca, pa_unmis, pa_quant, pa_darave, pa_importo, pa_importoda, pa_origine,pa_datreg,pa_note, pa_codcope,pa_mastro,
						pa_datini, pa_datfin, pa_quantda,pa_tipore)
                     VALUES(" & CStrSQL(strDitta) & "," & nProg & "," & Year(NTSCDate(dataReg)) & " , " & codconto & " , " & codcena & "," & codcomme & ",'H'," & CDblSQL(qtaOre) & ",
                     'D'," & CDblSQL(Importo) & " , " & CDblSQL(Importo) & " , 'O', " & CDataSQL(NTSCDate(dataReg)) & " , " & CStrSQL(NTSCStr(dr!rapportino)) & ", " & codcope & " , " & mastro & ",
                     " & CDataSQL(NTSCDate(dataReg)) & "," & CDataSQL(NTSCDate(dataReg)) & "," & CDblSQL(qtaOre) & "," & CStrSQL(tipoore) & ") end"
            Execute(strSql, CLE__APP.DBTIPO.DBAZI)

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    Public Overridable Function ModRigaPrianaMezzi(ByVal strDitta As String, ByVal codcope As Integer, ByVal rapportino As String, ByVal data As Date, ByVal id As Integer, ByVal qtaore As Decimal) As Boolean

        Dim strlog = ""
        Dim strSql As String = ""
        Dim dt As New DataTable

        Dim dttCope As New DataTable
        Dim dttDatiDaOperaio As New DataTable
        Dim costounitario As Decimal
        Dim importo As Decimal

        Try

            strSql = "SELECT * FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino) & " and pa_datreg=" & CDataOraSQL(data)
            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI)

            'se trovo solo una riga allora cancello altrimenti no
            Select Case dt.Rows.Count
                Case 0
                    strlog += "ATTENZIONE!Non sono state trovate righe in prima nota corrispondenti a mezzo " & codcope & ", rapportino " & rapportino & " data " & data & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
                Case 1

                    'prendo i dat
                    ValCodiceDb(codcope, strDitta, "TABCOPE", "", "N", dttCope)

                    If dttCope.Rows.Count >= 0 Then
                        'prelevo i dati necessari dall'operatore
                        TrovaDatiDaOperaioMezzo(strDitta, codcope, dttDatiDaOperaio)
                        If dttDatiDaOperaio.Rows.Count >= 0 Then
                            'setto il costo, il tipo ore, e  l'importo
                            costoUnitario = dttDatiDaOperaio.Rows(0)!Cordinario
                            importo = qtaore * costounitario
                        End If
                    Else
                        strlog = "Operaio " & codcope & " non presente in business"
                        Return True
                    End If

                    strSql = "UPDATE priana SET pa_importo=" & importo & ", pa_importoda=" & importo & ", pa_quant=" & qtaore & "  WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino) & " and pa_datreg=" & CDataOraSQL(data)
                    Execute(strSql, CLE__APP.DBTIPO.DBAZI)

                Case > 1
                    strlog += "ATTENZIONE!Sono state trovate più righe in prima nota corrispondenti a mezzo " & codcope & ", rapportino " & rapportino & " data " & data & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
            End Select

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function


    Public Overridable Function CancRigaPrianaMezzi(ByVal strDitta As String, ByVal codcope As Integer, ByVal rapportino As String, ByVal id As Integer, ByVal qta As Decimal) As Boolean

        Dim strlog = ""
        Dim strSql As String = ""
        Dim dt As New DataTable

        Try

            strSql = "SELECT * FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino)
            '& " and pa_quant= " & NTSCDec(qta)
            '& " and pa_datreg=" & CDataOraSQL(data)
            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI)

            ''se trovo solo una riga allora cancello altrimenti no
            'Select Case dt.Rows.Count
            '    Case 0
            '        strlog += "ATTENZIONE!Non sono state trovate righe in prima nota corrispondenti a mezzo " & codcope & ", rapportino " & rapportino & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
            '    Case 1
            '        strSql = "DELETE FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino)
            '        '& " and pa_datreg=" & CDataOraSQL(data)
            '        Execute(strSql, CLE__APP.DBTIPO.DBAZI)
            '    Case > 1
            '        strlog += "ATTENZIONE!Sono state trovate più righe in prima nota corrispondenti a mezzo " & codcope & ", rapportino " & rapportino & "[id tab Risrse=" & id & "]. Impossibile cancellare!"
            'End Select
            strSql = "DELETE FROM priana WHERE codditt= " & CStrSQL(strDitta) & " AND " & " pa_codcope=" & codcope & " AND pa_note=" & CStrSQL(rapportino)
            '        '& " and pa_datreg=" & CDataOraSQL(data)
            Execute(strSql, CLE__APP.DBTIPO.DBAZI)
            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function

    Public Overridable Function CancRigaImportataMezziDaDBFrontiera(ByVal id As Integer, ByVal nomedb As String) As Boolean
        Dim strSql As String = ""
        Dim nProg As Integer = 0
        Dim dQta As Decimal = 0
        Dim dImporto As Decimal = 0
        'Dim strMsg As String = "BDZCIMAN"
        Dim strErr As String = ""

        Try

            strSql = "DELETE FROM [" & nomedb & "].[DBO].[mezzi]  WHERE id=" & id
            Execute(strSql, CLE__APP.DBTIPO.DBAZI, dbConnCrm)
            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function




#End Region

    Public Overridable Function TrovaDatiDaOperaioMezzo(ByVal ditta As String, ByVal codcope As Integer, ByRef dt As DataTable) As Boolean
        Dim strSql As String = ""

        Try

            strSql = " select tabcope.tb_ocodcena as codcena, tb_concova as codconto,
                        tabcope.tb_ocosora as Cordinario,
                        tabcope.tb_ocosoras as Cstraodinario,
                        tabcope.tb_ocosorao as Cnotturno,
                        tabcope.tb_ocosoray as Cfestivo,
                        tabcope.tb_ocosorag as Cfestivonotturno,
                        ac_codmaca as mastro, 
                        * from tabcope
                        left join tabcove on tabcope.tb_ocodcontr=tabcove.tb_codcove and tabcove.codditt=tabcope.codditt 
                        left join anagca on tabcove.tb_concova=ac_conto and tabcove.codditt=anagca.codditt
                        where tb_codcope=" & codcope & " and tabcope.codditt=" & CStrSQL(ditta)
            dt = OpenRecordset(strSql, CLE__APP.DBTIPO.DBAZI)
            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function




End Class
