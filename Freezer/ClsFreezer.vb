Imports log4net
Imports System.Data.SqlClient
Imports ComPlusInterface
Imports Utility

''' <summary>
''' Classe che implementa le interfacce rese disponibili dal servizio
''' </summary>
Public Class ClsFreezer
    Private Shared ReadOnly Log As ILog = LogManager.GetLogger(GetType(ClsFreezer))
    'Public m_objHashTable As Hashtable

    'Public Overloads Sub InizializeObject(ByVal objHashTable As Hashtable)
    '    m_objHashTable = objHashTable
    'End Sub

    Public Function ViewCodaCalcolo(StringConnectionICI As String, IdEnte As String) As String
        Dim myDataSet As New DataSet
        Dim myRet As String = "Non ci sono elaborazioni in corso".ToUpper

        Try
            Dim FncDB As New ClsDBManager
            myDataSet = FncDB.ViewCalcolo(StringConnectionICI, IdEnte, "", True)
            For Each myRow As DataRow In myDataSet.Tables(0).Rows
                myRet = Utility.StringOperation.FormatString(myRow("esito"))
            Next
        Catch ex As Exception
            Log.Error("ViewCodaCalcolo::si è verificato il seguente errore::" & ex.Message)
        End Try
        Return myRet
    End Function

    Public Function ViewCalcolo(StringConnectionICI As String, IdEnte As String) As DataSet
        Dim myRet As New DataSet

        Try
            Dim FncDB As New ClsDBManager
            myRet = FncDB.ViewCalcolo(StringConnectionICI, IdEnte, "", False)
        Catch ex As Exception
            Log.Error("ViewCalcolo::si è verificato il seguente errore::" & ex.Message)
        End Try
        Return myRet
    End Function
    ''' <summary>
    ''' Riordino per abitazione principale per poter gestire correttamente la detrazione che nel caso non copra tutta l’abitazione principale deve andare anche sulla pertinenza; ciclo su ogni riga in ingresso e valorizzo la struttura dei parametri di calcolo; valorizzo gli importi. Salvataggio valori nel dataset passato come parametro.
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <param name="ListSituazioneFinale"></param>
    ''' <returns></returns>
    Public Function getCalcolo(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, ByVal TipoCalcolo As Integer, ByVal ListSituazioneFinale As objSituazioneFinale()) As objSituazioneFinale()
        Dim cmdMyCommand As New SqlCommand
        Log.Debug("si calcola")
        Try
            Dim fncCalc As New CALCOLO_ICI
            Dim AnnoCalcolo As Integer
            Dim objUtility As New Generale
            Dim myParamICI As New Generale.PARAMETRI_ICI
            Dim oImportiICI As New Generale.VALORI_ICI_CALCOLATA
            Dim IsEsente As Integer
            Dim blnEsenzione As Boolean
            Dim mesiEsenzione As Integer
            Dim oDatiPerCalcolo As New objCalcolo
            Dim oDatiCalcolati As New objCalcolo
            Dim ListResiduoRendita() As ObjResiduoRendita = Nothing
            Dim impDetrazioneResiduaAcc As Double = 0
            Dim impDetrazResiduaStandardAcc As Double = 0
            Dim impDetrazioneResiduaTot As Double = 0
            Dim impDetrazResiduaStandardTot As Double = 0
            Dim ContribPrec As Integer = 0
            Dim CodCatAAP As String = ""

            'Valorizzo la connessione
            cmdMyCommand.Connection = New SqlClient.SqlConnection(StringConnectionICI)
            cmdMyCommand.CommandTimeout = 0
            If cmdMyCommand.Connection.State = ConnectionState.Closed Then
                cmdMyCommand.Connection.Open()
            End If

            Log.Debug("ho trovato queste righe::" & ListSituazioneFinale.GetUpperBound(0).ToString)
            'riordino per abitazione principale
            Array.Sort(ListSituazioneFinale, New Utility.Comparatore(New String() {"FlagPrincipale"}, New Boolean() {Utility.TipoOrdinamento.Crescente}))
            For Each mySituazioneFinale As objSituazioneFinale In ListSituazioneFinale
                Log.Debug("sono su:" + mySituazioneFinale.Foglio + "|" + mySituazioneFinale.Numero + "|" + mySituazioneFinale.Subalterno + " esenzione->" + mySituazioneFinale.FlagEsente.ToString())
                If mySituazioneFinale.IdContribuente <> ContribPrec Then
                    impDetrazioneResiduaAcc = 0 : impDetrazioneResiduaTot = 0
                    impDetrazResiduaStandardAcc = 0 : impDetrazResiduaStandardTot = 0
                End If
                AnnoCalcolo = mySituazioneFinale.Anno
                blnEsenzione = mySituazioneFinale.FlagEsente
                If (mySituazioneFinale.FlagEsente = "0") Then
                    Log.Debug("ho FLAG_ESENTE")
                    blnEsenzione = True
                    mesiEsenzione = 12
                Else
                    blnEsenzione = False
                    mesiEsenzione = 0
                End If
                If Year(mySituazioneFinale.Dal) = AnnoCalcolo Then
                    If (mySituazioneFinale.MesiEsenzione > 0) Then
                        Log.Debug("ho MESI_ESCL_ESENZIONE")
                        mesiEsenzione = mySituazioneFinale.MesiEsenzione
                    End If
                End If
                myParamICI.strCATEGORIA = mySituazioneFinale.Categoria
                myParamICI.intANNO_CALCOLO = AnnoCalcolo
                myParamICI.strTIPO_RENDITA = mySituazioneFinale.TipoRendita.ToUpper
                myParamICI.intTIPO_ABITAZIONE = mySituazioneFinale.FlagPrincipale
                '*** 20140509 - TASI ***
                myParamICI.IdTipoUtilizzo = mySituazioneFinale.IdTipoUtilizzo
                '*** ***
                '*** 20120530 - IMU ***
                myParamICI.IsColtivatoreDiretto = mySituazioneFinale.IsColtivatoreDiretto
                '*** ***
                '*** 201805 - se la pertinenza è riferita ad una principale esclusa devo esludere anche lei ***
                If mySituazioneFinale.FlagPrincipale = Generale.ABITAZIONE_PRINCIPALE_PERTINENZA.ABITAZIONE_PRINCIPALE Then
                    CodCatAAP = mySituazioneFinale.Categoria
                End If
                myParamICI.Categoria_AAP = CodCatAAP
                '*** ***
                '******************************************************************************************************
                'CALCOLO ICI ACCONTO
                '******************************************************************************************************
                myParamICI.strACCONTO_TOTALE = Generale.ACCONTO
                oDatiPerCalcolo = SetDatiPerCalcolo(StringConnectionGOV, StringConnectionICI, IdEnte, mySituazioneFinale, myParamICI, IsEsente, ListResiduoRendita)
                If oDatiPerCalcolo Is Nothing Then
                    Throw New Exception("COMPlusOPENgovProvvedimenti.COMPLUSBaseCalcoloICI.getCalcolo.Errore durante la fase::SetDatiPerCalcolo")
                End If
                oDatiCalcolati = fncCalc.CalcolaICI(oDatiPerCalcolo, blnEsenzione, mesiEsenzione, impDetrazioneResiduaAcc, impDetrazResiduaStandardAcc)
                If IsEsente = 1 Then
                    oImportiICI = SetICIImporti(0, oDatiCalcolati, oImportiICI)
                Else
                    oImportiICI = SetICIImporti(1, oDatiCalcolati, oImportiICI)
                End If
                '******************************************************************************************************
                'FINE CALCOLO ICI ACCONTO
                '******************************************************************************************************

                '******************************************************************************************************
                'TOTALE
                '******************************************************************************************************
                myParamICI.strACCONTO_TOTALE = Generale.TOTALE

                oDatiPerCalcolo = SetDatiPerCalcolo(StringConnectionGOV, StringConnectionICI, IdEnte, mySituazioneFinale, myParamICI, IsEsente, ListResiduoRendita)
                oDatiCalcolati = fncCalc.CalcolaICI(oDatiPerCalcolo, blnEsenzione, mesiEsenzione, impDetrazioneResiduaTot, impDetrazResiduaStandardTot)
                'Log.Debug("COMPlusOPENgovProvvedimenti.COMPLUSBaseCalcoloICI.getCalcolo.calcolato importi saldo")
                If IsEsente = 1 Then
                    oImportiICI = SetICIImporti(0, oDatiCalcolati, oImportiICI)
                Else
                    oImportiICI = SetICIImporti(2, oDatiCalcolati, oImportiICI)
                End If

                '***************************************************************************************************************
                'SALVATAGGIO VALORI NEL DATASET PASSATO COME PARAMETRO
                '***************************************************************************************************************
                setDS_TabellaSituazioneFinaleICI(mySituazioneFinale, oImportiICI, myParamICI.IsColtivatoreDiretto, oDatiPerCalcolo.FigliACarico, oDatiPerCalcolo.PercentCaricoFigli)
                '***************************************************************************************************************
                'FINE SALVATAGGIO VALORI NEL DATASET PASSATO COME PARAMETRO
                '***************************************************************************************************************
                ContribPrec = mySituazioneFinale.IdContribuente
            Next
            '*** 20121203 - IMU calcolo dovuto al netto del versato ***
            'controllo se devo fare il calcolo al netto del versato
            If TipoCalcolo = Generale.NETTOVERSATO Then
                ListSituazioneFinale = CalcoloNettoVersato(StringConnectionICI, ListSituazioneFinale)
                If IsNothing(ListSituazioneFinale) Then
                    Throw New Exception("COMPlusOPENgovProvvedimenti.COMPLUSBaseCalcoloICI.getCalcolo.Errore durante la fase di calcolo netto versato")
                End If
            End If
            '*** ***
            Return ListSituazioneFinale
        Catch ex As Exception
            Log.Error("COMPlusOPENgovProvvedimenti.COMPLUSBaseCalcoloICI.getCalcolo.Errore->", ex)
            Throw New Exception("COMPlusOPENgovProvvedimenti.COMPLUSBaseCalcoloICI.getCalcolo.Errore durante la fase di calcolo dell'ICI")
        Finally
            cmdMyCommand.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' Funzione per il calcolo degli importi dovuti sugli immobili dei soggetti in ingresso
    ''' </summary>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="bIsMassivo"></param>
    ''' <param name="blnConfigurazioneDich"></param>
    ''' <param name="blnRibaltaVersatoSuDovuto"></param>
    ''' <param name="blnCalcolaArrotondamento"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <param name="sTipoTASI"></param>
    ''' <param name="ListSituazioneFinale"></param>
    ''' <returns></returns>
    Public Function CalcoloICIcompletoAsync(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, nIdContribuente As Integer, TributoCalcolo As String, Tributo As String, AnnoDa As String, AnnoA As String, ByVal bIsMassivo As Boolean, ByVal blnConfigurazioneDich As Boolean, ByVal blnRibaltaVersatoSuDovuto As Boolean, ByVal blnCalcolaArrotondamento As Boolean, ByVal TipoCalcolo As Integer, sTipoTASI As String, TASIAProprietario As String, TipoOperazione As String, Operatore As String, ByRef ListSituazioneFinale() As objSituazioneFinale) As Boolean
        Try
            'CALCOLO ICI 
            Return CalcoloICIcompleto(StringConnectionGOV, StringConnectionICI, IdEnte, nIdContribuente, TributoCalcolo, Tributo, AnnoDa, AnnoA, bIsMassivo, blnConfigurazioneDich, blnRibaltaVersatoSuDovuto, blnCalcolaArrotondamento, TipoCalcolo, sTipoTASI, TASIAProprietario, TipoOperazione, Operatore, ListSituazioneFinale)
        Catch ex As Exception
            Throw New Exception("Function::CalcoloICIcompletoAsync::COMPlusBussinesObject" & "::" & " " & ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="ListUI"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <returns></returns>
    Public Function CalcoloICISingle(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, ListUI As ArrayList, ByVal TipoCalcolo As Integer) As objSituazioneFinale()
        Log.Debug("parte CalcoloICISingle")
        Try
            Dim ListToCalc() As objSituazioneFinale = Nothing
            Dim nList As Integer = 0
            Log.Debug("dichiarato listotcalc")
            For x = 0 To ListUI.Count - 1 Step 91
                Log.Debug("ciclo listui")
                Dim myCalc As New objSituazioneFinale
                myCalc.Tributo = ListUI(x)
                x += 1
                myCalc.IdEnte = ListUI(x)
                x += 1
                myCalc.Anno = ListUI(x)
                x += 1
                myCalc.TipoRendita = ListUI(x)
                x += 1
                myCalc.Categoria = ListUI(x)
                x += 1
                myCalc.Classe = ListUI(x)
                x += 1
                myCalc.Zona = ListUI(x)
                x += 1
                myCalc.Foglio = ListUI(x)
                x += 1
                myCalc.Numero = ListUI(x)
                x += 1
                myCalc.Subalterno = ListUI(x)
                x += 1
                myCalc.Provenienza = ListUI(x)
                x += 1
                myCalc.Caratteristica = ListUI(x)
                x += 1
                myCalc.Via = ListUI(x)
                x += 1
                myCalc.NCivico = ListUI(x)
                x += 1
                myCalc.Esponente = ListUI(x)
                x += 1
                myCalc.Scala = ListUI(x)
                x += 1
                myCalc.Interno = ListUI(x)
                x += 1
                myCalc.Piano = ListUI(x)
                x += 1
                myCalc.Barrato = ListUI(x)
                x += 1
                myCalc.Sezione = ListUI(x)
                x += 1
                myCalc.Protocollo = ListUI(x)
                x += 1
                myCalc.DataScadenza = ListUI(x)
                x += 1
                myCalc.DataInizio = ListUI(x)
                x += 1
                myCalc.TipoOperazione = ListUI(x)
                x += 1
                myCalc.TitPossesso = ListUI(x)
                x += 1
                myCalc.Id = ListUI(x)
                x += 1
                myCalc.IdContribuente = ListUI(x)
                x += 1
                myCalc.IdContribuenteCalcolo = ListUI(x)
                x += 1
                myCalc.IdProcedimento = ListUI(x)
                x += 1
                myCalc.IdRiferimento = ListUI(x)
                x += 1
                myCalc.IdLegame = ListUI(x)
                x += 1
                myCalc.Progressivo = ListUI(x)
                x += 1
                myCalc.IdVia = ListUI(x)
                x += 1
                myCalc.NumeroFigli = ListUI(x)
                x += 1
                myCalc.MesiPossesso = ListUI(x)
                x += 1
                myCalc.Mesi = ListUI(x)
                x += 1
                myCalc.IdTipoUtilizzo = ListUI(x)
                x += 1
                myCalc.IdTipoPossesso = ListUI(x)
                x += 1
                myCalc.NUtilizzatori = ListUI(x)
                x += 1
                myCalc.FlagPrincipale = ListUI(x)
                x += 1
                myCalc.FlagRiduzione = ListUI(x)
                x += 1
                myCalc.FlagEsente = ListUI(x)
                x += 1
                myCalc.FlagStorico = ListUI(x)
                x += 1
                myCalc.FlagPosseduto = ListUI(x)
                x += 1
                myCalc.FlagProvvisorio = ListUI(x)
                x += 1
                myCalc.MesiRiduzione = ListUI(x)
                x += 1
                myCalc.MesiEsenzione = ListUI(x)
                x += 1
                myCalc.AccMesi = ListUI(x)
                x += 1
                myCalc.IdImmobile = ListUI(x)
                x += 1
                myCalc.IdImmobilePertinenza = ListUI(x)
                x += 1
                myCalc.IdImmobileDichiarato = ListUI(x)
                x += 1
                myCalc.MeseInizio = ListUI(x)
                x += 1
                myCalc.AbitazionePrincipaleAttuale = ListUI(x)
                x += 1
                myCalc.AccSenzaDetrazione = ListUI(x)
                x += 1
                myCalc.AccDetrazioneApplicata = ListUI(x)
                x += 1
                myCalc.AccDovuto = ListUI(x)
                x += 1
                myCalc.AccDetrazioneResidua = ListUI(x)
                x += 1
                myCalc.SalSenzaDetrazione = ListUI(x)
                x += 1
                myCalc.SalDetrazioneApplicata = ListUI(x)
                x += 1
                myCalc.SalDovuto = ListUI(x)
                x += 1
                myCalc.SalDetrazioneResidua = ListUI(x)
                x += 1
                myCalc.TotSenzaDetrazione = ListUI(x)
                x += 1
                myCalc.TotDetrazioneApplicata = ListUI(x)
                x += 1
                myCalc.TotDovuto = ListUI(x)
                x += 1
                myCalc.TotDetrazioneResidua = ListUI(x)
                x += 1
                myCalc.IdAliquota = ListUI(x)
                x += 1
                myCalc.Aliquota = ListUI(x)
                x += 1
                myCalc.AliquotaStatale = ListUI(x)
                x += 1
                myCalc.PercentCaricoFigli = ListUI(x)
                x += 1
                myCalc.AccDovutoStatale = ListUI(x)
                x += 1
                myCalc.AccDetrazioneApplicataStatale = ListUI(x)
                x += 1
                myCalc.AccDetrazioneResiduaStatale = ListUI(x)
                x += 1
                myCalc.SalDovutoStatale = ListUI(x)
                x += 1
                myCalc.SalDetrazioneApplicataStatale = ListUI(x)
                x += 1
                myCalc.SalDetrazioneResiduaStatale = ListUI(x)
                x += 1
                myCalc.TotDovutoStatale = ListUI(x)
                x += 1
                myCalc.TotDetrazioneApplicataStatale = ListUI(x)
                x += 1
                myCalc.TotDetrazioneResiduaStatale = ListUI(x)
                x += 1
                myCalc.PercPossesso = ListUI(x)
                x += 1
                myCalc.Rendita = ListUI(x)
                x += 1
                myCalc.Valore = ListUI(x)
                x += 1
                myCalc.ValoreReale = ListUI(x)
                x += 1
                myCalc.Consistenza = ListUI(x)
                x += 1
                myCalc.ImpDetrazione = ListUI(x)
                x += 1
                myCalc.DiffImposta = ListUI(x)
                x += 1
                myCalc.Totale = ListUI(x)
                x += 1
                myCalc.IsColtivatoreDiretto = ListUI(x)
                x += 1
                myCalc.Dal = ListUI(x)
                x += 1
                myCalc.Al = ListUI(x)
                x += 1
                myCalc.TipoTasi = ListUI(x)
                x += 1
                myCalc.DescrTipoTasi = ListUI(x)
                'ricalcolo i mesi
                myCalc.Mesi = New Generale().mesi_possesso(myCalc.Dal, myCalc.Al, myCalc.Anno, 0, 0)
                Log.Debug("da arraylist a oggetto")
                ReDim Preserve ListToCalc(nList)
                ListToCalc(nList) = myCalc
                nList += 1
                Log.Debug("valorizzato listtocalc")
            Next
            Log.Debug("richiamo getCalcolo")
            'CALCOLO ICI 
            Dim ListCalc() As objSituazioneFinale = getCalcolo(StringConnectionGOV, StringConnectionICI, IdEnte, TipoCalcolo, ListToCalc)
            Return ListCalc
        Catch ex As Exception
            Log.Debug("CalcoloICISingle.Errore->", ex)
            Throw New Exception("Function::CalcoloICISingle::COMPlusBussinesObject" & "::" & " " & ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Se il calcolo è massivo viene valorizzata la tabella dell'elaborazione.
    ''' Prima di tutto viene svuotata la banca dati dal calcolo precedente.
    ''' Viene prelevato l'elenco dei contribuenti per i quali calcolare; viene richiamata la funzione per memorizzare in una tabella di appoggio i dati principali delle unità immobiliari per il calcolo.
    ''' Ciclando su tutti i record delle anagrafiche ottenuti si richiama il calcolo degli importi; inserisco nel database il calcolo per singolo immobile ed il riepilogo per contribuente+anno+tributo+tipotasi; se previsto ribalto versato nel dovuto. Se il calcolo è massivo viene aggiornata la tabella dell'elaborazione.
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="nIdContribuente"></param>
    ''' <param name="TributoCalcolo"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="AnnoDa"></param>
    ''' <param name="AnnoA"></param>
    ''' <param name="bIsMassivo"></param>
    ''' <param name="blnConfigurazioneDich"></param>
    ''' <param name="blnRibaltaVersatoSuDovuto"></param>
    ''' <param name="blnCalcolaArrotondamento"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <param name="sTipoTASI"></param>
    ''' <param name="TASIAProprietario"></param>
    ''' <param name="TipoOperazione"></param>
    ''' <param name="Operatore"></param>
    ''' <param name="dsICIfinale"></param>
    ''' <returns></returns>
    ''' <revisionHistory><revision date="08/10/2020">Inserito controllo, in caso di elaborazione massiva, della presenza di un calcolo in corso. Se è presente un calcolo in corso esco senza dare errori.</revision></revisionHistory>
    Private Function CalcoloICIcompleto(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, nIdContribuente As Integer, TributoCalcolo As String, Tributo As String, AnnoDa As String, AnnoA As String, ByVal bIsMassivo As Boolean, ByVal blnConfigurazioneDich As Boolean, ByVal blnRibaltaVersatoSuDovuto As Boolean, ByVal blnCalcolaArrotondamento As Boolean, ByVal TipoCalcolo As Integer, sTipoTASI As String, TASIAProprietario As String, TipoOperazione As String, Operatore As String, ByRef dsICIfinale As objSituazioneFinale()) As Boolean
        Dim objDSAnagrafica As New DataSet
        Dim lngID_ELABORAZIONE, ID_TASK_REPOSITORY As Long
        Dim objDBOPENgovProvvedimentiUpdate As New ClsDBManager
        Dim DBselect As New ClsDBManager
        Dim iRetValFreezer As Boolean
        Dim iRetValSaveCalcoloICI As Long
        Dim dblSumImportoVersato As Double
        Dim sElabInCorso As String = ""

        Try
            Log.Debug(IdEnte & " - INIZIO Calcolo Massivo - Tributo::" & TributoCalcolo)
            lngID_ELABORAZIONE = -1

            If bIsMassivo = True Then
                sElabInCorso = ViewCodaCalcolo(StringConnectionICI, IdEnte)
                If sElabInCorso.ToUpper <> "ELABORAZIONE IN CORSO" Then
                    lngID_ELABORAZIONE = DBselect.getNewID(StringConnectionICI, "ID_ELABORAZIONE")
                    ID_TASK_REPOSITORY = DBselect.getNewID(StringConnectionICI, "ID_TASK_REPOSITORY")
                    iRetValSaveCalcoloICI = objDBOPENgovProvvedimentiUpdate.Set_TP_TASK_REPOSITORY(StringConnectionICI, ID_TASK_REPOSITORY, lngID_ELABORAZIONE, "C", "Elaborazione in Corso", 0, IdEnte, AnnoDa, Operatore)
                    nIdContribuente = 0
                Else
                    Log.Debug("ClsFreezer.CalcoloICIcompleto.NON calcolo perchè già in corso")
                    Return True
                End If
            End If
            Log.Debug("ClsFreezer.CalcoloICIcompleto.svuoto DB")
            objDBOPENgovProvvedimentiUpdate.DeleteFreezer(StringConnectionICI, nIdContribuente, IdEnte, "")
            objDBOPENgovProvvedimentiUpdate.Delete_SITUAZIONE_FINALE_ICI(StringConnectionICI, AnnoDa, TributoCalcolo, IdEnte, nIdContribuente)
            objDBOPENgovProvvedimentiUpdate.Delete_TP_CALCOLO_FINALE_ICI(StringConnectionICI, AnnoDa, TributoCalcolo, IdEnte, nIdContribuente)
            objDSAnagrafica = GetAnagraficheFreezer(StringConnectionICI, AnnoDa, Tributo, IdEnte, nIdContribuente)

            If Not objDSAnagrafica Is Nothing Then
                'freezer massivo
                'viene passato un dataset di anagrafiche selezionate
                iRetValFreezer = CreateFreezer(StringConnectionICI, IdEnte, TributoCalcolo, AnnoDa, objDSAnagrafica, blnConfigurazioneDich, sTipoTASI, TASIAProprietario)
                If iRetValFreezer = True Then
                    'calcolo ici massivo
                    Dim ii As Integer
                    For ii = 0 To objDSAnagrafica.Tables(0).Rows.Count - 1
                        Log.Debug(IdEnte & " - Calcolo ICI per Contribuente " & objDSAnagrafica.Tables(0).Rows(ii)("COD_CONTRIBUENTE").ToString() & " - " & vbTab & ii + 1 & " di " & objDSAnagrafica.Tables(0).Rows.Count)
                        dsICIfinale = Calcola(StringConnectionGOV, StringConnectionICI, IdEnte, TributoCalcolo, AnnoDa, AnnoA, objDSAnagrafica.Tables(0).Rows(ii)("COD_CONTRIBUENTE").ToString(), TipoCalcolo, TipoOperazione)

                        'insert into TP_SITUAZIONE_FINALE_ICI
                        iRetValSaveCalcoloICI = objDBOPENgovProvvedimentiUpdate.Set_SITUAZIONE_FINALE(StringConnectionICI, IdEnte, dsICIfinale, lngID_ELABORAZIONE, Operatore)
                        If iRetValSaveCalcoloICI > 0 Then
                            'insert into TP_CALCOLO_FINALE_ICI
                            iRetValSaveCalcoloICI = objDBOPENgovProvvedimentiUpdate.Set_TP_CALCOLO_FINALE_ICI(StringConnectionICI, dsICIfinale, lngID_ELABORAZIONE, blnCalcolaArrotondamento)
                            If blnRibaltaVersatoSuDovuto = True Then
                                If iRetValSaveCalcoloICI > 0 Then
                                    'ribalto versato nel dovuto
                                    dblSumImportoVersato = DBselect.GetImportoVersatoPerCalcoloICI(StringConnectionICI, dsICIfinale(0), IdEnte)
                                    If dblSumImportoVersato > 0 Then
                                        iRetValSaveCalcoloICI = objDBOPENgovProvvedimentiUpdate.Set_RibaltaVersatoNelDovuto(StringConnectionICI, dsICIfinale(0), dblSumImportoVersato)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    Log.Debug(IdEnte & " - Calcolo ICI Terminato Correttamente")
                    If bIsMassivo = True Then
                        iRetValSaveCalcoloICI = objDBOPENgovProvvedimentiUpdate.Set_TP_TASK_REPOSITORY(StringConnectionICI, ID_TASK_REPOSITORY, lngID_ELABORAZIONE, "C", "Calcolo ICI Massivo " & AnnoDa, ii, IdEnte, AnnoDa, Operatore)
                    End If
                    Return True
                End If
            Else
                If bIsMassivo = True Then
                    iRetValSaveCalcoloICI = objDBOPENgovProvvedimentiUpdate.Set_TP_TASK_REPOSITORY(StringConnectionICI, ID_TASK_REPOSITORY, lngID_ELABORAZIONE, "C", "Elaborazione " & AnnoDa & " terminata con errori", 0, IdEnte, AnnoDa, Operatore)
                End If
                Throw New Exception("00000")
            End If
        Catch ex As Exception
            Log.Error(IdEnte & " - Function::CalcoloICImassivo::COMPlusCalcoloICI::CalcoloICIcompletoAsync::" & ex.Message)
            If bIsMassivo = True Then
                iRetValSaveCalcoloICI = objDBOPENgovProvvedimentiUpdate.Set_TP_TASK_REPOSITORY(StringConnectionICI, ID_TASK_REPOSITORY, lngID_ELABORAZIONE, "C", "Elaborazione " & AnnoDa & " terminata con errori", 0, IdEnte, AnnoDa, Operatore)
            End If
            Throw New Exception("Function::CalcoloICImassivo::COMPlusCalcoloICI::CalcoloICIcompletoAsync::" & ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Prelevo l'elenco dei contribuenti per i quali calcolare
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="ANNO"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="ENTE"></param>
    ''' <param name="CONTRIB"></param>
    ''' <returns></returns>
    Public Function GetAnagraficheFreezer(myStringConnection As String, ByVal ANNO As String, ByVal Tributo As String, ByVal ENTE As String, ByVal CONTRIB As Integer) As DataSet
        Try
            GetAnagraficheFreezer = Nothing
            'Dim ListAnagrafica As ListAnagrafica= GetListAnagragraficaFreezer(myStringConnection, ANNO, Tributo, ENTE, CONTRIB)
            Dim objListaAnagrafica As New ListAnagrafica
            Dim objAnagrafica As New ClsDBManager

            objListaAnagrafica = objAnagrafica.getListaContribuentiFreezer(myStringConnection, ANNO, Tributo, ENTE, CONTRIB)

            GetAnagraficheFreezer = objListaAnagrafica.p_dsItemsANAGRAFICA
        Catch ex As Exception
            Log.Error("Function::GetAnagraficheFreezer::COMPlusFreezer:: " & ex.Message)
            Throw New Exception("Function::GetAnagraficheFreezer::COMPlusFreezer:: " & ex.Message)
        End Try
    End Function
    '''' <summary>
    '''' Prelevo l'elenco dei contribuenti per i quali calcolare
    '''' </summary>
    '''' <param name="myStringConnection"></param>
    '''' <param name="ANNO"></param>
    '''' <param name="Tributo"></param>
    '''' <param name="ENTE"></param>
    '''' <param name="CONTRIB"></param>
    '''' <returns></returns>
    'Public Function GetListAnagragraficaFreezer(myStringConnection As String, ByVal ANNO As String, ByVal Tributo As String, ByVal ENTE As String, ByVal CONTRIB As Integer) As ListAnagrafica
    '    Try
    '        Dim objListaAnagrafica As New ListAnagrafica
    '        Dim objAnagrafica As New ClsDBManager

    '        objListaAnagrafica = objAnagrafica.getListaContribuentiFreezer(myStringConnection, ANNO, Tributo, ENTE, CONTRIB)

    '        Return objListaAnagrafica

    '    Catch ex As Exception
    '        Throw New Exception("Function::GetListAnagragrafica::COMPlusService" & "::" & " " & ex.Message)
    '    End Try
    'End Function
    ''' <summary>
    ''' Prelevo l’elenco delle detrazioni.
    ''' Pulisco la situazione per il contribuente in questione; creo la struttura, che servirà per i calcoli, come quella della tabella tabella di appoggio i dati principali delle unità immobiliari.
    ''' Cerco le dichiarazioni per anno; ciclo sulle dichiarazioni ottenute e popolo la struttura con i dati.
    ''' Ricalcolo i mesi. Al termine del ciclo salvo il dataset finale.
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="TributoCalcolo"></param>
    ''' <param name="AnnoDa"></param>
    ''' <param name="objDSAnagrafica"></param>
    ''' <param name="blnConfigurazioneDich"></param>
    ''' <param name="sTipoTASI"></param>
    ''' <param name="TASIAProprietario"></param>
    ''' <returns></returns>
    ''' <revisionHistory>
    ''' <revision date="07/02/2020">
    ''' tolto test su blnConfigurazioneDich perchè mai usato come dichiarazione.
    ''' Tolto filtro su stesso perchè crea duplicati se presenti più inquilini per lo stesso anno
    ''' </revision>
    ''' </revisionHistory>
    Public Function CreateFreezer(myStringConnection As String, IdEnte As String, TributoCalcolo As String, AnnoDa As String, ByVal objDSAnagrafica As DataSet, ByVal blnConfigurazioneDich As Boolean, sTipoTASI As String, TASIAProprietario As String) As Boolean
        'blnConfigurazioneDich=true --> configurazione del verticale ici che considera la dichiarazione MAI GESTITO QUINDI TOLTO
        'blnConfigurazioneDich=false--> configurazione del verticale ici che considera l'immobile
        Try
            Dim intCountAna, strCOD_CONTRIBUENTE As Long
            Dim dsDetraz As New DataSet
            Dim strAnnoFreezer As String = ""
            Dim myDsTemp As DataSet
            Dim myDsFinale As DataSet
            Dim dsDich As New DataSet
            Dim strIdTestataFreezer, strIdImmobileFreezer, strAnnoFreezerFine As String
            Dim strFoglio, strNumero, strSubalterno As String
            Dim DataInizio, DataFine As DateTime
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            Dim objDBOPENgovProvvedimentiSelect As New ClsDBManager
            Log.Debug("CreateFreezer.inizio")
            dsDetraz = objDBOPENgovProvvedimentiSelect.getDetrazioni(myStringConnection, IdEnte, TributoCalcolo)
            Log.Debug("CreateFreezer.prelevato detrazioni")
            For intCountAna = 0 To objDSAnagrafica.Tables(0).Rows.Count - 1
                strCOD_CONTRIBUENTE = 0
                strCOD_CONTRIBUENTE = CType(objDSAnagrafica.Tables(0).Rows(intCountAna).Item("COD_CONTRIBUENTE"), String)
                Log.Debug(IdEnte & " - Analizzo Contribuente " & strCOD_CONTRIBUENTE & " - " & vbTab & intCountAna + 1 & " di " & objDSAnagrafica.Tables(0).Rows.Count)
                'pulisco la situazione del FREEZER per il contribuente in questione
                objDBOPENgovProvvedimentiSelect.DeleteFreezer(myStringConnection, strCOD_CONTRIBUENTE, IdEnte, TributoCalcolo)
                'creo la struttura dei 2 dataset come quella della tabella TP_SITUAZIONE_VIRTUALE_DICHIARATO
                myDsTemp = setObjFREEZER()
                myDsFinale = setObjFREEZER()
                'cerco le dichiarazioni per anno
                Log.Debug("cerco le dichiarazioni per anno")
                '*** 20150430 - TASI Inquilino ***
                dsDich = objDBOPENgovProvvedimentiSelect.getTutteDichiarazioni(myStringConnection, IdEnte, strCOD_CONTRIBUENTE, AnnoDa, TributoCalcolo, TASIAProprietario, sTipoTASI)
                '*** ***
                If dsDich.Tables(0).Rows.Count > 0 Then 'ciclo sulle dichiarazioni
                    'se trovo delle dichiarazioni.....      
                    For Each myRow As DataRow In dsDich.Tables(0).Rows
                        strIdTestataFreezer = CType(myRow.Item("ID_TESTATA"), String)
                        strIdImmobileFreezer = CType(myRow.Item("ID_IMMOBILE"), String)

                        strFoglio = "" : strNumero = "" : strSubalterno = ""
                        If Not IsDBNull(myRow.Item("FOGLIO")) Then
                            strFoglio = CType(myRow.Item("FOGLIO"), String)
                        End If
                        If Not IsDBNull(myRow.Item("NUMERO")) Then
                            strNumero = CType(myRow.Item("NUMERO"), String)
                        End If
                        If Not IsDBNull(myRow.Item("SUBALTERNO")) Then
                            strSubalterno = CType(myRow.Item("SUBALTERNO"), String)
                        End If

                        If Not IsDBNull(myRow.Item("DATAINIZIO")) Then
                            DataInizio = DateTime.Parse(myRow.Item("DATAINIZIO"), culture).ToString("dd/MM/yyyy")
                        End If
                        If Not IsDBNull(myRow.Item("DATAFINE")) Then
                            DataFine = DateTime.Parse(myRow.Item("DATAFINE"), culture).ToString("dd/MM/yyyy")
                        Else
                            DataFine = DateTime.MaxValue.Date
                        End If

                        If DataFine.Date = DateTime.MaxValue.Date Then                       'vuol dire che l'immobile è aperto
                            'arrivo fino all'anno di calcolo
                            strAnnoFreezerFine = AnnoDa 'Now.Year
                        Else
                            strAnnoFreezerFine = DataFine.Year
                        End If
                        strAnnoFreezer = DataInizio.Year
                        '*** 20150430 - TASI Inquilino ***
                        strAnnoFreezer = AnnoDa
                        If TributoCalcolo = "" Then
                            FillRowFREEZER(Utility.Costanti.TRIBUTO_ICI, myRow, strAnnoFreezer, myDsTemp)
                            FillRowFREEZER(Utility.Costanti.TRIBUTO_TASI, myRow, strAnnoFreezer, myDsTemp)
                        Else
                            FillRowFREEZER(TributoCalcolo, myRow, strAnnoFreezer, myDsTemp)
                        End If
                        Update_MesiPossesso_Con_Periodo(myDsTemp)
                        'se sto calcolando ICI devo inserire solo se proprietario
                        Dim bEscludi As Boolean = False
                        If TributoCalcolo = Utility.Costanti.TRIBUTO_ICI And myRow.Item("TIPOTASI") = Utility.Costanti.TIPOTASI_INQUILINO Then
                            bEscludi = True
                        End If
                        If bEscludi = False Then
                            FillObjFREEZER(strAnnoFreezer, strAnnoFreezer, dsDetraz, blnConfigurazioneDich, myDsTemp, myDsFinale)
                        End If
                        '*** ***
                        myDsTemp.Clear()
                    Next
                    'salvataggio datasetfinale in TP_SITUAZIONE_VIRTUALE_DICHIARATO
                    objDBOPENgovProvvedimentiSelect.Set_TP_SITUAZIONE_VIRTUALE_DICHIARATO(myStringConnection, myDsFinale)
                    myDsTemp = Nothing
                    myDsFinale = Nothing
                    dsDich.Dispose()
                End If
            Next
            Log.Debug(IdEnte & " - ClsFreezer.CreateFreezer terminata correttamente")
            Return True
        Catch ex As Exception
            Log.Debug(IdEnte & " - Si è verificato un errore in ClsFreezer.CreateFreezer.errore::", ex)
            Throw New Exception(IdEnte & " - ClsFreezer.CreateFreezer.errore::" & ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Prelevo i dati degli immobili per i quali calcolare; ciclando per gli anni da calcolare se sono nell’anno di calcolo popolo i dati con calcolo altrimenti popolo i dati senza calcolo.
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="TributoCalcolo"></param>
    ''' <param name="AnnoDA"></param>
    ''' <param name="AnnoA"></param>
    ''' <param name="strCOD_CONTRIBUENTE"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <param name="TipoOperazione"></param>
    ''' <returns></returns>
    Public Function Calcola(StringConnectionGOV As String, ByVal myStringConnection As String, IdEnte As String, TributoCalcolo As String, AnnoDA As String, AnnoA As String, ByVal strCOD_CONTRIBUENTE As String, ByVal TipoCalcolo As Integer, TipoOperazione As String) As objSituazioneFinale()
        'Dim dsDichiaratoIci As DataSet = Nothing
        Dim dsSituazioneFinaleIci As DataSet
        Dim dsAnagrafica As DataSet = Nothing
        Dim AnnoDaLiquidare As Integer
        Dim dsDichiarazioniPerAnno As DataRow()
        Dim nAnni As Integer
        Dim dsICI() As objSituazioneFinale
        Dim i As Integer = 0
        Dim objDBICI As New ClsDBManager

        dsICI = Nothing
        Try
            If Not (AnnoA = "-1") And Not (AnnoA = AnnoDA) Then
                nAnni = Integer.Parse(AnnoA) - Integer.Parse(AnnoDA)
            Else
                nAnni = 1
            End If
            dsSituazioneFinaleIci = objDBICI.GetSituazioneVirtualeImmobili(myStringConnection, IdEnte, AnnoDA, AnnoA, strCOD_CONTRIBUENTE)
            For i = 0 To nAnni - 1
                AnnoDaLiquidare = CInt(AnnoDA)
                dsDichiarazioniPerAnno = dsSituazioneFinaleIci.Tables(0).Select("ANNO='" & AnnoDaLiquidare & "'")
                If dsDichiarazioniPerAnno.Length > 0 Then
                    dsICI = addRowsCalcoloICI(StringConnectionGOV, myStringConnection, IdEnte, dsDichiarazioniPerAnno, TipoOperazione, TipoCalcolo)
                Else
                    If TributoCalcolo = "" Then
                        dsICI = addRowsSenzaCalcoloICI(IdEnte, strCOD_CONTRIBUENTE, AnnoDaLiquidare, Utility.Costanti.TRIBUTO_ICI)
                        dsICI = addRowsSenzaCalcoloICI(IdEnte, strCOD_CONTRIBUENTE, AnnoDaLiquidare, Utility.Costanti.TRIBUTO_TASI)
                    Else
                        dsICI = addRowsSenzaCalcoloICI(IdEnte, strCOD_CONTRIBUENTE, AnnoDaLiquidare, TributoCalcolo)
                    End If
                End If
            Next
        Catch ex As Exception
            Log.Debug("CalcoloICI::si è verificato il seguente errore::" & ex.Message)
            dsICI = Nothing
        End Try
        Return dsICI
    End Function
    ''' <summary>
    ''' Funzione di definizione dataset con i dati della tabella di appoggio i dati principali delle unità immobiliari 
    ''' </summary>
    ''' <returns></returns>
    Private Function setObjFREEZER() As DataSet
        Dim objDS As New DataSet

        Dim newTable As DataTable
        newTable = New DataTable("TP_SITUAZIONE_VIRTUALE_DICHIARATO")

        Dim NewColumn As New DataColumn
        NewColumn.ColumnName = "ANNO"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "CODTRIBUTO"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "COD_CONTRIBUENTE"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "ID_TESTATA"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = "0"
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "ID_IMMOBILE"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = "0"
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "COD_TIPO_PROCEDIMENTO"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "COD_ENTE"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "NUMERO_MESI_ACCONTO"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "NUMERO_MESI_TOTALI"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "NUMERO_UTILIZZATORI"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "PERC_POSSESSO"
        NewColumn.DataType = System.Type.GetType("System.Double")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "VALORE"
        NewColumn.DataType = System.Type.GetType("System.Double")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "POSSESSO_FINE_ANNO"
        NewColumn.DataType = System.Type.GetType("System.Boolean")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "FOGLIO"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "NUMERO"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "SUBALTERNO"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "MESIPOSSESSO"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "FLAG_PRINCIPALE"
        'NewColumn.DataType = System.Type.GetType("System.Boolean")
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "RIDUZIONE"
        'NewColumn.DataType = System.Type.GetType("System.Boolean")
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "ESENTE_ESCLUSO"
        'NewColumn.DataType = System.Type.GetType("System.Boolean")
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)
        '*** 20140509 - TASI ***
        'NewColumn = New DataColumn
        'NewColumn.ColumnName = "TIPO_POSSESSO"
        'NewColumn.DataType = System.Type.GetType("System.Int64")
        'NewColumn.DefaultValue = 0
        'newTable.Columns.Add(NewColumn)
        NewColumn = New DataColumn
        NewColumn.ColumnName = "IDTIPOUTILIZZO"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)
        NewColumn = New DataColumn
        NewColumn.ColumnName = "IDTIPOPOSSESSO"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)
        '*** ***
        NewColumn = New DataColumn
        NewColumn.ColumnName = "RENDITA"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "IMPORTO_DETRAZIONE"
        NewColumn.DataType = System.Type.GetType("System.Double")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "DATA_INIZIO"
        NewColumn.DataType = System.Type.GetType("System.DateTime")
        'NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "DATA_FINE"
        NewColumn.DataType = System.Type.GetType("System.DateTime")
        'NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)

        'Ale 24052007
        NewColumn = New DataColumn
        NewColumn.ColumnName = "COD_IMMOBILE_PERTINENZA"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        NewColumn = New DataColumn
        NewColumn.ColumnName = "COD_IMMOBILE_DA_ACCERTAMENTO"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = ""
        newTable.Columns.Add(NewColumn)

        'dipe 27/04/2010
        NewColumn = New DataColumn
        NewColumn.ColumnName = "CONTITOLARE"
        NewColumn.DataType = System.Type.GetType("System.Boolean")
        NewColumn.DefaultValue = False
        newTable.Columns.Add(NewColumn)
        '*** 20150430 - TASI Inquilino ***
        NewColumn = New DataColumn
        NewColumn.ColumnName = "TIPOTASI"
        NewColumn.DataType = System.Type.GetType("System.String")
        NewColumn.DefaultValue = Utility.Costanti.TIPOTASI_PROPRIETARIO
        newTable.Columns.Add(NewColumn)
        NewColumn = New DataColumn
        NewColumn.ColumnName = "IDCONTRIBUENTECALCOLO"
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)
        '*** ***
        NewColumn = New DataColumn
        NewColumn.ColumnName = "IDCONTRIBUENTEDICH" 'mi serve per accertamento TASI
        NewColumn.DataType = System.Type.GetType("System.Int64")
        NewColumn.DefaultValue = 0
        newTable.Columns.Add(NewColumn)
        objDS.Tables.Add(newTable)

        Return objDS
    End Function
    ''' <summary>
    ''' Valorizzazione dati per il calcolo con prelievo delle aliquote e ricalcolo del valore aggiornato anche In base alla soglia rendita configurata per l'aliquota.
    ''' Prima del calcolo degli importi ridetermino la % di acconto e il numero dei mesi.
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="mySituazioneFinale"></param>
    ''' <param name="oListParametri"></param>
    ''' <param name="IsEsente"></param>
    ''' <param name="ListResiduoRendita"></param>
    ''' <returns></returns>
    ''' <revisionHistory>
    ''' <revision date="20190910">
    ''' In caso di flag storico la riduzione deve essere applicata anche per l'abitazione principale
    ''' </revision>
    ''' <revision date="20210514">
    ''' Il valore <strong><em>non</em></strong> deve essere ricalcolato in caso di Aree Fabbricabili.
    ''' </revision>
    ''' </revisionHistory>
    Private Function SetDatiPerCalcolo(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, ByVal mySituazioneFinale As objSituazioneFinale, ByVal oListParametri As Generale.PARAMETRI_ICI, ByRef IsEsente As Integer, ByRef ListResiduoRendita() As ObjResiduoRendita) As objCalcolo
        Dim FncUtil As New Generale
        Dim oMyDati As New objCalcolo
        Dim myAliquote As New ListALIQUOTA_DETRAZIONE
        Dim nValore As Double

        Try
            oMyDati.MesiT = mySituazioneFinale.Mesi
            oMyDati.Possesso = mySituazioneFinale.PercPossesso
            oMyDati.Utilizzatori = mySituazioneFinale.NUtilizzatori
            oMyDati.AbitazionePrincipale = mySituazioneFinale.FlagPrincipale
            oMyDati.Storico = mySituazioneFinale.FlagStorico
            If mySituazioneFinale.MesiRiduzione > 0 Then
                Log.Debug("SetDatiPerCalcolo ho " & mySituazioneFinale.MesiRiduzione.ToString() & " mesi di riduzione")
                oMyDati.Riduzione = True
            Else
                Log.Debug("SetDatiPerCalcolo non ho mesi di riduzione")
                oMyDati.Riduzione = False
            End If
            oMyDati.Acconto = getPercentualeAcconto(oListParametri)
            Log.Debug("ClsFreezer.SetDatiPerCalcolo.devo valorizzare i mesi per il calcolo:: ho " + mySituazioneFinale.Mesi.ToString + " in origine; i mesi di acconto sono " + mySituazioneFinale.AccMesi.ToString + " sto calcolando " + oListParametri.strACCONTO_TOTALE)
            oMyDati.Mesi = getMesi(oListParametri, mySituazioneFinale.AccMesi, mySituazioneFinale.Mesi)
            Log.Debug("ClsFreezer.SetDatiPerCalcolo.ottengo Mesi=" + oMyDati.Mesi.ToString)
            oMyDati.MesiT = mySituazioneFinale.Mesi
            'prelevo le aliquote da applicare in fase di calcolo
            '*** 20140509 - TASI ***
            myAliquote = New CALCOLO_ICI().getAliquotaDetrazione(StringConnectionICI, IdEnte, oListParametri, mySituazioneFinale.Tributo)

            '*** 20120530 - IMU ***
            'devo ricalcolare il valore aggiornato
            Dim FncValore As New ComPlusInterface.FncICI
            Dim nRendita As Double = 0
            nRendita = mySituazioneFinale.Rendita
            If myAliquote.sTipoSoglia = "<" And myAliquote.nSogliaRendita <> 0 Then
                If nRendita > myAliquote.nSogliaRendita Then
                    nRendita = myAliquote.nSogliaRendita
                End If
            Else
                nRendita -= myAliquote.nSogliaRendita
            End If
            If mySituazioneFinale.IdImmobile <> mySituazioneFinale.IdImmobilePertinenza And mySituazioneFinale.IdImmobilePertinenza > 0 Then
                'se sono pretinenza alla rendita tolgo il residuo soglia del principale
                If Not ListResiduoRendita Is Nothing Then
                    For Each myResiduoRendita As ObjResiduoRendita In ListResiduoRendita
                        If myResiduoRendita.IdPrincipale = mySituazioneFinale.IdImmobilePertinenza And myResiduoRendita.CodTributo = mySituazioneFinale.Tributo Then
                            If myResiduoRendita.sTipoSoglia = "<" Then
                                If nRendita > myResiduoRendita.impResiduo Then
                                    nRendita = myResiduoRendita.impResiduo
                                End If
                            Else
                                nRendita -= myResiduoRendita.impResiduo
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
            Dim bForzaRendita As Boolean = False
            If myAliquote.sTipoSoglia = "<" And myAliquote.nSogliaRendita <> 0 Then
                If nRendita > 0 Then
                    ListResiduoRendita = GetResiduoSogliaRenditaPrincipale(ListResiduoRendita, mySituazioneFinale.Tributo, mySituazioneFinale.IdImmobilePertinenza, mySituazioneFinale.IdImmobile, myAliquote.nSogliaRendita - nRendita, myAliquote.sTipoSoglia)
                End If
            Else
                If nRendita < 0 Then
                    ListResiduoRendita = GetResiduoSogliaRenditaPrincipale(ListResiduoRendita, mySituazioneFinale.Tributo, mySituazioneFinale.IdImmobilePertinenza, mySituazioneFinale.IdImmobile, nRendita * -1, myAliquote.sTipoSoglia)
                    'la rendita non può essere negativa viene quindi forzata a zero
                    nRendita = 0
                    bForzaRendita = True
                End If
            End If
            If (mySituazioneFinale.TipoRendita <> "AF" Or mySituazioneFinale.Valore = 0) Then
                mySituazioneFinale.Valore = FncValore.CalcoloValore(Generale.DBType, StringConnectionGOV, StringConnectionICI, mySituazioneFinale.IdEnte, mySituazioneFinale.Anno, mySituazioneFinale.TipoRendita, oListParametri.strCATEGORIA, mySituazioneFinale.Classe, mySituazioneFinale.Zona, nRendita, mySituazioneFinale.Valore, mySituazioneFinale.Consistenza, mySituazioneFinale.Dal, oListParametri.IsColtivatoreDiretto)
            End If

            nValore = mySituazioneFinale.Valore
            'Nuova Gestione valore per immobili di tipo B
            If Left(oListParametri.strCATEGORIA, 1) = "B" Then
                nValore = CalcoloValoreImmobiliB(oListParametri.intANNO_CALCOLO, mySituazioneFinale.Rendita, mySituazioneFinale.Dal, mySituazioneFinale.Al)
            End If
            If bForzaRendita = True Then
                oMyDati.Valore = 0
            Else
                oMyDati.Valore = nValore
            End If
            '*** ***
            oMyDati.TipoAliquota = myAliquote.TipoAliquota
            oMyDati.Aliquota = myAliquote.p_dblVALORE_ALIQUOTA
            oMyDati.Detrazione = myAliquote.p_dblVALORE_DETRAZIONE
            oMyDati.DetrazioneDichiarata = myAliquote.p_dblVALORE_DETRAZIONE
            IsEsente = myAliquote.p_ESENTE

            '*** 20120530 - IMU ***
            oMyDati.AliquotaStatale = myAliquote.AliquotaStatale
            oMyDati.AnnoCalcolo = oListParametri.intANNO_CALCOLO
            oMyDati.FigliACarico = mySituazioneFinale.NumeroFigli
            oMyDati.DetrazioneFigli = myAliquote.nDetrazioneFigli
            oMyDati.PercentCaricoFigli = mySituazioneFinale.PercentCaricoFigli
            '*** ***
            '*** 20130422 - aggiornamento IMU ***
            oMyDati.nIdAliquota = myAliquote.nIdAliquota
            '*** ***
            '*** 20150430 - TASI Inquilino ***
            Try
                oMyDati.TipoTasi = mySituazioneFinale.TipoTasi
            Catch ex As Exception

            End Try
            oMyDati.nPercInquilino = myAliquote.nPercInquilino
            '*** ***
            Return oMyDati
        Catch ex As Exception
            Log.Error("COMPlusOPENgovProvvedimenti::COMPLUSBaseCalcoloICI::SetDatiPerCalcolo::Errore::" & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function GetResiduoSogliaRenditaPrincipale(ByVal ListResiduoRendita() As ObjResiduoRendita, ByVal CodTributo As String, ByVal nIdPertinenza As Integer, ByVal nIdPrincipale As Integer, ByVal impRendita As Double, ByVal sTipoSoglia As String) As ObjResiduoRendita()
        Dim ListMyResiduoRendita() As ObjResiduoRendita = ListResiduoRendita
        Dim nList As Integer = 0
        Dim myResiduoRendita As New ObjResiduoRendita
        Dim bNewPrincipale As Boolean = True

        Try
            'carico l'oggetto con i residui rendita da applicare
            If Not ListMyResiduoRendita Is Nothing Then
                nList = ListResiduoRendita.GetUpperBound(0) + 1
                'controllo che non ci sia già
                For Each myResiduoRendita In ListMyResiduoRendita
                    If myResiduoRendita.IdPrincipale = nIdPertinenza And myResiduoRendita.CodTributo = CodTributo Then
                        bNewPrincipale = False
                        Exit For
                    End If
                Next
            End If
            If bNewPrincipale = True Then
                myResiduoRendita.CodTributo = CodTributo
                myResiduoRendita.IdPrincipale = nIdPrincipale
                myResiduoRendita.impResiduo = impRendita
                myResiduoRendita.sTipoSoglia = sTipoSoglia
                ReDim Preserve ListMyResiduoRendita(nList)
                ListMyResiduoRendita(nList) = myResiduoRendita
            End If
        Catch ex As Exception
            Log.Debug("GetResiduoSogliaRenditaPrincipale::si è verificato il seguente errore::", ex)
            ListResiduoRendita = Nothing
        End Try
        Return ListMyResiduoRendita
    End Function

    Private Function SetICIImporti(ByVal nTipo As Integer, ByVal oDati As objCalcolo, ByVal oICICalcolata As Generale.VALORI_ICI_CALCOLATA) As Generale.VALORI_ICI_CALCOLATA
        Dim oMyICICalcolata As New Generale.VALORI_ICI_CALCOLATA
        'nTipo=1(Acconto), 2(Saldo), 0(Esente)
        Try
            Select Case nTipo
                Case 0
                    oMyICICalcolata.dblICI_ACCONTO_DOVUTA = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_ACCONTO_SENZA_DETRAZIONE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA = FormatNumber(0, 2)

                    oMyICICalcolata.dblICI_SALDO_DOVUTA = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_APPLICATA = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_SALDO_SENZA_DETRAZIONE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_RESIDUA = FormatNumber(0, 2)

                    oMyICICalcolata.dblICI_TOTALE_DOVUTA = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_TOTALE_SENZA_DETRAZIONE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_RESIDUA = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_APPLICATA = FormatNumber(0, 2)

                    oMyICICalcolata.dblICI_VALORE_ALIQUOTA = 0
                    '*** 20120530 - IMU ***
                    oMyICICalcolata.dblICI_VALORE_ALIQUOTA_STATALE = 0
                    oMyICICalcolata.dblICI_ACCONTO_DOVUTA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_SALDO_DOVUTA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_RESIDUA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_APPLICATA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_TOTALE_DOVUTA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_RESIDUA_STATALE = FormatNumber(0, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_APPLICATA_STATALE = FormatNumber(0, 2)
                    '*** ***
                    '*** 20130422 - aggiornamento IMU ***
                    oMyICICalcolata.nIdAliquota = 0
                    '*** ***
                Case 1
                    oMyICICalcolata.dblICI_ACCONTO_DOVUTA = FormatNumber(oDati.Ici_Dovuta, 2)
                    oMyICICalcolata.dblICI_ACCONTO_SENZA_DETRAZIONE = FormatNumber(oDati.Ici_Teorica, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA = FormatNumber(oDati.Detrazione_Residua, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA = FormatNumber(oDati.Detrazione_Applicabile, 2)
                    '*** 20120530 - IMU ***
                    oMyICICalcolata.dblICI_ACCONTO_DOVUTA_STATALE = FormatNumber(oDati.Ici_Dovuta_Statale, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE = FormatNumber(oDati.Detrazione_Residua_Statale, 2)
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE = FormatNumber(oDati.Detrazione_Applicabile_Statale, 2)
                    '*** ***
                    '*** 20130422 - aggiornamento IMU ***
                    oMyICICalcolata.nIdAliquota = oDati.nIdAliquota
                    '*** ***
                Case 2
                    oMyICICalcolata.dblICI_TOTALE_DOVUTA = FormatNumber(oDati.Ici_Dovuta, 2)
                    Log.Debug("SetICIImporti::ho TOTALE DOVUTO=" & FormatNumber(oDati.Ici_Dovuta, 2))
                    oMyICICalcolata.dblICI_TOTALE_SENZA_DETRAZIONE = FormatNumber(oDati.Ici_Teorica, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_RESIDUA = FormatNumber(oDati.Detrazione_Residua, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_APPLICATA = FormatNumber(oDati.Detrazione_Applicabile, 2)

                    '*************************************************************************
                    'DA VERIFICARE
                    '*************************************************************************
                    oMyICICalcolata.dblICI_SALDO_DOVUTA = FormatNumber(oMyICICalcolata.dblICI_TOTALE_DOVUTA - oICICalcolata.dblICI_ACCONTO_DOVUTA, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_APPLICATA = FormatNumber(oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_APPLICATA - oICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA, 2)
                    oMyICICalcolata.dblICI_SALDO_SENZA_DETRAZIONE = FormatNumber(oMyICICalcolata.dblICI_TOTALE_SENZA_DETRAZIONE - oICICalcolata.dblICI_ACCONTO_SENZA_DETRAZIONE, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_RESIDUA = FormatNumber(oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_RESIDUA - oICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA, 2)

                    oMyICICalcolata.dblICI_VALORE_ALIQUOTA = oDati.Aliquota
                    '*** 20120530 - IMU ***
                    oMyICICalcolata.dblICI_VALORE_ALIQUOTA_STATALE = oDati.AliquotaStatale
                    oMyICICalcolata.dblICI_TOTALE_DOVUTA_STATALE = FormatNumber(oDati.Ici_Dovuta_Statale, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_RESIDUA_STATALE = FormatNumber(oDati.Detrazione_Residua_Statale, 2)
                    oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_APPLICATA_STATALE = FormatNumber(oDati.Detrazione_Applicabile_Statale, 2)
                    oMyICICalcolata.dblICI_SALDO_DOVUTA_STATALE = FormatNumber(oMyICICalcolata.dblICI_TOTALE_DOVUTA_STATALE - oICICalcolata.dblICI_ACCONTO_DOVUTA_STATALE, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_RESIDUA_STATALE = FormatNumber(oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_RESIDUA_STATALE - oICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE, 2)
                    oMyICICalcolata.dblICI_SALDO_DETRAZIONE_APPLICATA_STATALE = FormatNumber(oMyICICalcolata.dblICI_TOTALE_DETRAZIONE_APPLICATA_STATALE - oICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE, 2)
                    '*** ***
                    'devo risettare anche gli importi di acconto altrimenti li perde
                    oMyICICalcolata.dblICI_ACCONTO_DOVUTA = oICICalcolata.dblICI_ACCONTO_DOVUTA
                    oMyICICalcolata.dblICI_ACCONTO_SENZA_DETRAZIONE = oICICalcolata.dblICI_ACCONTO_SENZA_DETRAZIONE
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA = oICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA = oICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA
                    '*** 20120530 - IMU ***
                    oMyICICalcolata.dblICI_ACCONTO_DOVUTA_STATALE = oICICalcolata.dblICI_ACCONTO_DOVUTA_STATALE
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE = oICICalcolata.dblICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE
                    oMyICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE = oICICalcolata.dblICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE
                    '*** ***
                    '*** 20130422 - aggiornamento IMU ***
                    oMyICICalcolata.nIdAliquota = oICICalcolata.nIdAliquota
                    '*** ***
            End Select
            Return oMyICICalcolata
        Catch ex As Exception
            Log.Error("COMPlusOPENgovProvvedimenti::COMPLUSBaseCalcoloICI::SetICIImporti::Errore durante la fase di calcolo dell'ICI " & ex.Message)
            Throw New Exception("COMPlusOPENgovProvvedimenti::COMPLUSBaseCalcoloICI::SetICIImporti::Errore durante la fase di calcolo dell'ICI")
        End Try
    End Function

    Private Sub setDS_TabellaSituazioneFinaleICI(ByRef mySituazioneFinale As objSituazioneFinale, ByVal oMyImpCalcolati As Generale.VALORI_ICI_CALCOLATA, ByVal bIsColtivatoreDiretto As Boolean, ByVal nNumeroFigli As Integer, ByVal nPercentCaricoFigli As Double) 'Private Sub setDS_TabellaSituazioneFinaleICI(ByRef rowTabellaSituazioneFinaleICI As DataRow, ByVal oMyImpCalcolati As Generale.VALORI_ICI_CALCOLATA, ByVal bIsColtivatoreDiretto As Boolean, ByVal nNumeroFigli As Integer, ByVal nPercentCaricoFigli As Double)
        mySituazioneFinale.AccSenzaDetrazione = oMyImpCalcolati.dblICI_ACCONTO_SENZA_DETRAZIONE
        mySituazioneFinale.AccDetrazioneApplicata = oMyImpCalcolati.dblICI_ACCONTO_DETRAZIONE_APPLICATA
        mySituazioneFinale.AccDovuto = oMyImpCalcolati.dblICI_ACCONTO_DOVUTA
        mySituazioneFinale.AccDetrazioneResidua = oMyImpCalcolati.dblICI_ACCONTO_DETRAZIONE_RESIDUA

        mySituazioneFinale.TotSenzaDetrazione = oMyImpCalcolati.dblICI_TOTALE_SENZA_DETRAZIONE
        mySituazioneFinale.TotDetrazioneApplicata = oMyImpCalcolati.dblICI_TOTALE_DETRAZIONE_APPLICATA
        mySituazioneFinale.TotDovuto = oMyImpCalcolati.dblICI_TOTALE_DOVUTA
        Log.Debug("TabellaSituazioneFinaleICI::ho TOTALE DOVUTA=" & oMyImpCalcolati.dblICI_TOTALE_DOVUTA.ToString)
        mySituazioneFinale.TotDetrazioneResidua = oMyImpCalcolati.dblICI_TOTALE_DETRAZIONE_RESIDUA

        mySituazioneFinale.SalSenzaDetrazione = oMyImpCalcolati.dblICI_SALDO_SENZA_DETRAZIONE
        mySituazioneFinale.SalDetrazioneApplicata = oMyImpCalcolati.dblICI_SALDO_DETRAZIONE_APPLICATA
        mySituazioneFinale.SalDovuto = oMyImpCalcolati.dblICI_SALDO_DOVUTA
        mySituazioneFinale.SalDetrazioneResidua = oMyImpCalcolati.dblICI_SALDO_DETRAZIONE_RESIDUA

        mySituazioneFinale.Aliquota = oMyImpCalcolati.dblICI_VALORE_ALIQUOTA
        '*** 20120530 - IMU ***
        mySituazioneFinale.AliquotaStatale = oMyImpCalcolati.dblICI_VALORE_ALIQUOTA_STATALE
        mySituazioneFinale.IsColtivatoreDiretto = bIsColtivatoreDiretto
        mySituazioneFinale.NumeroFigli = nNumeroFigli
        mySituazioneFinale.PercentCaricoFigli = nPercentCaricoFigli

        mySituazioneFinale.AccDovutoStatale = oMyImpCalcolati.dblICI_ACCONTO_DOVUTA_STATALE
        mySituazioneFinale.AccDetrazioneApplicataStatale = oMyImpCalcolati.dblICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE
        mySituazioneFinale.AccDetrazioneResiduaStatale = oMyImpCalcolati.dblICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE

        mySituazioneFinale.SalDovutoStatale = oMyImpCalcolati.dblICI_SALDO_DOVUTA_STATALE
        mySituazioneFinale.SalDetrazioneApplicataStatale = oMyImpCalcolati.dblICI_SALDO_DETRAZIONE_APPLICATA_STATALE
        mySituazioneFinale.SalDetrazioneResiduaStatale = oMyImpCalcolati.dblICI_SALDO_DETRAZIONE_RESIDUA_STATALE

        mySituazioneFinale.TotDovutoStatale = oMyImpCalcolati.dblICI_TOTALE_DOVUTA_STATALE
        mySituazioneFinale.TotDetrazioneApplicataStatale = oMyImpCalcolati.dblICI_TOTALE_DETRAZIONE_APPLICATA_STATALE
        mySituazioneFinale.TotDetrazioneResidua = oMyImpCalcolati.dblICI_TOTALE_DETRAZIONE_RESIDUA_STATALE
        '*** ***
        '*** 20130422 - aggiornamento IMU ***
        mySituazioneFinale.IdAliquota = oMyImpCalcolati.nIdAliquota
    End Sub

    '**************************************************************************************
    'GESTIONE DELLE PERTINENZE
    '**************************************************************************************
    'Private Function Gestione_Pertinenze(ByVal cmdMyCommand As SqlCommand, ByVal dsTabellaSituazioneFinaleICI As DataSet, ByVal enmTIPO_ABITAZIONE As Generale.ABITAZIONE_PRINCIPALE_PERTINENZA) As DataSet

    '    'Dim objDSTabellaSituazioneFinaleICIClonePertinenza As DataSet
    '    'Dim objDSTabellaSituazioneFinaleICICloneReale As DataSet
    '    Dim intCount, intCountReale As Integer

    '    Dim intID_SITUAZIONE_FINALE As Integer

    '    Dim objUtility As New Generale
    '    '*************************************************************
    '    'Dim dblICI_ACCONTO_DOVUTA_PERTINENZA As Double
    '    'Dim dblICI_TOTALE_DOVUTA_PERTINENZA As Double
    '    'Dim dblICI_SALDO_DOVUTA_PERTINENZA As Double
    '    '*************************************************************
    '    'Dim dblICI_ACCONTO_SENZA_DETRAZIONE As Double
    '    'Dim dblICI_ACCONTO_DETRAZIONE_APPLICATA As Double
    '    Dim dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE As Double
    '    'Dim dblICI_TOTALE_SENZA_DETRAZIONE As Double
    '    'Dim dblICI_TOTALE_DETRAZIONE_APPLICATA As Double
    '    Dim dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE As Double
    '    'Dim dblICI_SALDO_SENZA_DETRAZIONE As Double
    '    'Dim dblICI_SALDO_DETRAZIONE_APPLICATA As Double
    '    Dim dblICI_SALDO_DETRAZIONE_RESIDUA_REALE As Double
    '    '*** 20120530 - IMU ***
    '    Dim dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE_STATALE As Double
    '    Dim dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE_STATALE As Double
    '    Dim dblICI_SALDO_DETRAZIONE_RESIDUA_REALE_STATALE As Double
    '    '*** ***
    '    Dim intID_Immobile_Pertinenza As Integer
    '    Dim objDBOPENgovProvvedimentiSelect As New ClsDBManager
    '    Dim bEscludi As Boolean

    '    'objDSTabellaSituazioneFinaleICIClonePertinenza = dsTabellaSituazioneFinaleICI.Copy
    '    'objDSTabellaSituazioneFinaleICICloneReale = dsTabellaSituazioneFinaleICI.Copy

    '    Dim objRowsDichiarazioniDaBonificarePertinenza() As DataRow
    '    Dim objRowsDichiarazioniDaBonificareReale() As DataRow
    '    'Log.Debug("Inizio Gestione_Pertinenze")
    '    Try
    '        objRowsDichiarazioniDaBonificarePertinenza = dsTabellaSituazioneFinaleICI.Tables("TP_SITUAZIONE_FINALE_ICI").Select("FLAG_PRINCIPALE=" & enmTIPO_ABITAZIONE.ABITAZIONE_PERTINENZA & " AND CATEGORIA LIKE 'C%'")

    '        For intCount = 0 To objRowsDichiarazioniDaBonificarePertinenza.Length - 1
    '            bEscludi = False

    '            intID_Immobile_Pertinenza = objUtility.CToInt(objRowsDichiarazioniDaBonificarePertinenza(intCount)("COD_IMMOBILE_PERTINENZA"))

    '            'dblICI_ACCONTO_DOVUTA_PERTINENZA = objUtility.cToDbl(objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO"))
    '            'dblICI_TOTALE_DOVUTA_PERTINENZA = objUtility.cToDbl(objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA"))
    '            'dblICI_SALDO_DOVUTA_PERTINENZA = objUtility.cToDbl(objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO"))
    '            '***************************************************************************
    '            'Un immobile può avere più pertinenze
    '            'ma una pertinenza un solo immobile
    '            'Il risultato della query nel dataset sarà un unico immobile
    '            '***************************************************************************
    '            objRowsDichiarazioniDaBonificareReale = dsTabellaSituazioneFinaleICI.Tables("TP_SITUAZIONE_FINALE_ICI").Select("COD_IMMOBILE=" & intID_Immobile_Pertinenza)
    '            If objRowsDichiarazioniDaBonificareReale.Length = 1 Then
    '                Dim dsCatDaEscludereAP, dsCatDaEscludereAUG1, dsCatDaEscludereAUG2, dsCatDaEscludereAUG3 As DataSet

    '                Dim strAnno, strCategoria, sTipoUtilizzo As String ', strTipoPossesso
    '                Dim strAbitPr As String
    '                strAnno = objRowsDichiarazioniDaBonificareReale(0)("ANNO")
    '                strCategoria = objRowsDichiarazioniDaBonificareReale(0)("CATEGORIA")
    '                'dipe mofificato da ABITAZIONE_PRINCIPALE_ATTUALE  a FLAG_PRINCIPALE
    '                strAbitPr = objRowsDichiarazioniDaBonificareReale(0)("FLAG_PRINCIPALE")
    '                '*** 20140509 - TASI ***
    '                'strTipoPossesso = objRowsDichiarazioniDaBonificareReale(0)("TIPO_POSSESSO")
    '                sTipoUtilizzo = objRowsDichiarazioniDaBonificareReale(0)("IDTIPOUTILIZZO")
    '                '*** ***
    '                dsCatDaEscludereAP = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(cmdMyCommand, strAnno, m_objHashTable, Generale.TipoAliquote_AAP, objRowsDichiarazioniDaBonificareReale(0)("CODTRIBUTO"))
    '                dsCatDaEscludereAUG1 = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(cmdMyCommand, strAnno, m_objHashTable, Generale.TipoAliquote_AUG1, objRowsDichiarazioniDaBonificareReale(0)("CODTRIBUTO"))
    '                dsCatDaEscludereAUG2 = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(cmdMyCommand, strAnno, m_objHashTable, Generale.TipoAliquote_AUG2, objRowsDichiarazioniDaBonificareReale(0)("CODTRIBUTO"))
    '                dsCatDaEscludereAUG3 = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(cmdMyCommand, strAnno, m_objHashTable, Generale.TipoAliquote_AUG3, objRowsDichiarazioniDaBonificareReale(0)("CODTRIBUTO"))
    '                'se l'immobile è 
    '                '   abitazione principale e appartiene a gategoria da escludere
    '                '   o uso gratuito 1 e appartiene a categoria da escludere
    '                '   o uso gratuito 2 e appartiene a categoria da escludere
    '                '   o uso gratuito 3 e appartiene a categoria da escludere
    '                '----
    '                'allora ici=0 si per immobile principale che per pertinenza
    '                '*** 20140509 - TASI ***
    '                'If (strAbitPr = "1" And dsCatDaEscludereAP.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0) _
    '                ' Or (strTipoPossesso = "2" And dsCatDaEscludereAUG1.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0) _
    '                ' Or (strTipoPossesso = "3" And dsCatDaEscludereAUG2.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0) _
    '                ' Or (strTipoPossesso = "4" And dsCatDaEscludereAUG3.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0) Then
    '                If (strAbitPr = "1" And dsCatDaEscludereAP.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0) _
    '                 Or (sTipoUtilizzo = "5" And (dsCatDaEscludereAUG1.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0 Or dsCatDaEscludereAUG2.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0 Or dsCatDaEscludereAUG3.Tables(0).Select("COD_CAT='" & strCategoria & "'").Length > 0)) Then
    '                    '*** ***
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_TOTALE_DETRAZIONE_RESIDUA") = 0
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_DETRAZIONE_RESIDUA") = 0
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_ACCONTO_DETRAZIONE_RESIDUA") = 0

    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_TOTALE_DOVUTA") = 0
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_SALDO") = 0
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_ACCONTO") = 0

    '                    objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA") = 0
    '                    objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO") = 0
    '                    objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO") = 0
    '                    '*** 20120530 - IMU ***
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_TOTALE_DOVUTA_STATALE") = 0
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_SALDO_STATALE") = 0
    '                    objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_ACCONTO_STATALE") = 0

    '                    objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA_STATALE") = 0
    '                    objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO_STATALE") = 0
    '                    objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO_STATALE") = 0
    '                    '*** ***
    '                    bEscludi = True
    '                End If
    '            End If

    '            If Not bEscludi Then
    '                For intCountReale = 0 To objRowsDichiarazioniDaBonificareReale.Length - 1
    '                    '*************************************************************************
    '                    'ID della tabella sulla quale si dovranno portare gli aggiornamenti
    '                    '*************************************************************************
    '                    intID_SITUAZIONE_FINALE = objRowsDichiarazioniDaBonificareReale(intCountReale)("ID_SITUAZIONE_FINALE")

    '                    dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE = objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_ACCONTO_DETRAZIONE_RESIDUA"), False)
    '                    dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE = objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_TOTALE_DETRAZIONE_RESIDUA"), False)
    '                    dblICI_SALDO_DETRAZIONE_RESIDUA_REALE = objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_DOVUTA_DETRAZIONE_RESIDUA"), False)
    '                    If objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_TOTALE_DETRAZIONE_RESIDUA"), False) < 0 Then
    '                        dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE *= -1
    '                        dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE *= -1
    '                        dblICI_SALDO_DETRAZIONE_RESIDUA_REALE *= -1
    '                    End If
    '                    '*** 20120530 - IMU ***
    '                    dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE_STATALE = objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE"), False)
    '                    dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE_STATALE = objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_TOTALE_DETRAZIONE_RESIDUA_STATALE"), False)
    '                    dblICI_SALDO_DETRAZIONE_RESIDUA_REALE_STATALE = objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_DOVUTA_DETRAZIONE_RESIDUA_STATALE"), False)
    '                    If objUtility.CToDouble(objRowsDichiarazioniDaBonificareReale(intCountReale)("ICI_TOTALE_DETRAZIONE_RESIDUA_STATALE"), False) < 0 Then
    '                        dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE_STATALE *= -1
    '                        dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE_STATALE *= -1
    '                        dblICI_SALDO_DETRAZIONE_RESIDUA_REALE_STATALE *= -1
    '                    End If
    '                    '*** ***
    '                Next

    '                'Dim DBRow As DataRow

    '                'AGGIORNO L'ICI DOVUTA PER LA PERTINENZA IN QUESTIONE
    '                'E AGGIORNO LA DETRAZIONE RESIDUA DELL'ABITAZIONE PRINCIPALE TOGLIENDO IL DOVUTO DELLA PERTINENZA IN QUESTIONE
    '                'If dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE > 0 Then 
    '                'alep 09042008 
    '                If dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE > 0 And objRowsDichiarazioniDaBonificareReale.Length > 0 Then
    '                    If objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA") - dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE < 0 Then
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_TOTALE_DETRAZIONE_RESIDUA") = dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE - objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA")
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA") = 0
    '                    Else
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA") = objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA") - dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE
    '                        Log.Debug("GestionePertinenze::ho TOTALE DOVUTA=" & objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA").ToString)
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_TOTALE_DETRAZIONE_RESIDUA") = 0
    '                    End If

    '                    If objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO") - dblICI_SALDO_DETRAZIONE_RESIDUA_REALE < 0 Then
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_DETRAZIONE_RESIDUA") = dblICI_SALDO_DETRAZIONE_RESIDUA_REALE - objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO")
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO") = 0
    '                    Else
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO") = objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO") - dblICI_SALDO_DETRAZIONE_RESIDUA_REALE
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_DETRAZIONE_RESIDUA") = 0
    '                    End If

    '                    If objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO") - dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE < 0 Then
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_ACCONTO_DETRAZIONE_RESIDUA") = dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE - objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO")
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO") = 0
    '                    Else
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO") = objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO") - dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_ACCONTO_DETRAZIONE_RESIDUA") = 0
    '                    End If
    '                End If
    '                '*** 20120530 - IMU ***
    '                If dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE_STATALE > 0 And objRowsDichiarazioniDaBonificareReale.Length > 0 Then
    '                    If objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA_STATALE") - dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE_STATALE < 0 Then
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_TOTALE_DETRAZIONE_RESIDUA_STATALE") = dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE_STATALE - objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA_STATALE")
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA_STATALE") = 0
    '                    Else
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA_STATALE") = objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_TOTALE_DOVUTA_STATALE") - dblICI_TOTALE_DETRAZIONE_RESIDUA_REALE_STATALE
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_TOTALE_DETRAZIONE_RESIDUA_STATALE") = 0
    '                    End If

    '                    If objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO_STATALE") - dblICI_SALDO_DETRAZIONE_RESIDUA_REALE_STATALE < 0 Then
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_DETRAZIONE_RESIDUA_STATALE") = dblICI_SALDO_DETRAZIONE_RESIDUA_REALE_STATALE - objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO_STATALE")
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO_STATALE") = 0
    '                    Else
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO_STATALE") = objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_SALDO_STATALE") - dblICI_SALDO_DETRAZIONE_RESIDUA_REALE_STATALE
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_DOVUTA_DETRAZIONE_RESIDUA_STATALE") = 0
    '                    End If

    '                    If objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO_STATALE") - dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE_STATALE < 0 Then
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE") = dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE_STATALE - objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO_STATALE")
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO_STATALE") = 0
    '                    Else
    '                        objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO_STATALE") = objRowsDichiarazioniDaBonificarePertinenza(intCount)("ICI_DOVUTA_ACCONTO_STATALE") - dblICI_ACCONTO_DETRAZIONE_RESIDUA_REALE_STATALE
    '                        objRowsDichiarazioniDaBonificareReale(0)("ICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE") = 0
    '                    End If
    '                End If
    '                '*** ***
    '            End If

    '            objRowsDichiarazioniDaBonificarePertinenza(intCount).AcceptChanges()
    '            'Dipe 21/04/2009 aggiunto controllo perchè capita che non venga trovata la pertinenza
    '            If objRowsDichiarazioniDaBonificareReale.Length > 0 Then
    '                objRowsDichiarazioniDaBonificareReale(0).AcceptChanges()
    '            End If
    '        Next

    '    Catch ex As Exception
    '        Log.Error("Gestione_Pertinenze:" & ex.Message & " " & ex.StackTrace)
    '    End Try
    '    Return dsTabellaSituazioneFinaleICI
    'End Function

    '*** 20121203 - IMU calcolo dovuto al netto del versato ***
    Private Function CalcoloNettoVersato(myStringConnection As String, ByVal dsCalcolato As objSituazioneFinale()) As objSituazioneFinale()
        Dim myDataView As New DataView
        Dim sSQL As String

        '*** ATTENZIONE L'ARROTONDAMENTO PUò ESSERE SOLO SU TOTALE O SU SINGOLE VOCI ***
        Try
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_CalcoloNettoVersato", "ente", "annoRiferimento", "idAnagrafico")
                myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("ente", dsCalcolato(0).IdEnte) _
                                    , ctx.GetParam("annoRiferimento", dsCalcolato(0).Anno) _
                                    , ctx.GetParam("idAnagrafico", dsCalcolato(0).IdContribuente)
                                )
                For Each myRow As DataRowView In myDataView
                    For Each mySituazioneFinale As objSituazioneFinale In dsCalcolato
                        If mySituazioneFinale.FlagPrincipale = 1 Or mySituazioneFinale.IdImmobilePertinenza > 0 Then
                            mySituazioneFinale.AccDovuto = mySituazioneFinale.AccDovuto - myRow("ABIPRIN")
                            mySituazioneFinale.SalDovuto = mySituazioneFinale.SalDovuto - myRow("ABIPRIN")
                            mySituazioneFinale.TotDovuto = mySituazioneFinale.TotDovuto - myRow("ABIPRIN")
                            Log.Debug("CalcoloNettoVersato::AP::ho TOTALE DOVUTA=" & mySituazioneFinale.TotDovuto.ToString)

                            mySituazioneFinale.AccDetrazioneApplicata = mySituazioneFinale.AccDetrazioneApplicata - myRow("DETRAZIONE")
                            mySituazioneFinale.SalDetrazioneApplicata = mySituazioneFinale.SalDetrazioneApplicata - myRow("DETRAZIONE")
                            mySituazioneFinale.TotDetrazioneApplicata = mySituazioneFinale.TotDetrazioneApplicata - myRow("DETRAZIONE")
                        End If
                        If mySituazioneFinale.TipoRendita <> "AF" And mySituazioneFinale.TipoRendita <> "TA" And mySituazioneFinale.FlagPrincipale <> 1 And mySituazioneFinale.IdImmobilePertinenza = -1 Then
                            mySituazioneFinale.AccDovuto = mySituazioneFinale.AccDovuto - (myRow("ALTRIFAB") + myRow("ALTRIFABSTATALE"))
                            mySituazioneFinale.SalDovuto = mySituazioneFinale.SalDovuto - (myRow("ALTRIFAB") + myRow("ALTRIFABSTATALE"))
                            mySituazioneFinale.TotDovuto = mySituazioneFinale.TotDovuto - (myRow("ALTRIFAB") + myRow("ALTRIFABSTATALE"))

                            mySituazioneFinale.AccDovutoStatale = mySituazioneFinale.AccDovuto - myRow("ALTRIFABSTATALE")
                            mySituazioneFinale.SalDovutoStatale = mySituazioneFinale.SalDovuto - myRow("ALTRIFABSTATALE")
                            mySituazioneFinale.TotDovutoStatale = mySituazioneFinale.TotDovuto - myRow("ALTRIFABSTATALE")
                        End If
                        If mySituazioneFinale.TipoRendita = "AF" Then
                            mySituazioneFinale.AccDovuto = mySituazioneFinale.AccDovuto - (myRow("AREEFAB") + myRow("AREEFABSTATALE"))
                            mySituazioneFinale.SalDovuto = mySituazioneFinale.SalDovuto - (myRow("AREEFAB") + myRow("AREEFABSTATALE"))
                            mySituazioneFinale.TotDovuto = mySituazioneFinale.TotDovuto - (myRow("AREEFAB") + myRow("AREEFABSTATALE"))
                            Log.Debug("CalcoloNettoVersatoAF : ho TOTALE DOVUTA=" & mySituazioneFinale.TotDovuto.ToString)

                            mySituazioneFinale.AccDovutoStatale = mySituazioneFinale.AccDovuto - myRow("AREEFABSTATALE")
                            mySituazioneFinale.SalDovutoStatale = mySituazioneFinale.SalDovuto - myRow("AREEFABSTATALE")
                            mySituazioneFinale.TotDovutoStatale = mySituazioneFinale.TotDovuto - myRow("AREEFABSTATALE")
                        End If
                        If mySituazioneFinale.TipoRendita = "TA" Then
                            mySituazioneFinale.AccDovuto = mySituazioneFinale.AccDovuto - (myRow("TERAGR") + myRow("TERAGRSTATALE"))
                            mySituazioneFinale.SalDovuto = mySituazioneFinale.SalDovuto - (myRow("TERAGR") + myRow("TERAGRSTATALE"))
                            mySituazioneFinale.TotDovuto = mySituazioneFinale.TotDovuto - (myRow("TERAGR") + myRow("TERAGRSTATALE"))
                            Log.Debug("CalcoloNettoVersato::TA::ho TOTALE DOVUTA=" & mySituazioneFinale.TotDovuto.ToString)

                            mySituazioneFinale.AccDovutoStatale = mySituazioneFinale.AccDovuto - myRow("TERAGRSTATALE")
                            mySituazioneFinale.SalDovutoStatale = mySituazioneFinale.SalDovuto - myRow("TERAGRSTATALE")
                            mySituazioneFinale.TotDovutoStatale = mySituazioneFinale.TotDovuto - myRow("TERAGRSTATALE")
                        End If
                    Next
                Next
                ctx.Dispose()
            End Using
        Catch ex As Exception
            Log.Debug("ClsFreezer.CalcoloNettoVersato.Errore.", ex)
            dsCalcolato = Nothing
        Finally
            myDataView.Dispose()
        End Try
        Return dsCalcolato
    End Function
    '*** ***

    Private Sub DeleteRowFREEZER(ByVal strFoglio As String, ByVal strNumero As String, ByVal strSubalterno As String, ByRef myDsTemp As DataSet)
        Try
            Dim iCount As Integer
            Dim pippo As Integer

            pippo = 0
            For iCount = 0 To myDsTemp.Tables(0).Rows.Count - 1
                If myDsTemp.Tables(0).Rows(iCount - pippo).Item("FOGLIO") = strFoglio And myDsTemp.Tables(0).Rows(iCount - pippo).Item("NUMERO") = strNumero And myDsTemp.Tables(0).Rows(iCount - pippo).Item("SUBALTERNO") = strSubalterno Then
                    myDsTemp.Tables(0).Rows(iCount - pippo).Delete()
                    myDsTemp.Tables(0).Rows(iCount - pippo).AcceptChanges()
                    pippo = pippo + 1
                End If
            Next
            myDsTemp.AcceptChanges()
        Catch ex As Exception
            Throw New Exception("Function:: DeleteRowFREEZER : COMPlusFreezer" & ": " & " " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Funzione di valorizzazione dataset con i dati della tabella di appoggio i dati principali delle unità immobiliari.
    ''' </summary>
    ''' <param name="Tributo"></param>
    ''' <param name="objDataRowSorgente"></param>
    ''' <param name="strANNO"></param>
    ''' <param name="myDataSet"></param>
    ''' <returns></returns>
    Private Function FillRowFREEZER(ByVal Tributo As String, ByVal objDataRowSorgente As DataRow, ByVal strANNO As String, ByRef myDataSet As DataSet) As Boolean
        Try
            Dim Row1 As DataRow

            Row1 = myDataSet.Tables(0).NewRow()
            Row1.Item("ANNO") = strANNO
            Row1.Item("CODTRIBUTO") = Tributo
            If Not IsDBNull(objDataRowSorgente("COD_CONTRIBUENTE")) Then
                Row1.Item("COD_CONTRIBUENTE") = CType(objDataRowSorgente("COD_CONTRIBUENTE"), Long)
            Else
                Row1.Item("COD_CONTRIBUENTE") = 0
            End If
            If Not IsDBNull(objDataRowSorgente("ID_TESTATA")) Then
                Row1.Item("ID_TESTATA") = CType(objDataRowSorgente("ID_TESTATA"), String)
            Else
                Row1.Item("ID_TESTATA") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("ID_IMMOBILE")) Then
                Row1.Item("ID_IMMOBILE") = CType(objDataRowSorgente("ID_IMMOBILE"), String)
            Else
                Row1.Item("ID_IMMOBILE") = ""
            End If
            Row1.Item("COD_TIPO_PROCEDIMENTO") = ""
            If Not IsDBNull(objDataRowSorgente("COD_ENTE")) Then
                Row1.Item("COD_ENTE") = CType(objDataRowSorgente("COD_ENTE"), String)
            Else
                Row1.Item("COD_ENTE") = ""
            End If
            Row1.Item("NUMERO_MESI_ACCONTO") = 0
            Row1.Item("NUMERO_MESI_TOTALI") = 0
            If Not IsDBNull(objDataRowSorgente("NUMEROUTILIZZATORI")) Then
                Row1.Item("NUMERO_UTILIZZATORI") = CType(objDataRowSorgente("NUMEROUTILIZZATORI"), String)
            Else
                Row1.Item("NUMERO_UTILIZZATORI") = 0
            End If
            'abitazione principale combo
            If Not IsDBNull(objDataRowSorgente("FLAG_PRINCIPALE")) Then
                Row1.Item("FLAG_PRINCIPALE") = CType(objDataRowSorgente("FLAG_PRINCIPALE"), String)
            Else
                Row1.Item("FLAG_PRINCIPALE") = 1
            End If
            If Not IsDBNull(objDataRowSorgente("PERC_POSSESSO")) Then
                Row1.Item("PERC_POSSESSO") = CType(objDataRowSorgente("PERC_POSSESSO"), Double)
            Else
                Row1.Item("PERC_POSSESSO") = 0
            End If
            If Not IsDBNull(objDataRowSorgente("VALORE")) Then
                Row1.Item("VALORE") = CType(objDataRowSorgente("VALORE"), Double)
            Else
                Row1.Item("VALORE") = 0
            End If
            If Not IsDBNull(objDataRowSorgente("FOGLIO")) Then
                Row1.Item("FOGLIO") = CType(objDataRowSorgente("FOGLIO"), String)
            Else
                Row1.Item("FOGLIO") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("NUMERO")) Then
                Row1.Item("NUMERO") = CType(objDataRowSorgente("NUMERO"), String)
            Else
                Row1.Item("NUMERO") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("SUBALTERNO")) Then
                Row1.Item("SUBALTERNO") = CType(objDataRowSorgente("SUBALTERNO"), String)
            Else
                Row1.Item("SUBALTERNO") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("MESIPOSSESSO")) Then
                Row1.Item("MESIPOSSESSO") = CType(objDataRowSorgente("MESIPOSSESSO"), Integer)
            Else
                Row1.Item("MESIPOSSESSO") = 0
            End If
            'abitazione principale checkbox                     
            If Not IsDBNull(objDataRowSorgente("AbitazionePrincipaleAttuale")) Then
                If objDataRowSorgente("AbitazionePrincipaleAttuale") = 1 Then
                    Row1.Item("FLAG_PRINCIPALE") = 1
                Else
                    Row1.Item("FLAG_PRINCIPALE") = 0
                End If
            Else
                Row1.Item("FLAG_PRINCIPALE") = 0
            End If

            'riduzione
            '0->si
            '1->no
            '2->non compilato
            If Not IsDBNull(objDataRowSorgente("RIDUZIONE")) Then
                Row1.Item("RIDUZIONE") = objDataRowSorgente("RIDUZIONE")
            Else
                Row1.Item("RIDUZIONE") = 1
            End If
            'esente/escluso
            If Not IsDBNull(objDataRowSorgente("ESCLUSIONEESENZIONE")) Then
                Row1.Item("ESENTE_ESCLUSO") = objDataRowSorgente("ESCLUSIONEESENZIONE")
            Else
                Row1.Item("ESENTE_ESCLUSO") = 1
            End If
            '*** 20140509 - TASI ***
            'tipo possesso
            'If Not IsDBNull(objDataRowSorgente("TIPOPOSSESSO")) Then
            '    Row1.Item("TIPO_POSSESSO") = CType(objDataRowSorgente("TIPOPOSSESSO"), String)
            'Else
            '    Row1.Item("TIPO_POSSESSO") = ""
            'End If
            If Not IsDBNull(objDataRowSorgente("IDTIPOUTILIZZO")) Then
                Row1.Item("IDTIPOUTILIZZO") = CType(objDataRowSorgente("IDTIPOUTILIZZO"), String)
            Else
                Row1.Item("IDTIPOUTILIZZO") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("IDTIPOPOSSESSO")) Then
                Row1.Item("IDTIPOPOSSESSO") = CType(objDataRowSorgente("IDTIPOPOSSESSO"), String)
            Else
                Row1.Item("IDTIPOPOSSESSO") = ""
            End If
            '*** ***
            'importo detrazione
            If Not IsDBNull(objDataRowSorgente("ImpDetrazAbitazPrincipale")) Then
                Row1.Item("IMPORTO_DETRAZIONE") = CType(objDataRowSorgente("ImpDetrazAbitazPrincipale"), Double)
            Else
                Row1.Item("IMPORTO_DETRAZIONE") = 0
            End If
            If Not IsDBNull(objDataRowSorgente("RENDITA")) Then
                Row1.Item("RENDITA") = CType(objDataRowSorgente("RENDITA"), String)
            Else
                Row1.Item("RENDITA") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("DataInizio")) Then
                Row1.Item("DATA_INIZIO") = CType(objDataRowSorgente("DataInizio"), DateTime)
            Else
                Row1.Item("DATA_INIZIO") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("DataFine")) Then
                Row1.Item("DATA_FINE") = CType(objDataRowSorgente("DataFine"), DateTime)
            Else
                Row1.Item("DATA_FINE") = ""
            End If
            If Not IsDBNull(objDataRowSorgente("DataFine")) Then
                Dim miadata As Date = New Date(CInt(strANNO), 12, 31)
                If DateDiff(DateInterval.Day, miadata, Row1.Item("DATA_FINE")) >= 0 Then
                    Row1.Item("POSSESSO_FINE_ANNO") = 1
                Else
                    Row1.Item("POSSESSO_FINE_ANNO") = 0
                End If
            Else
                Row1.Item("POSSESSO_FINE_ANNO") = 1
            End If
            If Not IsDBNull(objDataRowSorgente("IDIMMOBILEPERTINENTE")) And objDataRowSorgente("IDIMMOBILEPERTINENTE") <> "-1" Then
                Row1.Item("COD_IMMOBILE_PERTINENZA") = CType(objDataRowSorgente("IDIMMOBILEPERTINENTE"), String)
            Else
                Row1.Item("COD_IMMOBILE_PERTINENZA") = ""
            End If

            If Not IsDBNull(objDataRowSorgente("ID_IMMOBILE")) Then
                Row1.Item("COD_IMMOBILE_DA_ACCERTAMENTO") = CType(objDataRowSorgente("ID_IMMOBILE"), String)
            Else
                Row1.Item("COD_IMMOBILE_DA_ACCERTAMENTO") = ""
            End If
            '*** 20150430 - TASI Inquilino ***
            Row1.Item("TIPOTASI") = CType(objDataRowSorgente("TIPOTASI"), String)
            Row1.Item("IDCONTRIBUENTECALCOLO") = CType(objDataRowSorgente("IDCONTRIBUENTECALCOLO"), Integer)
            '*** ***
            Row1.Item("IDCONTRIBUENTEDICH") = CType(objDataRowSorgente("IDCONTRIBUENTEDICH"), Integer)
            myDataSet.Tables(0).Rows.Add(Row1)
            myDataSet.AcceptChanges()
            Return True
        Catch ex As Exception
            Log.Debug("Si è verificato un errore in Function FillRowFREEZER" & ex.Message)
            Throw New Exception
        End Try
    End Function

    Private Sub FillObjFREEZER(ByVal strAnnoFreezer As String, ByVal strAnnoDichiarazioneTrovata As String, ByVal dsDetraz As DataSet, ByVal blnConfigurazioneDich As Boolean, ByVal myDataSet As DataSet, ByRef myDsFinale As DataSet)
        'Dim OBJDS As DataSet
        Dim myNewRow As DataRow
        'Dim iCount As Integer
        Dim culture As IFormatProvider
        Dim myCalcolo As New CALCOLO_ICI

        Try
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            For Each myRow As DataRow In myDataSet.Tables(0).Rows
                myNewRow = myDsFinale.Tables(0).NewRow()
                If blnConfigurazioneDich = True Then
                    myNewRow.Item("ANNO") = strAnnoFreezer
                Else
                    myNewRow.Item("ANNO") = CType(myRow.Item("ANNO"), String)
                End If
                If Not IsDBNull(myRow.Item("CODTRIBUTO")) Then
                    myNewRow.Item("CODTRIBUTO") = CType(myRow.Item("CODTRIBUTO"), String)
                Else
                    myNewRow.Item("CODTRIBUTO") = Utility.Costanti.TRIBUTO_ICI
                End If
                If Not IsDBNull(myRow.Item("COD_CONTRIBUENTE")) Then
                    myNewRow.Item("COD_CONTRIBUENTE") = CType(myRow.Item("COD_CONTRIBUENTE"), Long)
                Else
                    myNewRow.Item("COD_CONTRIBUENTE") = 0
                End If

                If Not IsDBNull(myRow.Item("ID_TESTATA")) Then
                    myNewRow.Item("ID_TESTATA") = CType(myRow.Item("ID_TESTATA"), String)
                Else
                    myNewRow.Item("ID_TESTATA") = ""
                End If

                If Not IsDBNull(myRow.Item("ID_IMMOBILE")) Then
                    myNewRow.Item("ID_IMMOBILE") = CType(myRow.Item("ID_IMMOBILE"), String)
                Else
                    myNewRow.Item("ID_IMMOBILE") = ""
                End If

                myNewRow.Item("COD_TIPO_PROCEDIMENTO") = "L"                ' CType(myrow.Item("COD_TIPO_PROCEDIMENTO"), String)
                If Not IsDBNull(myRow.Item("COD_ENTE")) Then
                    myNewRow.Item("COD_ENTE") = CType(myRow.Item("COD_ENTE"), String)
                Else
                    myNewRow.Item("COD_ENTE") = ""
                End If

                If Not IsDBNull(myRow.Item("NUMERO_UTILIZZATORI")) Then
                    myNewRow.Item("NUMERO_UTILIZZATORI") = CType(myRow.Item("NUMERO_UTILIZZATORI"), Integer)
                Else
                    myNewRow.Item("NUMERO_UTILIZZATORI") = 0
                End If
                If Not IsDBNull(myRow.Item("PERC_POSSESSO")) Then
                    myNewRow.Item("PERC_POSSESSO") = CType(myRow.Item("PERC_POSSESSO"), Double)
                Else
                    myNewRow.Item("PERC_POSSESSO") = 0
                End If
                If Not IsDBNull(myRow.Item("VALORE")) Then
                    If blnConfigurazioneDich = True Then
                        myNewRow.Item("VALORE") = GetValore(CType(myRow.Item("VALORE"), Double), myRow.Item("RENDITA"), strAnnoFreezer, strAnnoDichiarazioneTrovata)
                    Else
                        myNewRow.Item("VALORE") = GetValore(CType(myRow.Item("VALORE"), Double), myRow.Item("RENDITA"), myNewRow.Item("ANNO"), strAnnoDichiarazioneTrovata)
                    End If
                Else
                    myNewRow.Item("VALORE") = 0
                End If
                If Not IsDBNull(myRow.Item("POSSESSO_FINE_ANNO")) Then
                    myNewRow.Item("POSSESSO_FINE_ANNO") = CType(myRow.Item("POSSESSO_FINE_ANNO"), String)
                Else
                    myNewRow.Item("POSSESSO_FINE_ANNO") = 0
                End If
                If Not IsDBNull(myRow.Item("FOGLIO")) Then
                    myNewRow.Item("FOGLIO") = CType(myRow.Item("FOGLIO"), String)
                Else
                    myNewRow.Item("FOGLIO") = ""
                End If
                If Not IsDBNull(myRow.Item("NUMERO")) Then
                    myNewRow.Item("NUMERO") = CType(myRow.Item("NUMERO"), String)
                Else
                    myNewRow.Item("NUMERO") = ""
                End If
                If Not IsDBNull(myRow.Item("SUBALTERNO")) Then
                    myNewRow.Item("SUBALTERNO") = CType(myRow.Item("SUBALTERNO"), String)
                Else
                    myNewRow.Item("SUBALTERNO") = ""
                End If
                Dim objRowFinale As DataRow()
                If blnConfigurazioneDich = True Then
                    objRowFinale = myDsFinale.Tables(0).Select("FOGLIO='" & myNewRow.Item("FOGLIO") & "' AND NUMERO='" & myNewRow.Item("NUMERO") & "' AND SUBalterno='" & myNewRow.Item("SUBALTERNO") & "' AND ANNO='" & strAnnoFreezer - 1 & "' AND POSSESSO_FINE_ANNO=true")
                    If objRowFinale.Length > 0 And myNewRow.Item("POSSESSO_FINE_ANNO") = True Then
                        myNewRow.Item("MESIPOSSESSO") = 12
                        myNewRow.Item("NUMERO_MESI_TOTALI") = 12
                        myNewRow.Item("NUMERO_MESI_ACCONTO") = 6
                    Else
                        If Not IsDBNull(myRow.Item("MESIPOSSESSO")) Then
                            myNewRow.Item("MESIPOSSESSO") = CType(myRow.Item("MESIPOSSESSO"), Integer)
                        Else
                            myNewRow.Item("MESIPOSSESSO") = 0
                        End If
                        If Not IsDBNull(myRow.Item("NUMERO_MESI_ACCONTO")) Then
                            myNewRow.Item("NUMERO_MESI_ACCONTO") = CType(myRow.Item("NUMERO_MESI_ACCONTO"), Integer)
                        Else
                            myNewRow.Item("NUMERO_MESI_ACCONTO") = 0
                        End If
                        If Not IsDBNull(myRow.Item("NUMERO_MESI_TOTALI")) Then
                            myNewRow.Item("NUMERO_MESI_TOTALI") = CType(myRow.Item("NUMERO_MESI_TOTALI"), Integer)
                        Else
                            myNewRow.Item("NUMERO_MESI_TOTALI") = 0
                        End If
                    End If
                Else
                    myNewRow.Item("NUMERO_MESI_TOTALI") = CType(myRow.Item("NUMERO_MESI_TOTALI"), Integer)
                    myNewRow.Item("NUMERO_MESI_ACCONTO") = CType(myRow.Item("NUMERO_MESI_ACCONTO"), Integer)
                End If
                If Not IsDBNull(myRow.Item("FLAG_PRINCIPALE")) Then
                    myNewRow.Item("FLAG_PRINCIPALE") = CType(myRow.Item("FLAG_PRINCIPALE"), String)
                Else
                    myNewRow.Item("FLAG_PRINCIPALE") = 0
                End If
                If Not IsDBNull(myRow.Item("RIDUZIONE")) Then
                    myNewRow.Item("RIDUZIONE") = CType(myRow.Item("RIDUZIONE"), String)
                Else
                    myNewRow.Item("RIDUZIONE") = 0
                End If
                If Not IsDBNull(myRow.Item("ESENTE_ESCLUSO")) Then
                    myNewRow.Item("ESENTE_ESCLUSO") = CType(myRow.Item("ESENTE_ESCLUSO"), String)
                Else
                    myNewRow.Item("ESENTE_ESCLUSO") = 0
                End If
                '*** 20140509 - TASI ***
                'If Not IsDBNull(myrow.Item("TIPO_POSSESSO")) Then
                '    Row1.Item("TIPO_POSSESSO") = CType(myrow.Item("TIPO_POSSESSO"), String)
                'Else
                '    Row1.Item("TIPO_POSSESSO") = 0
                'End If
                If Not IsDBNull(myRow.Item("IDTIPOUTILIZZO")) Then
                    myNewRow.Item("IDTIPOUTILIZZO") = CType(myRow.Item("IDTIPOUTILIZZO"), String)
                Else
                    myNewRow.Item("IDTIPOUTILIZZO") = 0
                End If
                If Not IsDBNull(myRow.Item("IDTIPOPOSSESSO")) Then
                    myNewRow.Item("IDTIPOPOSSESSO") = CType(myRow.Item("IDTIPOPOSSESSO"), String)
                Else
                    myNewRow.Item("IDTIPOPOSSESSO") = 0
                End If
                '*** ***
                'importo detrazione
                If Not IsDBNull(myRow.Item("IMPORTO_DETRAZIONE")) Then
                    'Row1.Item("IMPORTO_DETRAZIONE") = CType(myrow.Item("IMPORTO_DETRAZIONE"), String)
                    If blnConfigurazioneDich = True Then
                        myNewRow.Item("IMPORTO_DETRAZIONE") = GestioneDetrazioni(dsDetraz, CType(myRow.Item("IMPORTO_DETRAZIONE"), Double), strAnnoDichiarazioneTrovata, strAnnoFreezer)
                    Else
                        myNewRow.Item("IMPORTO_DETRAZIONE") = GestioneDetrazioni(dsDetraz, CType(myRow.Item("IMPORTO_DETRAZIONE"), Double), strAnnoDichiarazioneTrovata, myNewRow.Item("ANNO"))
                    End If
                    'dblImportDetrazCalcolato = GestioneDetrazioni(dsDetraz, dblImportDetrazDichiarato, strAnnoDichiarazioneTrovata, strAnnoFreezer)
                Else
                    myNewRow.Item("IMPORTO_DETRAZIONE") = 0
                End If
                'ale 24052007
                If Not IsDBNull(myRow.Item("COD_IMMOBILE_PERTINENZA")) Then
                    myNewRow.Item("COD_IMMOBILE_PERTINENZA") = CType(myRow.Item("COD_IMMOBILE_PERTINENZA"), String)
                Else
                    myNewRow.Item("COD_IMMOBILE_PERTINENZA") = ""
                End If
                If Not IsDBNull(myRow.Item("COD_IMMOBILE_DA_ACCERTAMENTO")) Then
                    myNewRow.Item("COD_IMMOBILE_DA_ACCERTAMENTO") = CType(myRow.Item("COD_IMMOBILE_DA_ACCERTAMENTO"), String)
                Else
                    myNewRow.Item("COD_IMMOBILE_DA_ACCERTAMENTO") = ""
                End If
                If Not IsDBNull(myRow.Item("CONTITOLARE")) Then
                    myNewRow.Item("CONTITOLARE") = CType(myRow.Item("CONTITOLARE"), String)
                Else
                    myNewRow.Item("CONTITOLARE") = False
                End If
                '*** 20150430 - TASI Inquilino ***
                myNewRow.Item("TIPOTASI") = CType(myRow.Item("TIPOTASI"), String)
                myNewRow.Item("IDCONTRIBUENTECALCOLO") = CType(myRow.Item("IDCONTRIBUENTECALCOLO"), Integer)
                '*** ***
                myNewRow.Item("IDCONTRIBUENTEDICH") = CType(myRow.Item("IDCONTRIBUENTEDICH"), Integer)
                myDsFinale.Tables(0).Rows.Add(myNewRow)
            Next
            myDsFinale.AcceptChanges()
        Catch ex As Exception
            Log.Debug("Si è verificato un errore in Function::FillObjFREEZER::COMPlusFreezer::" & ex.Message)
            Throw New Exception("Function::FillObjFREEZER::COMPlusFreezer" & "::" & " " & ex.Message)
        End Try
    End Sub

    Private Sub DeleteRowFREEZERPossessoFineAnno(ByRef myDsTemp As DataSet)
        Try
            Dim iCount As Integer

            For iCount = 0 To myDsTemp.Tables(0).Rows.Count - 1
                If myDsTemp.Tables(0).Rows(iCount).Item("POSSESSO_FINE_ANNO") = 0 Then
                    myDsTemp.Tables(0).Rows(iCount).Delete()
                End If
            Next
            myDsTemp.AcceptChanges()
        Catch ex As Exception
            Throw New Exception("Function::DeleteRowFREEZERPossessoFineAnno::COMPlusFreezer" & "::" & " " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Determino Mesi Acconto e Totali
    ''' Se inizio nell'anno uso la data di inizio+la data di fine e calcolo i mesi altrimenti se finisco nell'anno uso inizio anno+la data di fine e calcolo i mesi altrimenti uso inizio anno+fine anno e calcolo i mesi.
    ''' Dal calcolo di mesi ottendo il mese di inizio ed il mese di fine che servono per determinare i mesi in acconto ed i mesi in saldo.
    ''' </summary>
    ''' <param name="myDsTemp"></param>
    Private Sub Update_MesiPossesso_Con_Periodo(ByRef myDsTemp As DataSet)
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")
        Dim nMesiAcconto, nMeseTotali As Integer
        Dim DataInizio, DataFine As String
        Dim FncUtil As New Generale
        Dim MeseAcconto As Integer
        Dim MeseInizio As Integer = 1
        Dim MeseFine As Integer = 12

        Try
            For Each myRow As DataRow In myDsTemp.Tables(0).Rows
                DataInizio = ""
                DataFine = ""
                If StringOperation.FormatInt(myRow("ANNO")) = StringOperation.FormatDateTime(myRow("DATA_INIZIO")).Year Then
                    DataInizio = myRow("DATA_INIZIO")
                    DataFine = myRow("DATA_FINE")
                    nMeseTotali = FncUtil.mesi_possesso(DataInizio, DataFine, StringOperation.FormatInt(myRow("ANNO")), MeseInizio, MeseFine)
                Else
                    If StringOperation.FormatInt(myRow("ANNO")) = StringOperation.FormatDateTime(myRow("DATA_FINE")).Year Then
                        DataInizio = "01/01/" & myRow("ANNO")
                        DataFine = myRow("DATA_FINE")
                        nMeseTotali = FncUtil.mesi_possesso(DataInizio, DataFine, StringOperation.FormatInt(myRow("ANNO")), MeseInizio, MeseFine)
                    Else
                        DataInizio = "01/01/" & myRow("ANNO")
                        DataFine = "31/12/" & myRow("ANNO")
                        nMeseTotali = FncUtil.mesi_possesso(DataInizio, DataFine, StringOperation.FormatInt(myRow("ANNO")), MeseInizio, MeseFine)
                    End If
                End If

                If MeseFine > 6 Then
                    MeseAcconto = 6
                Else
                    MeseAcconto = MeseFine
                End If
                nMesiAcconto = (MeseAcconto - MeseInizio) + 1
                If nMesiAcconto < 0 Then
                    nMesiAcconto = 0
                End If
                myRow.Item("NUMERO_MESI_ACCONTO") = nMesiAcconto
                myRow.Item("NUMERO_MESI_TOTALI") = nMeseTotali
            Next
            myDsTemp.AcceptChanges()
        Catch ex As Exception
            Log.Debug("ClsFreezer.Update_MesiPossesso_Utilizzatori.errore::" & ex.Message)
            Throw New Exception
        End Try
    End Sub
    'Private Sub Update_MesiPossesso_Con_Periodo(ByVal strCOD_CONTRIBUENTE As Long, ByVal strANNO As String, ByVal strAnnoDichiarazioneTrovata As String, ByRef myDsTemp As DataSet)
    '    System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")
    '    Dim iCount As Integer
    '    Dim intNumero_MesiPossesso As Integer
    '    Dim blnPossessoFineAnno As Boolean
    '    Dim nMesiAcconto, nMeseTotali As Integer
    '    Dim DataInizio, DataFine As String
    '    Dim blntrovato As Boolean = False
    '    Dim FncUtil As New Generale
    '    Dim MeseAcconto As Integer
    '    Dim MeseInizio As Integer = 1
    '    Dim MeseFine As Integer = 12

    '    Try
    '        For iCount = 0 To myDsTemp.Tables(0).Rows.Count - 1
    '            DataInizio = ""
    '            DataFine = ""
    '            blntrovato = False
    '            '**********************************************************************
    '            'Determino Mesi Acconto e Totali
    '            '**********************************************************************
    '            If myDsTemp.Tables(0).Rows(iCount)("ANNO") = DateTime.Parse(myDsTemp.Tables(0).Rows(iCount)("DATA_INIZIO")).Year Then
    '                DataInizio = myDsTemp.Tables(0).Rows(iCount)("DATA_INIZIO")
    '                DataFine = myDsTemp.Tables(0).Rows(iCount)("DATA_FINE")
    '                nMeseTotali = FncUtil.mesi_possesso(DataInizio, DataFine, myDsTemp.Tables(0).Rows(iCount)("ANNO"), MeseInizio, MeseFine)
    '            Else
    '                If myDsTemp.Tables(0).Rows(iCount)("ANNO") = DateTime.Parse(myDsTemp.Tables(0).Rows(iCount)("DATA_FINE")).Year Then
    '                    DataInizio = "01/01/" & myDsTemp.Tables(0).Rows(iCount)("ANNO")
    '                    DataFine = myDsTemp.Tables(0).Rows(iCount)("DATA_FINE")
    '                    nMeseTotali = FncUtil.mesi_possesso(DataInizio, DataFine, myDsTemp.Tables(0).Rows(iCount)("ANNO"), MeseInizio, MeseFine)
    '                Else
    '                    nMeseTotali = 12
    '                    intNumero_MesiPossesso = nMeseTotali
    '                    blntrovato = True
    '                End If
    '            End If
    '            blnPossessoFineAnno = CType(myDsTemp.Tables(0).Rows(iCount).Item("POSSESSO_FINE_ANNO"), Long)
    '            If blntrovato = True Then
    '                If blnPossessoFineAnno = True Then
    '                    DataInizio = "01/" & (12 - intNumero_MesiPossesso) + 1 & "/" & myDsTemp.Tables(0).Rows(iCount)("ANNO")
    '                    DataFine = "31/12/" & myDsTemp.Tables(0).Rows(iCount)("ANNO")
    '                Else
    '                    DataInizio = "01/01/" & myDsTemp.Tables(0).Rows(iCount)("ANNO")
    '                    If intNumero_MesiPossesso = 2 Then
    '                        If Date.IsLeapYear(myDsTemp.Tables(0).Rows(iCount)("ANNO")) = False Then
    '                            DataFine = "28" & "/" & intNumero_MesiPossesso & "/" & myDsTemp.Tables(0).Rows(iCount)("ANNO")
    '                        Else
    '                            DataFine = "29" & "/" & intNumero_MesiPossesso & "/" & myDsTemp.Tables(0).Rows(iCount)("ANNO")
    '                        End If
    '                    ElseIf intNumero_MesiPossesso = 0 Then
    '                        DataFine = DataInizio
    '                    Else
    '                        DataFine = FncUtil.giorni_mese(intNumero_MesiPossesso) & "/" & intNumero_MesiPossesso & "/" & myDsTemp.Tables(0).Rows(iCount)("ANNO")
    '                    End If
    '                End If
    '                nMeseTotali = FncUtil.mesi_possesso(DataInizio, DataFine, myDsTemp.Tables(0).Rows(iCount)("ANNO"), MeseInizio, MeseFine)
    '            End If

    '            If MeseFine > 6 Then
    '                MeseAcconto = 6
    '            Else
    '                MeseAcconto = MeseFine
    '            End If
    '            nMesiAcconto = (MeseAcconto - MeseInizio) + 1
    '            If nMesiAcconto < 0 Then
    '                nMesiAcconto = 0
    '            End If
    '            myDsTemp.Tables(0).Rows(iCount).Item("NUMERO_MESI_ACCONTO") = nMesiAcconto
    '            myDsTemp.Tables(0).Rows(iCount).Item("NUMERO_MESI_TOTALI") = nMeseTotali

    '            '**********************************************************************
    '            'Fine Calcolo Mesi
    '            '**********************************************************************
    '        Next
    '        myDsTemp.AcceptChanges()
    '    Catch ex As Exception
    '        Log.Debug("Si è verificato un errore in Function Update_MesiPossesso_Utilizzatori" & ex.Message)
    '        Throw New Exception
    '    End Try
    'End Sub
    ''' <summary>
    ''' Funzione che valorizza, per ogni riga della tabella di appoggio, l'oggetto del calcolo con i suoi dati; ricalcolo il valore.
    ''' Al termine del ciclo richiamo la funzione getCalcolo.
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="drDichiarazioni"></param>
    ''' <param name="TipoOperazione"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <returns></returns>
    ''' <revisionHistory>
    ''' <revision date="20210514">
    ''' Il valore <strong><em>non</em></strong> deve essere ricalcolato in caso di Aree Fabbricabili.
    ''' </revision>
    ''' </revisionHistory>
    Private Function addRowsCalcoloICI(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, ByVal drDichiarazioni As DataRow(), TipoOperazione As String, ByVal TipoCalcolo As Integer) As objSituazioneFinale()
        Try
            Dim fncGen As New Generale
            Dim dsImmobiliAppoggio As objSituazioneFinale()
            Dim myArray As New ArrayList()
            Dim x As Integer = 0

            For x = 0 To drDichiarazioni.Length - 1
                Dim mySituazioneFinale As New objSituazioneFinale
                mySituazioneFinale.Id = x + 1
                If Not IsDBNull(drDichiarazioni(x)("COD_CONTRIBUENTE")) Then
                    mySituazioneFinale.IdContribuente = drDichiarazioni(x)("COD_CONTRIBUENTE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("ANNO")) Then
                    mySituazioneFinale.Anno = drDichiarazioni(x)("ANNO")
                End If
                '*** 20140509 - TASI ***
                If Not IsDBNull(drDichiarazioni(x)("CODTRIBUTO")) Then
                    mySituazioneFinale.Tributo = drDichiarazioni(x)("CODTRIBUTO")
                End If
                '*** ***
                If Not IsDBNull(drDichiarazioni(x)("NUMERO_MESI_ACCONTO")) Then
                    mySituazioneFinale.AccMesi = drDichiarazioni(x)("NUMERO_MESI_ACCONTO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("NUMERO_MESI_TOTALI")) Then
                    mySituazioneFinale.Mesi = drDichiarazioni(x)("NUMERO_MESI_TOTALI")
                End If
                If Not IsDBNull(drDichiarazioni(x)("NUMERO_UTILIZZATORI")) Then
                    mySituazioneFinale.NUtilizzatori = drDichiarazioni(x)("NUMERO_UTILIZZATORI")
                End If
                If Not IsDBNull(drDichiarazioni(x)("FLAG_PRINCIPALE")) Then
                    If Boolean.Parse(drDichiarazioni(x)("FLAG_PRINCIPALE").ToString) = True Then
                        mySituazioneFinale.FlagPrincipale = 1
                    Else
                        If drDichiarazioni(x)("IDIMMOBILEPERTINENTE").ToString.Length > 0 Then
                            If drDichiarazioni(x)("IDIMMOBILEPERTINENTE").ToString = "-1" Then
                                mySituazioneFinale.FlagPrincipale = 0
                            Else
                                mySituazioneFinale.FlagPrincipale = 2
                            End If
                        ElseIf drDichiarazioni(x)("IDIMMOBILEPERTINENTE").ToString.Length = 0 Then
                            mySituazioneFinale.FlagPrincipale = 0
                        End If
                    End If
                End If
                If Not IsDBNull(drDichiarazioni(x)("PERC_POSSESSO")) Then
                    mySituazioneFinale.PercPossesso = drDichiarazioni(x)("PERC_POSSESSO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("ENTE")) Then
                    mySituazioneFinale.IdEnte = drDichiarazioni(x)("ENTE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("CARATTERISTICA")) Then
                    mySituazioneFinale.Caratteristica = drDichiarazioni(x)("CARATTERISTICA")
                End If
                If Not IsDBNull(drDichiarazioni(x)("VIA")) Then
                    mySituazioneFinale.Via = drDichiarazioni(x)("VIA")
                End If
                If Not IsDBNull(drDichiarazioni(x)("NUMEROCIVICO")) Then
                    mySituazioneFinale.NCivico = drDichiarazioni(x)("NUMEROCIVICO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("SEZIONE")) Then
                    mySituazioneFinale.Sezione = drDichiarazioni(x)("SEZIONE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("FOGLIO")) Then
                    mySituazioneFinale.Foglio = drDichiarazioni(x)("FOGLIO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("NUMERO")) Then
                    mySituazioneFinale.Numero = drDichiarazioni(x)("NUMERO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("SUBALTERNO")) Then
                    mySituazioneFinale.Subalterno = drDichiarazioni(x)("SUBALTERNO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("CODCATEGORIACATASTALE")) Then
                    mySituazioneFinale.Categoria = drDichiarazioni(x)("CODCATEGORIACATASTALE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("CODCLASSE")) Then
                    mySituazioneFinale.Classe = drDichiarazioni(x)("CODCLASSE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("STORICO")) Then
                    mySituazioneFinale.FlagStorico = drDichiarazioni(x)("STORICO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("FLAGVALOREPROVV")) Then
                    mySituazioneFinale.FlagProvvisorio = drDichiarazioni(x)("FLAGVALOREPROVV")
                End If
                If Not IsDBNull(drDichiarazioni(x)("MESIPOSSESSO")) Then
                    mySituazioneFinale.MesiPossesso = drDichiarazioni(x)("MESIPOSSESSO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("MESIESCLUSIONEESENZIONE")) Then
                    mySituazioneFinale.MesiEsenzione = drDichiarazioni(x)("MESIESCLUSIONEESENZIONE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("MESIRIDUZIONE")) Then
                    mySituazioneFinale.MesiRiduzione = drDichiarazioni(x)("MESIRIDUZIONE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("IMPDETRAZABITAZPRINCIPALE")) Then
                    mySituazioneFinale.ImpDetrazione = drDichiarazioni(x)("IMPDETRAZABITAZPRINCIPALE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("POSSESSO")) Then
                    mySituazioneFinale.FlagPosseduto = drDichiarazioni(x)("POSSESSO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("ESENTE_ESCLUSO")) Then
                    mySituazioneFinale.FlagEsente = drDichiarazioni(x)("ESENTE_ESCLUSO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("RIDUZIONE")) Then
                    mySituazioneFinale.FlagRiduzione = drDichiarazioni(x)("RIDUZIONE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("ID")) Then
                    mySituazioneFinale.IdImmobile = drDichiarazioni(x)("ID")
                End If
                If Not IsDBNull(drDichiarazioni(x)("IDIMMOBILEPERTINENTE")) Then
                    mySituazioneFinale.IdImmobilePertinenza = drDichiarazioni(x)("IDIMMOBILEPERTINENTE")
                End If
                If Not IsDBNull(drDichiarazioni(x)("DataInizio")) Then
                    mySituazioneFinale.Dal = CDate(drDichiarazioni(x)("DataInizio"))
                    mySituazioneFinale.DataInizio = fncGen.GiraData(drDichiarazioni(x)("DataInizio"))
                End If
                If Not IsDBNull(drDichiarazioni(x)("DataFine")) Then
                    mySituazioneFinale.Al = CDate(drDichiarazioni(x)("DataFine"))
                End If
                If Not IsDBNull(drDichiarazioni(x)("TIPO_RENDITA")) Then
                    mySituazioneFinale.TipoRendita = drDichiarazioni(x)("TIPO_RENDITA")
                End If
                '*** 20140509 - TASI ***
                'row("TIPO_POSSESSO") = objDSDichiarazioniTotalePerAnno(iCount)("TIPOPOSSESSO")
                If Not IsDBNull(drDichiarazioni(x)("IDTIPOUTILIZZO")) Then
                    mySituazioneFinale.IdTipoUtilizzo = drDichiarazioni(x)("IDTIPOUTILIZZO")
                End If
                If Not IsDBNull(drDichiarazioni(x)("IDTIPOPOSSESSO")) Then
                    mySituazioneFinale.IdTipoPossesso = drDichiarazioni(x)("IDTIPOPOSSESSO")
                End If
                '*** ***
                If Not IsDBNull(drDichiarazioni(x)("ZONA")) Then
                    mySituazioneFinale.Zona = drDichiarazioni(x)("ZONA")
                End If
                If Not IsDBNull(drDichiarazioni(x)("consistenza")) Then
                    mySituazioneFinale.Consistenza = drDichiarazioni(x)("consistenza")
                End If
                If Not IsDBNull(drDichiarazioni(x)("ABITAZIONEPRINCIPALEATTUALE")) Then
                    mySituazioneFinale.AbitazionePrincipaleAttuale = drDichiarazioni(x)("ABITAZIONEPRINCIPALEATTUALE")
                End If
                If IsDBNull(drDichiarazioni(x)("RENDITA")) Then
                    mySituazioneFinale.Rendita = 0
                Else
                    mySituazioneFinale.Rendita = drDichiarazioni(x)("RENDITA")
                End If
                mySituazioneFinale.TipoOperazione = TipoOperazione
                mySituazioneFinale.IdProcedimento = 0
                mySituazioneFinale.IdRiferimento = 0
                mySituazioneFinale.Provenienza = ""
                mySituazioneFinale.Protocollo = "0"
                mySituazioneFinale.MeseInizio = 0
                mySituazioneFinale.DataScadenza = ""
                '*** 20120530 - IMU ***
                'devo ricalcolare il valore aggiornato
                Dim FncValore As New ComPlusInterface.FncICI
                '*** ***
                If Not IsDBNull(drDichiarazioni(x)("COLTIVATOREDIRETTO")) Then
                    mySituazioneFinale.IsColtivatoreDiretto = drDichiarazioni(x)("COLTIVATOREDIRETTO")
                Else
                    mySituazioneFinale.IsColtivatoreDiretto = False
                End If
                If Not IsDBNull(drDichiarazioni(x)("NUMEROFIGLI")) Then
                    mySituazioneFinale.NumeroFigli = drDichiarazioni(x)("NUMEROFIGLI")
                End If
                If Not IsDBNull(drDichiarazioni(x)("PERCENTCARICOFIGLI")) Then
                    mySituazioneFinale.PercentCaricoFigli = drDichiarazioni(x)("PERCENTCARICOFIGLI")
                End If
                '*** 20120709 - IMU per AF e LC devo usare il campo valore ***
                Dim nValoreDich As Double = 0
                If Not IsDBNull(drDichiarazioni(x)("valore")) Then
                    nValoreDich = drDichiarazioni(x)("valore")
                End If
                If (mySituazioneFinale.TipoRendita <> "AF" Or mySituazioneFinale.Valore = 0) Then
                    mySituazioneFinale.Valore = FncValore.CalcoloValore(Generale.DBType, StringConnectionGOV, StringConnectionICI, mySituazioneFinale.IdEnte, mySituazioneFinale.Anno, mySituazioneFinale.TipoRendita, mySituazioneFinale.Categoria, mySituazioneFinale.Classe, mySituazioneFinale.Zona, mySituazioneFinale.Rendita, nValoreDich, mySituazioneFinale.Consistenza, mySituazioneFinale.Dal, mySituazioneFinale.IsColtivatoreDiretto)
                End If
                '*** ***
                '*** 20140509 - TASI ***
                mySituazioneFinale.ValoreReale = mySituazioneFinale.Valore
                '*** ***
                '*** 20150430 - TASI Inquilino ***
                If Not IsDBNull(drDichiarazioni(x)("TIPOTASI")) Then
                    mySituazioneFinale.TipoTasi = drDichiarazioni(x)("TIPOTASI")
                End If
                If Not IsDBNull(drDichiarazioni(x)("IDCONTRIBUENTECALCOLO")) Then
                    mySituazioneFinale.IdContribuenteCalcolo = drDichiarazioni(x)("IDCONTRIBUENTECALCOLO")
                End If
                '*** ***
                If Not IsDBNull(drDichiarazioni(x)("IDCONTRIBUENTEDICH")) Then
                    mySituazioneFinale.IdContribuenteDich = drDichiarazioni(x)("IDCONTRIBUENTEDICH")
                End If
                myArray.Add(mySituazioneFinale)
            Next
            dsImmobiliAppoggio = CType(myArray.ToArray(GetType(objSituazioneFinale)), objSituazioneFinale())

            Return getCalcolo(StringConnectionGOV, StringConnectionICI, IdEnte, TipoCalcolo, dsImmobiliAppoggio)

        Catch ex As Exception
            Log.Error("Function::addRowsCalcoloICI::CalcoloICI::" & ex.Message)
            Throw New Exception("Function::addRowsCalcoloICI::CalcoloICI::" & ex.Message)
        End Try
    End Function

    Private Function addRowsSenzaCalcoloICI(IdEnte As String, ByVal CodContrib As String, ByVal strAnno As String, ByVal Tributo As String) As objSituazioneFinale()
        Try
            Dim mySituazioneFinale As New objSituazioneFinale
            Dim objDSImmobiliAppoggio As objSituazioneFinale()
            'Log.Debug("Function::addRowsSenzaCalcoloICI")

            mySituazioneFinale.Id = 1
            mySituazioneFinale.IdContribuente = CodContrib
            mySituazioneFinale.Anno = strAnno
            '*** 20140509 - TASI ***
            mySituazioneFinale.ValoreReale = mySituazioneFinale.Valore
            '*** ***
            mySituazioneFinale.IdEnte = IdEnte
            mySituazioneFinale.DataInizio = Date.MaxValue
            '*** 20140509 - TASI ***
            mySituazioneFinale.Tributo = Tributo
            '*** **
            '*** 20150430 - TASI Inquilino ***
            mySituazioneFinale.TipoTasi = Utility.Costanti.TIPOTASI_PROPRIETARIO
            mySituazioneFinale.IdContribuenteCalcolo = 0
            ReDim Preserve objDSImmobiliAppoggio(0)
            objDSImmobiliAppoggio(0) = mySituazioneFinale

            Return objDSImmobiliAppoggio
        Catch ex As Exception
            Log.Error("Function:: addRowsSenzaCalcoloICI()::CalcoloICI::" & ex.Message)
            Throw New Exception("Function::addRowsSenzaCalcoloICI::CalcoloICI::" & ex.Message)
        End Try
    End Function

    Private Function getPercentualeAcconto(ByVal structPARAMETRI_ICI As Generale.PARAMETRI_ICI) As Double
        Dim intANNO_RIFERIMENTO As Integer

        If structPARAMETRI_ICI.strACCONTO_TOTALE.CompareTo(Generale.ACCONTO) = 0 Then
            intANNO_RIFERIMENTO = structPARAMETRI_ICI.intANNO_CALCOLO
            Select Case intANNO_RIFERIMENTO
                Case Is < Generale.ANNO_CALCOLO
                    getPercentualeAcconto = 90
                Case Is >= Generale.ANNO_CALCOLO
                    getPercentualeAcconto = 100
            End Select
        End If
        If structPARAMETRI_ICI.strACCONTO_TOTALE.CompareTo(Generale.TOTALE) = 0 Then
            getPercentualeAcconto = 100
        End If
        Return getPercentualeAcconto
    End Function

    Private Function getMesi(ByVal structPARAMETRI_ICI As Generale.PARAMETRI_ICI, ByVal intNUMERO_MESI_ACCONTO As Integer, ByVal intNUMERO_MESI_TOTALE As Integer) As Integer
        If structPARAMETRI_ICI.strACCONTO_TOTALE.CompareTo(Generale.ACCONTO) = 0 Then
            getMesi = intNUMERO_MESI_ACCONTO
        End If

        If structPARAMETRI_ICI.strACCONTO_TOTALE.CompareTo(Generale.TOTALE) = 0 Then
            getMesi = intNUMERO_MESI_TOTALE
        End If

        Return getMesi
    End Function

    Public Function CalcoloValoreImmobiliB(ByVal Anno_calcolo As Integer, ByVal rendita As Double, ByVal data_inizio As String, ByVal data_fine As String) As Double
        Dim mesi_poss As Integer
        Dim dblvalore As Double
        Dim mesi100, mesi140 As Integer
        Dim data_controllo As Date = "01/10/2006"
        Dim clsGeneralFunction As New Generale

        Try
            If Anno_calcolo < 2006 Then
                dblvalore = rendita * 100
            End If
            If Anno_calcolo > 2006 Then
                dblvalore = rendita * 140
            End If

            If Anno_calcolo = 2006 Then
                If IsDBNull(data_fine) Then
                    data_fine = "20061231"
                ElseIf data_fine > "20061231" Then
                    data_fine = "20061231"
                End If
                If data_inizio < "20060101" Then
                    data_inizio = "20060101"
                End If

                data_inizio = clsGeneralFunction.GiraDataFromDB(data_inizio)
                data_fine = clsGeneralFunction.GiraDataFromDB(data_fine)

                mesi_poss = DateDiff(DateInterval.Month, CDate(data_inizio), CDate(data_fine)) + 1

                If mesi_poss >= 12 Then
                    mesi100 = 9
                    mesi140 = 3
                Else
                    mesi100 = DateDiff(DateInterval.Month, CDate(data_inizio), CDate(data_controllo))
                    mesi140 = DateDiff(DateInterval.Month, CDate(data_controllo), CDate(data_fine)) + 1
                End If

                dblvalore = mesi100 / 12 * (rendita * 100) + mesi140 / 12 * (rendita * 140)
            End If

            If Anno_calcolo > 1997 Then
                'rivalutazione
                dblvalore = dblvalore + dblvalore * 5 / 100
            End If

            CalcoloValoreImmobiliB = dblvalore
        Catch ex As Exception
            Log.Error("CalcoloValoreImmobiliB::" & ex.Message)
            CalcoloValoreImmobiliB = 0
        End Try
    End Function

    Protected Function GetValore(ByVal Valore As Double, ByVal TipoRendita As String, ByVal Anno As String, ByVal AnnoDichOrig As String) As Double
        Try
            Dim dblValoreRet As Double

            If IsDBNull(Valore) Then
                dblValoreRet = 0
            Else

                If Anno >= "1997" And AnnoDichOrig < "1997" Then
                    Select Case UCase(TipoRendita)
                        Case "AF"
                            'non faccio nulla
                            dblValoreRet = Valore
                        Case "LC"
                            'non faccio nulla
                            dblValoreRet = Valore
                        Case ""
                            'non faccio nulla
                            dblValoreRet = Valore
                        Case "TA"
                            dblValoreRet = Valore + ((Valore * 25) / 100)
                        Case Else
                            dblValoreRet = Valore + ((Valore * 5) / 100)
                    End Select
                Else
                    dblValoreRet = Valore
                End If
            End If
            GetValore = dblValoreRet
        Catch ex As Exception
            Log.Debug("GetValore::si è verificato il seguente errore::" & ex.Message)
            GetValore = -1
        End Try
    End Function

    Private Function GestioneDetrazioni(ByVal dsDetraz As DataSet, ByVal ImportoDetrazDichiarato As Double, ByVal strAnnoOrig As String, ByVal strAnnoFreezer As String) As Double
        Dim objDetrazAnnoOrig() As DataRow
        Dim objDetrazAnnoFreezer() As DataRow
        Dim importoDetrazAnnoOrig As Double
        Dim importoDetrazAnnoFreezer As Double
        Dim ImportoDetrazioneRet As Double

        Try
            objDetrazAnnoOrig = dsDetraz.Tables(0).Select("ANNO='" & strAnnoOrig & "'")
            If objDetrazAnnoOrig.Length > 0 Then
                importoDetrazAnnoOrig = objDetrazAnnoOrig(0).Item("VALORE")
            Else
                importoDetrazAnnoOrig = 0
            End If
            objDetrazAnnoFreezer = dsDetraz.Tables(0).Select("ANNO='" & strAnnoFreezer & "'")
            If objDetrazAnnoFreezer.Length > 0 Then
                importoDetrazAnnoFreezer = objDetrazAnnoFreezer(0).Item("VALORE")
            Else
                importoDetrazAnnoFreezer = 0
            End If
            If importoDetrazAnnoOrig = 0 Or importoDetrazAnnoFreezer = 0 Then
                Return ImportoDetrazDichiarato
            End If
            ImportoDetrazioneRet = (ImportoDetrazDichiarato * importoDetrazAnnoFreezer) / importoDetrazAnnoOrig

            Return ImportoDetrazioneRet
        Catch ex As Exception
            Log.Debug("GestioneDetrazioni::si è verificato il seguente errore::" & ex.Message)
            Throw New Exception("GestioneDetrazioni::si è verificato il seguente errore::" & ex.Message)
        End Try
    End Function

    'Public Function CreateDSperCalcoloICI() As DataSet

    '    Dim objDS As DataSet = New DataSet
    '    Dim newTable As DataTable

    '    newTable = New DataTable("TP_SITUAZIONE_FINALE_ICI")
    '    Dim NewColumn As DataColumn = New DataColumn

    '    NewColumn.ColumnName = "ID_SITUAZIONE_FINALE"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    NewColumn = New DataColumn

    '    NewColumn.ColumnName = "ANNO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    NewColumn = New DataColumn

    '    NewColumn.ColumnName = "COD_ENTE"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ID_PROCEDIMENTO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ID_RIFERIMENTO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "PROVENIENZA"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "CARATTERISTICA"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "INDIRIZZO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "SEZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "FOGLIO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "NUMERO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "SUBALTERNO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "CATEGORIA"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "CLASSE"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "PROTOCOLLO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "FLAG_STORICO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "VALORE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)
    '    '*** 20140509 - TASI ***
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "VALORE_REALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)
    '    '*** ***

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "FLAG_PROVVISORIO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "PERC_POSSESSO"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "MESI_POSSESSO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "MESI_ESCL_ESENZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "MESI_RIDUZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "IMPORTO_DETRAZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "FLAG_POSSEDUTO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "FLAG_ESENTE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "FLAG_RIDUZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "FLAG_PRINCIPALE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "COD_CONTRIBUENTE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "COD_IMMOBILE_PERTINENZA"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "COD_IMMOBILE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "DAL"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "AL"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "NUMERO_MESI_ACCONTO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "NUMERO_MESI_TOTALI"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "NUMERO_UTILIZZATORI"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "TIPO_RENDITA"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "RIDUZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "MESE_INIZIO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "DATA_SCADENZA"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    '*** 20140509 - TASI ***
    '    'NewColumn = New DataColumn
    '    'NewColumn.ColumnName = "TIPO_POSSESSO"
    '    'NewColumn.DataType = System.Type.GetType("System.String")
    '    'NewColumn.DefaultValue = ""
    '    'newTable.Columns.Add(NewColumn)
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "IDTIPOUTILIZZO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "IDTIPOPOSSESSO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ZONA"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "DataInizio"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    '*** ***
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "TIPO_OPERAZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = ""
    '    newTable.Columns.Add(NewColumn)
    '    'DIPE 11/02/2009

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "CONSISTENZA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ABITAZIONE_PRINCIPALE_ATTUALE"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "RENDITA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    'DIPE 02/03/2011
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_VALORE_ALIQUOTA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    '--------------------------------------------------------------------------------------
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_SENZA_DETRAZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_DETRAZIONE_APPLICATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_ACCONTO"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_DETRAZIONE_RESIDUA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_SENZA_DETRAZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DETRAZIONE_APPLICATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DOVUTA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DETRAZIONE_RESIDUA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_SALDO"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_DETRAZIONE_SALDO"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_SENZA_DETRAZIONE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_DETRAZIONE_RESIDUA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)
    '    '*** Campi Detrazione Statale usati per non so cosa ***
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_DETRAZIONE_STATALE_CALCOLATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_DETRAZIONE_STATALE_APPLICATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_DETRAZIONE_STATALE_RESIDUA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_SALDO_DETRAZIONE_STATALE_CALCOLATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_SALDO_DETRAZIONE_STATALE_APPLICATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_SALDO_DETRAZIONE_STATALE_RESIDUA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DETRAZIONE_STATALE_CALCOLATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DETRAZIONE_STATALE_APPLICATA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DETRAZIONE_STATALE_RESIDUA"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)
    '    '*** ***
    '    '*** 20120530 - IMU ***
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "COLTIVATOREDIRETTO"
    '    NewColumn.DataType = System.Type.GetType("System.Boolean")
    '    NewColumn.DefaultValue = False
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "NUMEROFIGLI"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "PERCENTCARICOFIGLI"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_VALORE_ALIQUOTA_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_ACCONTO_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DOVUTA_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DETRAZIONE_APPLICATA_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_TOTALE_DETRAZIONE_RESIDUA_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_SALDO_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_DETRAZIONE_SALDO_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)

    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ICI_DOVUTA_DETRAZIONE_RESIDUA_STATALE"
    '    NewColumn.DataType = System.Type.GetType("System.Double")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)
    '    '*** ***
    '    '*** 20130422 - aggiornamento IMU ***
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "ID_ALIQUOTA"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)
    '    '*** ***
    '    '*** 20140509 - TASI ***
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "CODTRIBUTO"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = Utility.Costanti.TRIBUTO_ICI
    '    newTable.Columns.Add(NewColumn)
    '    '*** ***
    '    '*** 20150430 - TASI Inquilino ***
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "TIPOTASI"
    '    NewColumn.DataType = System.Type.GetType("System.String")
    '    NewColumn.DefaultValue = Utility.Costanti.TIPOTASI_PROPRIETARIO
    '    newTable.Columns.Add(NewColumn)
    '    NewColumn = New DataColumn
    '    NewColumn.ColumnName = "IDCONTRIBUENTECALCOLO"
    '    NewColumn.DataType = System.Type.GetType("System.Int64")
    '    NewColumn.DefaultValue = "0"
    '    newTable.Columns.Add(NewColumn)
    '    '*** ***
    '    objDS.Tables.Add(newTable)

    '    Return objDS
    'End Function
End Class

Public Class objFreezerRow
    Private Property _Anno As String = String.Empty
    Private Property _Codtributo As String = String.Empty
    Private Property _Cod_Contribuente As Integer = 0
    Private Property _Id_Testata As String = "0"
    Private Property _Id_Immobile As String = "0"
    Private Property _Cod_Tipo_Procedimento As String = String.Empty
    Private Property _Cod_Ente As String = String.Empty
    Private Property _Numero_Mesi_Acconto As Integer = 0
    Private Property _Numero_Mesi_Totali As Integer = 0
    Private Property _Numero_Utilizzatori As Integer = 0
    Private Property _Perc_Possesso As Double = 0
    Private Property _Valore As Double = 0
    Private Property _Possesso_Fine_Anno As Boolean = False
    Private Property _Foglio As String = String.Empty
    Private Property _Numero As String = String.Empty
    Private Property _Subalterno As String = String.Empty
    Private Property _MesiPossesso As Integer = 0
    Private Property _Flag_Principale As Integer = 1
    Private Property _Riduzione As Integer = 0
    Private Property _Esente_Escluso As Integer = 0
    Private Property _IdTipoUtilizzo As Integer = 0
    Private Property _IdTipoPossesso As Integer = 0
    Private Property _Rendita As String = String.Empty
    Private Property _Importo_Detrazione As Double = 0
    Private Property _Data_Inizio As DateTime = DateTime.MaxValue
    Private Property _Data_Fine As DateTime = DateTime.MaxValue
    Private Property _Cod_Immobile_Pertinenza As String = String.Empty
    Private Property _Cod_Immobile_Da_Accertamento As String = String.Empty
    Private Property _Contitolare As Boolean = False
    Private Property _Tipotasi As String = Utility.Costanti.TIPOTASI_PROPRIETARIO
    Private Property _IdContribuenteCalcolo As Integer = 0
    Private Property _IdContribuenteDich As Integer = 0

    Public Property Anno As String
        Get
            Return _Anno
        End Get
        Set(ByVal Value As String)
            _Anno = Value
        End Set
    End Property
    Public Property Codtributo As String
        Get
            Return _Codtributo
        End Get
        Set(ByVal Value As String)
            _Codtributo = Value
        End Set
    End Property
    Public Property Cod_Contribuente As Integer
        Get
            Return _Cod_Contribuente
        End Get
        Set(ByVal Value As Integer)
            _Cod_Contribuente = Value
        End Set
    End Property
    Public Property Id_Testata As String
        Get
            Return _Id_Testata
        End Get
        Set(ByVal Value As String)
            _Id_Testata = Value
        End Set
    End Property
    Public Property Id_Immobile As String
        Get
            Return _Id_Immobile
        End Get
        Set(ByVal Value As String)
            _Id_Immobile = Value
        End Set
    End Property
    Public Property Cod_Tipo_Procedimento As String
        Get
            Return _Cod_Tipo_Procedimento
        End Get
        Set(ByVal Value As String)
            _Cod_Tipo_Procedimento = Value
        End Set
    End Property
    Public Property Cod_Ente As String
        Get
            Return _Cod_Ente
        End Get
        Set(ByVal Value As String)
            _Cod_Ente = Value
        End Set
    End Property
    Public Property Numero_Mesi_Acconto As Integer
        Get
            Return _Numero_Mesi_Acconto
        End Get
        Set(ByVal Value As Integer)
            _Numero_Mesi_Acconto = Value
        End Set
    End Property
    Public Property Numero_Mesi_Totali As Integer
        Get
            Return _Numero_Mesi_Totali
        End Get
        Set(ByVal Value As Integer)
            _Numero_Mesi_Totali = Value
        End Set
    End Property
    Public Property Numero_Utilizzatori As Integer
        Get
            Return _Numero_Utilizzatori
        End Get
        Set(ByVal Value As Integer)
            _Numero_Utilizzatori = Value
        End Set
    End Property
    Public Property Perc_Possesso As Double
        Get
            Return _Perc_Possesso
        End Get
        Set(ByVal Value As Double)
            _Perc_Possesso = Value
        End Set
    End Property
    Public Property Valore As Double
        Get
            Return _Valore
        End Get
        Set(ByVal Value As Double)
            _Valore = Value
        End Set
    End Property
    Public Property Possesso_Fine_Anno As Boolean
        Get
            Return _Possesso_Fine_Anno
        End Get
        Set(ByVal Value As Boolean)
            _Possesso_Fine_Anno = Value
        End Set
    End Property
    Public Property Foglio As String
        Get
            Return _Foglio
        End Get
        Set(ByVal Value As String)
            _Foglio = Value
        End Set
    End Property
    Public Property Numero As String
        Get
            Return _Numero
        End Get
        Set(ByVal Value As String)
            _Numero = Value
        End Set
    End Property
    Public Property Subalterno As String
        Get
            Return _Subalterno
        End Get
        Set(ByVal Value As String)
            _Subalterno = Value
        End Set
    End Property
    Public Property MesiPossesso As Integer
        Get
            Return _MesiPossesso
        End Get
        Set(ByVal Value As Integer)
            _MesiPossesso = Value
        End Set
    End Property
    Public Property Flag_Principale As Integer
        Get
            Return _Flag_Principale
        End Get
        Set(ByVal Value As Integer)
            _Flag_Principale = Value
        End Set
    End Property
    Public Property Riduzione As Integer
        Get
            Return _Riduzione
        End Get
        Set(ByVal Value As Integer)
            _Riduzione = Value
        End Set
    End Property
    Public Property Esente_Escluso As Integer
        Get
            Return _Esente_Escluso
        End Get
        Set(ByVal Value As Integer)
            _Esente_Escluso = Value
        End Set
    End Property
    Public Property IdTipoUtilizzo As Integer
        Get
            Return _IdTipoUtilizzo
        End Get
        Set(ByVal Value As Integer)
            _IdTipoUtilizzo = Value
        End Set
    End Property
    Public Property IdTipoPossesso As Integer
        Get
            Return _IdTipoPossesso
        End Get
        Set(ByVal Value As Integer)
            _IdTipoPossesso = Value
        End Set
    End Property
    Public Property Rendita As String
        Get
            Return _Rendita
        End Get
        Set(ByVal Value As String)
            _Rendita = Value
        End Set
    End Property
    Public Property Importo_Detrazione As Double
        Get
            Return _Importo_Detrazione
        End Get
        Set(ByVal Value As Double)
            _Importo_Detrazione = Value
        End Set
    End Property
    Public Property Data_Inizio As DateTime
        Get
            Return _Data_Inizio
        End Get
        Set(ByVal Value As DateTime)
            _Data_Inizio = Value
        End Set
    End Property
    Public Property Data_Fine As DateTime
        Get
            Return _Data_Fine
        End Get
        Set(ByVal Value As DateTime)
            _Data_Fine = Value
        End Set
    End Property
    Public Property Cod_Immobile_Pertinenza As String
        Get
            Return _Cod_Immobile_Pertinenza
        End Get
        Set(ByVal Value As String)
            _Cod_Immobile_Pertinenza = Value
        End Set
    End Property
    Public Property Cod_Immobile_Da_Accertamento As String
        Get
            Return _Cod_Immobile_Da_Accertamento
        End Get
        Set(ByVal Value As String)
            _Cod_Immobile_Da_Accertamento = Value
        End Set
    End Property
    Public Property Contitolare As Boolean
        Get
            Return _Contitolare
        End Get
        Set(ByVal Value As Boolean)
            _Contitolare = Value
        End Set
    End Property
    Public Property Tipotasi As String
        Get
            Return _Tipotasi
        End Get
        Set(ByVal Value As String)
            _Tipotasi = Value
        End Set
    End Property
    Public Property IdContribuenteCalcolo As Integer
        Get
            Return _IdContribuenteCalcolo
        End Get
        Set(ByVal Value As Integer)
            _IdContribuenteCalcolo = Value
        End Set
    End Property
    Public Property IdContribuenteDich As Integer
        Get
            Return _IdContribuenteDich
        End Get
        Set(ByVal Value As Integer)
            _IdContribuenteDich = Value
        End Set
    End Property
End Class
