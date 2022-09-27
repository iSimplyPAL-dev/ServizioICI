Imports log4net
Imports System.Data.SqlClient
Imports ComPlusInterface
Imports Utility
''' <summary>
''' 
''' </summary>
Public Class ListAnagrafica
    ''' <summary>
    ''' 
    ''' </summary>
    Public p_dsItemsANAGRAFICA As DataSet = Nothing
    ''' <summary>
    ''' 
    ''' </summary>
    Public p_daItemsANAGRAFICA As SqlDataAdapter = Nothing
End Class
''' <summary>
''' 
''' </summary>
Public Class ListALIQUOTA_DETRAZIONE
    Public p_dblVALORE_ALIQUOTA As Double
    Public p_dblVALORE_DETRAZIONE As Double = 0
    Public p_ESENTE As Integer = 0
    Public nDetrazioneFigli As Double = 0
    Public AliquotaStatale As Double = 0
    '*** 20130422 - aggiornamento IMU ***
    Public nIdAliquota As Integer = 0
    '*** ***
    '*** 20140509 - TASI ***
    Public nSogliaRendita As Double = 0
    Public sTipoSoglia As String = ">"
    '*** ***
    '*** 20150430 - TASI Inquilino ***
    Public nPercInquilino As Double = 0
    '*** ***
    Public TipoAliquota As String = ""
End Class
Public Class ObjResiduoRendita
    Private _CodTributo As String = Utility.Costanti.TRIBUTO_ICI
    Private _IdPrincipale As Integer = 0
    Private _impResiduo As Double = 0
    Private _sTipoSoglia As String = "<"

    Public Property CodTributo() As String
        Get
            Return _CodTributo
        End Get
        Set(ByVal value As String)
            _CodTributo = value
        End Set
    End Property
    Public Property IdPrincipale() As Integer
        Get
            Return _IdPrincipale
        End Get
        Set(ByVal value As Integer)
            _IdPrincipale = value
        End Set
    End Property
    Public Property impResiduo() As Double
        Get
            Return _impResiduo
        End Get
        Set(ByVal value As Double)
            _impResiduo = value
        End Set
    End Property
    Public Property sTipoSoglia() As String
        Get
            Return _sTipoSoglia
        End Get
        Set(ByVal value As String)
            _sTipoSoglia = value
        End Set
    End Property
End Class
''' <summary>
''' 
''' </summary>
''' <revisionHistory><revision date="11/06/2021">Nuove Tipologie di Utilizzo</revision></revisionHistory>
Public Class Generale
    Private Shared Log As ILog = LogManager.GetLogger(GetType(Generale))
    Public Const DBType = "SQL"
    Public Const INIT_VALUE_NUMBER As Integer = -1
    Public Const VALUE_NUMBER_ZERO As Integer = 0
    Public Const ACCONTO As String = "ACCONTO"
    Public Const TOTALE As String = "TOTALE"
    Public Const NETTOVERSATO As Integer = 1
    Public Const ANNO_CALCOLO As Integer = 2001
#Region "Tipo Rendita"
    'TIPO RENDITA -- stesse voci presenti nella tabella TAB_TIPO_RENDITA di OPENgovTerritorio
    Public Const TipoRendita_RE As String = "RE"       'RENDITA EFFETTIVA
    Public Const TipoRendita_RP As String = "RP"       'RENDITA PRESUNTA
    Public Const TipoRendita_RPM As String = "RPM"     'RENDITA PRESUNTA MODIFICATA
    Public Const TipoRendita_AF As String = "AF"       'AREE EDIFICABILI
    Public Const TipoRendita_LC As String = "LC"       'LIBRI CONTABILI
    Public Const TipoRendita_TA As String = "TA"       'TERRENI AGRICOLI
#End Region
#Region "Tipo Aliquote"
    'TIPO ALIQUOTE -- stesse voci presenti nella tabella TP_ALIQUOTE_ICI di OPENgovProvvedimenti
    Public Const TipoAliquote_A As String = "A"    'ALTRI FABBRICATI
    Public Const TipoAliquote_AAF As String = "AAF"    'AREE EDIFICABILI
    Public Const TipoAliquote_AAP As String = "AAP"    'ABITAZIONE PRINCIPALE
    Public Const TipoAliquote_AC As String = "AC"     'AFFITTI CONVENZIONATI
    Public Const TipoAliquote_AUG1 As String = "AUG1"     'USO GRATUITO 1° GRADO
    Public Const TipoAliquote_AUG2 As String = "AUG2"     'USO GRATUITO 2° GRADO
    Public Const TipoAliquote_AUG3 As String = "AUG3"     'USO GRATUITO 3° GRADO
    Public Const TipoAliquote_P As String = "P"       'PERTINENZA
    Public Const TipoAliquote_S As String = "S"       'SFITTI/A DISPOSIZIONE
    Public Const TipoAliquote_TTAA As String = "TA"       'TERRENI AGRICOLI
    Public Const TipoAliquote_STO As String = "STO"       'Storico
    Public Const TipoAliquote_RUR As String = "RUR"       'Rurale/Fabbricati Strumentali Agricoli
    Public Const TipoAliquote_IACP As String = "IACP"       'IMU per l'Agenzia Territoriale per la Casa (ex IACP)
    'DIPE 25/03/2009
    Public Const TipoAliquote_AAIRE As String = "AAIRE"       'AIRE
    'DIPE 15/09/2009 Per Pomarance
    Public Const TipoAliquote_BO As String = "BO"     'Immobili C/1 e C/3
    Public Const TipoAliquote_DSAAP As String = "DSAAP"       'DETRAZIONE STATALE ABITAZIONE PRINCIPALE
    'DIPE 23/06/2010
    Public Const TipoAliquote_AAPEX As String = "APEX"
    '*** 20120530 - IMU ***
    Public Const TipoAliquote_CD As String = "CD"     'coltivatore diretto
    Public Const TipoAliquote_DFAAP As String = "DFAAP"  'DETRAZIONE FIGLI MINORI DI 26 ANNI
    '*** ***
    '*** 20130422 - aggiornamento IMU ****
    Public Const TipoAliquote_AFD As String = "AFD"          'altri fabbricati categoria D
    Public Const TipoAliquote_CDD10 As String = "CDD10"          'coltivatore diretto categoria D/10
    '*** ***
    'PER LE DETRAZIONI STABILISCO UN PREFISSO 'D' DA CONCATENARE A TUTTE LE TIPOLOGIE DI ALIQUOTE
    Public Const TipoAliquote_D As String = "D"
    '*** 20140509 - TASI ***
    Public Const TipoAliquote_AS As String = "AS" 'Abitazione Signorile A/1 A/8 A/9
    Public Const TipoAliquote_C2 As String = "C2" 'Immobili C/2
    Public Const TipoAliquote_C6 As String = "C6" 'Immobili C/6
    '*** ***
    '*** 201801 - aliquote specifiche sui D ***
    Public Const TipoAliquote_D1 As String = "D1" 'Immobili D/1
    Public Const TipoAliquote_D5 As String = "D5" 'Immobili D/5
    Public Const TipoAliquote_D8 As String = "D8" 'Immobili D/8
    '*** ***
    Public Const TipoAliquote_LO As String = "LO"     'Immobili Locati
#End Region
#Region "Titolo Possesso"
    'TITOLO DI POSSESSO -- stesse voci presenti nella tabella TAB_TIPO_POSSESSO di OPENgovTerritorio 
    'ci sono solo i titoli di possesso che influiscono sul calcolo ICI
    Public Const TitoloPossesso_MANCANTE As Integer = 1  '[MANCANTE]
    Public Const TitoloPossesso_AP As Integer = 2  'Abitazione principale
    Public Const TitoloPossesso_UG1 As Integer = 3  'uso gratuito primo grado
    Public Const TitoloPossesso_UG2 As Integer = 4  'uso gratuito secondo grado
    Public Const TitoloPossesso_UG3 As Integer = 5  'uso gratuito terzo grado
    Public Const TitoloPossesso_SAD As Integer = 6  'a disposizione
    Public Const TitoloPossesso_AFC As Integer = 7  'affitti convenzionati
    Public Const TitoloPossesso_LOC As Integer = 8  'Immobile Locato
    Public Const TitoloPossesso_APEX As Integer = 9  'Beni Merci 'Abitazione principale ex 104/92
    Public Const TitoloPossesso_MICRO As Integer = 10  'Contribuente in micro comunità
    Public Const TitoloPossesso_VUOTO As Integer = 11  'Non occupato e vuoto
    Public Const TitoloPossesso_STO As Integer = 12  'Storico	
    Public Const TitoloPossesso_RUR As Integer = 13  'Rurale	
    Public Const TitoloPossesso_IACP As Integer = 14  'IMU per l'Agenzia Territoriale per la Casa (ex IACP)	
    Public Const TitoloPossesso_AIRE As Integer = 15  'A.I.R.E. (Abitazione Residente all'Estero)
    Public Const TitoloPossesso_CD As Integer = 16  'Coltivatore Diretto
    Public Const TitoloPossesso_CDD10 As Integer = 17  'Cat.D/10 su Coltivatore Diretto
    Public Const TitoloPossesso_PS As Integer = 18 'Pertinenza di Abitazione Signorile
#End Region

    Enum ABITAZIONE_PRINCIPALE_PERTINENZA
        ABITAZIONE_PRINCIPALE = 1
        ABITAZIONE_PERTINENZA = 2
        NO_ABITAZIONE_PERTINENZA = 0          'Ne abitazione Principale ne Pertinenza
    End Enum
    Structure PARAMETRI_ICI
        Dim intTIPO_ABITAZIONE As Integer
        Dim strTIPO_RENDITA As String
        Dim strACCONTO_TOTALE As String
        Dim intANNO_CALCOLO As Integer
        'GIULIA 20060619
        Dim intTITOLO_POSSESSO As Integer
        Dim strCATEGORIA As String
        Dim dblDETRAZIONE_DICHIARATA As Double
        '*** 20140509 - TASI ***
        'Dim strTIPO_POSSESSO As String
        Dim IdTipoUtilizzo As Integer
        '*** ***
        Dim IsColtivatoreDiretto As Boolean
        '*** 201805 - se la pertinenza è riferita ad una principale esclusa devo esludere anche lei ***
        Dim Categoria_AAP As String
        Dim Categoria_AS As String
    End Structure
    Structure VALORI_ICI_CALCOLATA
        Dim dblICI_ACCONTO_SENZA_DETRAZIONE As Double
        Dim dblICI_ACCONTO_DETRAZIONE_APPLICATA As Double
        Dim dblICI_ACCONTO_DOVUTA As Double
        Dim dblICI_ACCONTO_DETRAZIONE_RESIDUA As Double

        Dim dblICI_TOTALE_SENZA_DETRAZIONE As Double
        Dim dblICI_TOTALE_DETRAZIONE_APPLICATA As Double
        Dim dblICI_TOTALE_DOVUTA As Double
        Dim dblICI_TOTALE_DETRAZIONE_RESIDUA As Double

        Dim dblICI_SALDO_SENZA_DETRAZIONE As Double
        Dim dblICI_SALDO_DETRAZIONE_APPLICATA As Double
        Dim dblICI_SALDO_DOVUTA As Double
        Dim dblICI_SALDO_DETRAZIONE_RESIDUA As Double

        'dipe 25/02/2011
        Dim dblICI_VALORE_ALIQUOTA As Double
        '*** 20120530 - IMU***
        Dim dblICI_VALORE_ALIQUOTA_STATALE As Double
        Dim dblICI_ACCONTO_DOVUTA_STATALE As Double
        Dim dblICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE As Double
        Dim dblICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE As Double
        Dim dblICI_SALDO_DOVUTA_STATALE As Double
        Dim dblICI_SALDO_DETRAZIONE_RESIDUA_STATALE As Double
        Dim dblICI_SALDO_DETRAZIONE_APPLICATA_STATALE As Double
        Dim dblICI_TOTALE_DOVUTA_STATALE As Double
        Dim dblICI_TOTALE_DETRAZIONE_RESIDUA_STATALE As Double
        Dim dblICI_TOTALE_DETRAZIONE_APPLICATA_STATALE As Double
        '*** ***
        '*** 20130422 - aggiornamento IMU ***
        Dim nIdAliquota As Integer
        '*** ***
    End Structure

    'Public glbmese_inizio_p, glbmese_fine_p, glbmese_inizio_s, glbmese_fine_s As Integer

    '*******************************************************
    '
    ' CStrToDB() Function <a name="CStrToDB"></a>
    '
    ' Il metodo CStrToDB ritorna una stringa
    ' Viene utilizzato quando il valore di una stringa deve essere contatenata ad una stringa SQL
    ' 
    '*******************************************************
    Public Function CToStr(ByVal vInput As Object, ByRef blnClearSpace As Boolean, ByVal blnUseNull As Boolean, ByVal bUseApici As Boolean) As String
        Dim myRet As String = ""

        Try
            If blnUseNull Then
                myRet = "Null"
            End If

            If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
                myRet = CStr(vInput)
                If blnClearSpace Then
                    myRet = Trim(myRet)
                End If
                If Trim(myRet) <> "" Then
                    myRet = Replace(myRet, "'", "''")
                Else
                    myRet = ""
                End If
            End If
            If bUseApici = True Then
                myRet = "'" & myRet & "'"
            End If
        Catch ex As Exception
            Log.Debug("Utility::CStrToDB:si è verificato il seguente errore::", ex)
        End Try
        Return myRet
    End Function

    Public Function CToInt(ByVal objInput As Object) As Integer
        Dim myRet As Integer = 0

        Try
            If Not IsDBNull(objInput) And Not IsNothing(objInput) Then
                If IsNumeric(objInput) Then
                    myRet = Convert.ToInt32(objInput)
                End If
            End If
        Catch ex As Exception
            Log.Debug("Utility::CToInt:si è verificato il seguente errore::", ex)
        End Try
        Return myRet
    End Function

    Public Function CToBit(ByVal vInput As Object) As Short
        Dim myRet As Short = 0

        Try
            If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
                If CBool(vInput) Then
                    myRet = 1
                End If
            End If
        Catch ex As Exception
            Log.Debug("Utility::CToBit:si è verificato il seguente errore::", ex)
        End Try
        Return myRet
    End Function

    Public Function CToBool(ByVal vInput As Object) As Boolean
        Dim myRet As Boolean = False

        Try
            If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
                myRet = Convert.ToBoolean(vInput)
            End If
        Catch ex As Exception
            Log.Debug("Utility::CToBool:si è verificato il seguente errore::", ex)
        End Try
        Return myRet
    End Function

    Public Function CToDouble(ByVal vInput As Object, ByVal bReplaceSeparator As Boolean) As String
        Dim myRet As String = "0"

        Try
            If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
                If CStr(vInput) <> "" Then
                    myRet = CStr(vInput)
                    If bReplaceSeparator = True Then
                        myRet = Replace(myRet, ".", "")
                        myRet = Replace(myRet, ",", ".")
                    End If
                End If
            End If
        Catch ex As Exception
            Log.Debug("Utility::CToDouble:si è verificato il seguente errore::", ex)
        End Try
        Return myRet
    End Function

    Public Function GiraData(ByVal data As Object) As String
        'leggo la data nel formato gg/mm/aaaa e la metto nel formato aaaammgg
        Dim Giorno As String
        Dim Mese As String
        Dim Anno As String

        GiraData = ""
        data = CToStr(data, True, False, False)
        If Not IsDBNull(data) And Not IsNothing(data) Then
            If data <> "" Then
                Giorno = Right("0" & Mid(data, 1, 2), 2)
                Mese = Right("0" & Mid(data, 4, 2), 2)
                Anno = Mid(data, 7, 4)
                GiraData = Anno & Mese & Giorno
            End If
        End If
        Return GiraData
    End Function

    Public Function GiraDataFromDB(ByVal data As Object) As String
        'leggo la data nel formato aaaammgg  e la metto nel formato gg/mm/aaaa
        Dim Giorno As String
        Dim Mese As String
        Dim Anno As String
        GiraDataFromDB = ""
        data = CToStr(data, True, False, False)
        If Not IsDBNull(data) And Not IsNothing(data) Then
            If data <> "" Then
                Giorno = Mid(data, 7, 2)
                Mese = Mid(data, 5, 2)
                Anno = Mid(data, 1, 4)
                GiraDataFromDB = Giorno & "/" & Mese & "/" & Anno
            End If

            If IsDate(GiraDataFromDB) = False Then
                Giorno = Mid(data, 7, 2)
                Mese = Mid(data, 5, 2)
                Anno = Mid(data, 1, 4)
                GiraDataFromDB = Mese & "/" & Giorno & "/" & Anno
            End If
        End If
        Return GiraDataFromDB
    End Function

    Public Function giorni_mese(ByVal mese As Integer)
        Select Case mese
            Case 1, 3, 5, 7, 8, 12, 10
                Return 31
            Case 2
                Return 29
            Case 4, 6, 9, 11
                Return 30
            Case Else
                Return 30
        End Select
    End Function
    ''' <summary>
    ''' Funzione che calcola i mesi di possesso IMU.
    ''' Se nel mese ho meno di 15GG il mese non deve essere conteggiato
    ''' Se dalla data di inizio a fine mese ho meno di 15gg sposto il mese di inizio al mese successivo.
    ''' Se dalla data di fine a fine mese ho meno di 15gg sposto il mese di fine al mese precedente.
    ''' Calcolo i mesi utilizzando le nuove date determinate.
    ''' Se ho zero mesi ma la data di inizio e la data di fine sono nello stesso mese ed ho almeno 15gg allora forzo ad un mese.
    ''' </summary>
    ''' <param name="dInizio"></param>
    ''' <param name="dFine"></param>
    ''' <param name="nAnnoAccertamento"></param>
    ''' <param name="nMonthStartInYear"></param>
    ''' <param name="nMonthEndInYear"></param>
    ''' <revisionHistory><revision date="21/04/2021">I 15gg minimi devono essere all'interno dello stesso mese se quindi ricadono su mesi diversi non deve essere conteggiato</revision></revisionHistory>
    Public Function mesi_possesso(ByVal dInizio As DateTime, ByVal dFine As DateTime, ByVal nAnnoAccertamento As Integer, ByRef nMonthStartInYear As Integer, ByRef nMonthEndInYear As Integer) As Integer
        Dim nMesi As Integer = 0
        Dim dMonthEnd As DateTime = Now
        Dim IsAllMonthStart As Integer
        Dim IsAllMonthEnd As Integer

        Try
            Log.Debug("Generale.mesi_possesso.inizio=" + dInizio.ToString + "  fine=" + dFine.ToString)
            'determino mese inizio
            If dInizio.Year < nAnnoAccertamento Then
                Log.Debug("Generale.mesi_possesso.anno di inizio < anno calcolo mese partenza=1")
                nMonthStartInYear = 1
            Else
                dMonthEnd = New DateTime(dInizio.Year, dInizio.Month, 1).AddMonths(1).AddDays(-1)
                Log.Debug("Generale.mesi_possesso.ultimo giorno mese inizio=" + dMonthEnd.ToString)
                If DateDiff(DateInterval.Day, dInizio, dMonthEnd) > 14 Then
                    IsAllMonthStart = 0
                    Log.Debug("Generale.mesi_possesso.no aggiungo mese perchè ho più di 15gg nel mese=" + dInizio.ToString + "-" + dMonthEnd.ToString + "=" + DateDiff(DateInterval.Day, dInizio, dMonthEnd).ToString)
                Else
                    IsAllMonthStart = 1
                    Log.Debug("Generale.mesi_possesso.aggiungo mese perchè ho meno di 15gg nel mese=" + dInizio.ToString + "-" + dMonthEnd.ToString + "=" + DateDiff(DateInterval.Day, dInizio, dMonthEnd).ToString)
                End If
                nMonthStartInYear = dInizio.Month + IsAllMonthStart
                Log.Debug("Generale.mesi_possesso.anno di inizio > anno calcolo mese partenza=" + nMonthStartInYear.ToString)
            End If
            'determino mese fine
            dMonthEnd = New DateTime(dFine.Year, dFine.Month, 1)
            Log.Debug("Generale.mesi_possesso.primo giorno mese fine=" + dMonthEnd.ToString)
            If dFine.Year > nAnnoAccertamento Then
                nMonthEndInYear = 12
                dFine = "31/12/" + nAnnoAccertamento.ToString
            Else
                If DateDiff(DateInterval.Day, dMonthEnd, dFine) > 14 Then
                    IsAllMonthEnd = 0
                    Log.Debug("Generale.mesi_possesso.no tolgo mese perchè ho almeno 15gg nel mese=" + dFine.ToString + "-" + dMonthEnd.ToString + "=" + DateDiff(DateInterval.Day, dMonthEnd, dFine).ToString)
                Else
                    IsAllMonthEnd = -1
                    Log.Debug("Generale.mesi_possesso.tolgo mese perchè ho meno di 15gg nel mese=" + dFine.ToString + "-" + dMonthEnd.ToString + "=" + DateDiff(DateInterval.Day, dMonthEnd, dFine).ToString)
                End If
                nMonthEndInYear = dFine.Month + IsAllMonthEnd
            End If
            Log.Debug("Generale.mesi_possesso.ultimo mese nell'anno " + nMonthEndInYear.ToString + " primo mese nell'anno " + nMonthStartInYear.ToString)
            'calcolo i mesi
            nMesi = nMonthEndInYear - (nMonthStartInYear - 1)
            If nMesi = 0 And dInizio.Month = dFine.Month And DateDiff(DateInterval.Day, dInizio, dFine) > 14 Then
                Log.Debug("Generale.mesi_possesso.inizio e fine in stesso anno+mese e per almeno 15GG")
                nMesi = 1
            End If
            Log.Debug("Generale.mesi_possesso.mesi calcolati=" + nMesi.ToString)
        Catch ex As Exception
            Log.Debug("Generale.mesi_possesso.errore:: ", ex)
            Throw New Exception("Generale.mesi_possesso.errore::" & ex.Message)
        End Try
        Return nMesi
    End Function
    'Public Sub mesi_possesso(ByRef mesipossesso As Integer, ByVal dal As String, ByVal al As String, ByVal tipo_periodo As Integer, ByVal annoAccertamento As Integer)
    '    '********************************************************
    '    'Input:
    '    '       tipo_periodo: 1 Possesso, 2 Catasto_s, 3 Catasto_p
    '    '       dal : Data partenza periodo
    '    '       al : data chiusura periodo
    '    '
    '    'Output:
    '    '       mesipossesso: mesi di possesso del periodo in input rispetto
    '    'all()		'anno in esame
    '    '       impostazioni variabili globali
    '    '
    '    '********************************************************

    '    Dim mese As Integer
    '    Dim aggiunta_mese As Integer
    '    Dim data_ultimo_gg_mese As String


    '    'Azzero le variabili globali
    '    Select Case tipo_periodo
    '        Case 1              'Periodo di possesso
    '            glbmese_inizio_p = 0
    '            glbmese_fine_p = 0
    '        Case 2              'Periodo di classamento
    '            glbmese_inizio_s = 0
    '            glbmese_fine_s = 0
    '    End Select

    '    If Year(dal) < annoAccertamento Then
    '        If al = "" Then
    '            mesipossesso = 12
    '        Else
    '            If Year(al) > annoAccertamento Then
    '                mesipossesso = 12
    '            Else
    '                If Year(al) < annoAccertamento Then
    '                    mesipossesso = 0
    '                Else
    '                    'Forzo la data Dal all'inizio anno
    '                    dal = "01/01" + "/" + Trim(annoAccertamento)
    '                    'Verifico quanti gg di possosseso ci sono nel mese della Data(al)
    '                    If Day(al) > 14 Then
    '                        aggiunta_mese = aggiunta_mese + 1
    '                    End If
    '                    'Calcolo i mesi di possesso per l'anno in esame
    '                    mesipossesso = DateDiff("M", dal, al) + aggiunta_mese
    '                End If
    '            End If
    '        End If
    '        Select Case tipo_periodo
    '            Case 1                  'Periodo di possesso
    '                glbmese_inizio_p = 1
    '                glbmese_fine_p = mesipossesso
    '            Case 2                  'Periodo di classamento
    '                glbmese_inizio_s = 1
    '                glbmese_fine_s = mesipossesso
    '        End Select
    '        'glbmese_inizio = 1
    '        'glbmese_fine = mesipossesso
    '    Else
    '        If Year(dal) = annoAccertamento Then
    '            'Determino quanti giorni di possesso ci sono nel mese
    '            'del dal
    '            mese = Month(dal)
    '            data_ultimo_gg_mese = Trim(Str(giorni_mese(mese))) + "/" + Trim(Str(mese)) + "/" + Trim(Str(annoAccertamento))
    '            'Imposto i mesi di inizio/fine del periodo
    '            Select Case tipo_periodo
    '                Case 1                      'Periodo di possesso
    '                    glbmese_inizio_p = mese
    '                    glbmese_fine_p = 12
    '                Case 2                      'Periodo di classamento
    '                    glbmese_inizio_s = mese
    '                    glbmese_fine_s = 12
    '            End Select
    '            'glbmese_inizio = mese
    '            'glbmese_fine = 12
    '            'Verfico se, in presenza di febbraio, ho 28 o 29 gg
    '            If mese = 2 Then
    '                'data_ultimo_gg_mese = DateValue(data_ultimo_gg_mese)
    '                If Date.IsLeapYear(annoAccertamento) = False Then
    '                    data_ultimo_gg_mese = "28" + "/" + Trim(Str(mese)) + "/" + Trim(Str(annoAccertamento))
    '                End If
    '                'data_ultimo_gg_mese = "28" + "/" + Trim(Str(mese)) + "/" + Trim(Str(annoAccertamento))
    '            End If
    '            'Verifico quanti gg di possosseso ci sono nel mese della data Dal
    '            If mese = 2 Then
    '                If DateDiff("d", dal, data_ultimo_gg_mese) < 14 Then                       'modifica 8.1
    '                    'aggiunto = perchè in caso di un rif. cat
    '                    aggiunta_mese = -1
    '                    'secondario dal 16/11/1993 al ........ già di un secondo
    '                    'Sposto di uno il mese di inizio
    '                    'proprietario mi calcolava un mese per il primo
    '                    Select Case tipo_periodo
    '                        Case 1                             'Periodo di possesso
    '                            glbmese_inizio_p = Month(dal) + 1
    '                        Case 2                             'Periodo di classamento
    '                            glbmese_inizio_s = Month(dal) + 1
    '                    End Select
    '                    'glbmese_inizio = Month(dal) + 1
    '                End If
    '            Else

    '                If DateDiff("d", dal, data_ultimo_gg_mese) <= 14 Then                      'modifica 8.1
    '                    'aggiunto = perchè in caso di un rif. cat
    '                    aggiunta_mese = -1
    '                    'secondario dal 16/11/1993 al ........ già di un secondo
    '                    'Sposto di uno il mese di inizio
    '                    'proprietario mi calcolava un mese per il primo
    '                    Select Case tipo_periodo
    '                        Case 1                             'Periodo di possesso
    '                            glbmese_inizio_p = Month(dal) + 1
    '                        Case 2                             'Periodo di classamento
    '                            glbmese_inizio_s = Month(dal) + 1
    '                    End Select
    '                    'glbmese_inizio = Month(dal) + 1
    '                End If
    '            End If
    '            'Verifico data Al
    '            If al = "" Then
    '                al = "31/12" + "/" + Trim(Str(annoAccertamento))
    '            Else
    '                If Year(al) > annoAccertamento Then
    '                    al = "31/12" + "/" + Trim(Str(annoAccertamento))
    '                End If
    '            End If
    '            'Verifico quanti gg di possosseso ci sono nel mese della data Al
    '            If Day(al) > 14 Then
    '                aggiunta_mese = aggiunta_mese + 1
    '                'Inposto il mese di fine nel caso non fosse Dicembre (12)
    '                Select Case tipo_periodo
    '                    Case 1                          'Periodo di possesso
    '                        glbmese_fine_p = Month(al)
    '                    Case 2                          'Periodo di classamento
    '                        glbmese_fine_s = Month(al)
    '                End Select
    '                'glbmese_fine = Month(al)
    '            Else
    '                'Sposto indietro di uno il mese di fine nel caso il mese in	esame
    '                'non avesse i giorni sufficenti da essere considerato
    '                Select Case tipo_periodo
    '                    Case 1                          'Periodo di possesso
    '                        glbmese_fine_p = Month(al) - 1
    '                    Case 2                          'Periodo di classamento
    '                        glbmese_fine_s = Month(al) - 1
    '                End Select
    '                'glbmese_fine = Month(al) - 1
    '            End If
    '            'Calcolo i mesi di possesso per l'anno in esame
    '            mesipossesso = DateDiff("M", dal, al) + aggiunta_mese
    '        Else
    '            mesipossesso = 0
    '        End If
    '    End If
    'End Sub
End Class

Public Class ClsDBManager
    Private Shared ReadOnly Log As ILog = LogManager.GetLogger(GetType(ClsDBManager))
    Public myUtility As New Generale
    ''' <summary>
    ''' Restituisce il primo progressivo disponibile in base ai parametri in ingresso
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="strNomeTabella"></param>
    ''' <returns></returns>
    Public Function getNewID(myStringConnection As String, ByVal strNomeTabella As String) As Long
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GetNewID", "NOMETABELLA", "MAXID")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("NOMETABELLA", strNomeTabella) _
                            , ctx.GetParam("MAXID", 0)
                        )
                Catch ex As Exception
                    Log.Debug("ClsDBManager.getNewID.errore: ", ex)
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("maxid"))
                Next
            End Using
        Catch ex As Exception
            Log.Debug("ClsDBManager.getNewID.errore: ", ex)
            Throw New Exception("getNewIDdbICI.si è verificato il seguente errore." & ex.Message)
        End Try
        Return nMyReturn
    End Function
    ''' <summary>
    ''' Svuoto la tabella con il calcolo per singolo immobile per il contribuente
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="ANNO"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="ENTE"></param>
    ''' <param name="CONTRIB"></param>
    ''' <returns></returns>
    Public Function Delete_SITUAZIONE_FINALE_ICI(myStringConnection As String, ByVal ANNO As String, ByVal Tributo As String, ByVal ENTE As String, ByVal CONTRIB As Long) As Integer
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_SITUAZIONE_FINALE_ICI_D", "CodEnte", "Anno", "Tributo", "CodContribuente")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("CodEnte", ENTE) _
                            , ctx.GetParam("Anno", ANNO) _
                            , ctx.GetParam("Tributo", Tributo) _
                            , ctx.GetParam("CodContribuente", CONTRIB)
                        )
                Catch ex As Exception
                    Log.Debug(ENTE & " - ClsDBManager.Delete_SITUAZIONE_FINALE_ICI_dbICI.errore: ", ex)
                    nMyReturn = -1
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("id"))
                Next
            End Using
        Catch ex As Exception
            Log.Debug(ENTE & " - ClsDBManager.Delete_SITUAZIONE_FINALE_ICI_dbICI.errore: ", ex)
            nMyReturn = -1
        End Try
        Return nMyReturn
    End Function
    ''' <summary>
    ''' Svuoto la tabella con il riepilogo per contribuente+anno+tributo+tipotasi
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="ANNO"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="ENTE"></param>
    ''' <param name="CONTRIB"></param>
    ''' <returns></returns>
    Public Function Delete_TP_CALCOLO_FINALE_ICI(myStringConnection As String, ByVal ANNO As String, ByVal Tributo As String, ByVal ENTE As String, ByVal CONTRIB As Long) As Long
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_CALCOLO_FINALE_ICI_D", "CodEnte", "Anno", "Tributo", "CodContribuente")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("CodEnte", ENTE) _
                            , ctx.GetParam("Anno", ANNO) _
                            , ctx.GetParam("Tributo", Tributo) _
                            , ctx.GetParam("CodContribuente", CONTRIB)
                        )
                Catch ex As Exception
                    Log.Debug(ENTE & " - ClsDBManager.Delete_TP_CALCOLO_FINALE_ICI.errore: ", ex)
                    nMyReturn = -1
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("id"))
                Next
            End Using
        Catch ex As Exception
            Log.Debug(ENTE & " - ClsDBManager.Delete_TP_CALCOLO_FINALE_ICI.errore: ", ex)
            nMyReturn = -1
        End Try
        Return nMyReturn
    End Function
    ''' <summary>
    ''' Prelevo da database l'elenco dei contribuenti per i quali calcolare.
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="ANNO"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="ENTE"></param>
    ''' <param name="CONTRIB"></param>
    ''' <returns></returns>
    Public Function getListaContribuentiFreezer(myStringConnection As String, ByVal ANNO As String, ByVal Tributo As String, ByVal ENTE As String, ByVal CONTRIB As Integer) As ListAnagrafica
        Dim myAdapter As New SqlClient.SqlDataAdapter
        Dim myDataSet As New DataSet
        Dim myListAnagrafica As New ListAnagrafica
        Dim sSQL As String = ""

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_getListaContribuentiFreezer", "IDENTE", "IdContribuente", "Anno", "Tributo")
                    myDataSet = ctx.GetDataSet(sSQL, "TBL", ctx.GetParam("IDENTE", ENTE) _
                            , ctx.GetParam("IdContribuente", CONTRIB) _
                            , ctx.GetParam("Anno", ANNO) _
                            , ctx.GetParam("Tributo", Tributo)
                        )
                Catch ex As Exception
                    Log.Debug(ENTE & " - ClsDBManager.getListaContribuentiFreezer.errore: ", ex)
                    Throw New Exception("ClsDBManager.getListaContribuentiFreezer.errore: " & ex.Message)
                Finally
                    ctx.Dispose()
                End Try
                myListAnagrafica.p_dsItemsANAGRAFICA = myDataSet
            End Using
        Catch ex As Exception
            Log.Debug(ENTE & " - ClsDBManager.getListaContribuentiFreezer.errore: ", ex)
            Throw New Exception("ClsDBManager.getListaContribuentiFreezer.errore: " & ex.Message)
        End Try
        Return myListAnagrafica
    End Function
    ''' <summary>
    ''' Funzione per il salvataggio degli importi calcolati per singolo immobile
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="objICI"></param>
    ''' <param name="nIDElaborazione"></param>
    ''' <param name="Operatore"></param>
    ''' <returns></returns>
    ''' <revisionHistory>
    ''' <revision date="30/04/2015">
    ''' <strong>TASI Inquilino</strong>
    ''' </revision>
    ''' </revisionHistory>
    ''' <revisionHistory>
    ''' <revision date="12/04/2019">
    ''' <strong>Qualificazione AgID-analisi_rel01</strong>
    ''' <em>Analisi eventi</em>
    ''' </revision>
    ''' </revisionHistory>
    Public Function Set_SITUAZIONE_FINALE(myStringConnection As String, IdEnte As String, ByVal objICI As objSituazioneFinale(), ByVal nIDElaborazione As Long, Operatore As String) As Integer
        Dim lngID As Long
        Dim intAP As Integer
        Dim strTipoRendita As String
        Dim NumeroFabbricati As Integer
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            If nIDElaborazione = -1 Then 'calcolo puntuale
                'prelevo l'id elaborazione (eventuale) dei record del contribuente che sto per eliminare
                nIDElaborazione = getDati_TP_SITUAZIONE_FINALE_ICI_CALCOLO_MASSIVO(myStringConnection, IdEnte, objICI(0).Anno, objICI(0).IdContribuente)
            End If
            For Each mySituazioneFinale As objSituazioneFinale In objICI
                lngID = getNewID(myStringConnection, "TP_SITUAZIONE_FINALE_ICI")
                intAP = mySituazioneFinale.FlagPrincipale
                strTipoRendita = mySituazioneFinale.TipoRendita
                If (intAP = 1) Or ((strTipoRendita.ToUpper() <> "AF") And (strTipoRendita.ToUpper() <> "TA") And (intAP <> 1)) Then
                    NumeroFabbricati = 1
                Else
                    NumeroFabbricati = 0
                End If
                Using ctx As New DBModel(Generale.DBType, myStringConnection)
                    Try
                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_SITUAZIONE_FINALE_ICI_IU", "ID_SITUAZIONE_FINALE" _
                            , "ANNO" _
                            , "COD_ENTE" _
                            , "PROVENIENZA" _
                            , "CARATTERISTICA" _
                            , "INDIRIZZO" _
                            , "SEZIONE" _
                            , "FOGLIO" _
                            , "NUMERO" _
                            , "SUBALTERNO" _
                            , "CATEGORIA" _
                            , "CLASSE" _
                            , "PROTOCOLLO" _
                            , "FLAG_STORICO" _
                            , "VALORE" _
                            , "VALORE_REALE" _
                            , "FLAG_PROVVISORIO" _
                            , "PERC_POSSESSO" _
                            , "MESI_POSSESSO" _
                            , "MESI_ESCL_ESENZIONE" _
                            , "MESI_RIDUZIONE" _
                            , "IMPORTO_DETRAZIONE" _
                            , "FLAG_POSSEDUTO" _
                            , "FLAG_ESENTE" _
                            , "FLAG_RIDUZIONE" _
                            , "FLAG_PRINCIPALE" _
                            , "COD_CONTRIBUENTE" _
                            , "COD_IMMOBILE_PERTINENZA" _
                            , "COD_IMMOBILE" _
                            , "DAL" _
                            , "AL" _
                            , "NUMERO_MESI_ACCONTO" _
                            , "NUMERO_MESI_TOTALI" _
                            , "NUMERO_UTILIZZATORI" _
                            , "TIPO_RENDITA" _
                            , "ICI_ACCONTO_SENZA_DETRAZIONE" _
                            , "ICI_ACCONTO_DETRAZIONE_APPLICATA" _
                            , "ICI_DOVUTA_ACCONTO" _
                            , "ICI_ACCONTO_DETRAZIONE_RESIDUA" _
                            , "ICI_TOTALE_SENZA_DETRAZIONE" _
                            , "ICI_TOTALE_DETRAZIONE_APPLICATA" _
                            , "ICI_TOTALE_DOVUTA" _
                            , "ICI_TOTALE_DETRAZIONE_RESIDUA" _
                            , "ICI_DOVUTA_SALDO" _
                            , "ICI_DOVUTA_DETRAZIONE_SALDO" _
                            , "ICI_DOVUTA_SENZA_DETRAZIONE" _
                            , "ICI_DOVUTA_DETRAZIONE_RESIDUA" _
                            , "RIDUZIONE" _
                            , "MESE_INIZIO" _
                            , "DATA_SCADENZA" _
                            , "TIPO_OPERAZIONE" _
                            , "RITORNATA" _
                            , "DATA_ELABORAZIONE" _
                            , "idtestata" _
                            , "PROGRESSIVO_ELABORAZIONE" _
                            , "NUMERO_FABBRICATI" _
                            , "ICI_ACCONTO_DETRAZIONE_STATALE_APPLICATA" _
                            , "ICI_ACCONTO_DETRAZIONE_STATALE_CALCOLATA" _
                            , "ICI_ACCONTO_DETRAZIONE_STATALE_RESIDUA" _
                            , "ICI_SALDO_DETRAZIONE_STATALE_APPLICATA" _
                            , "ICI_SALDO_DETRAZIONE_STATALE_CALCOLATA" _
                            , "ICI_SALDO_DETRAZIONE_STATALE_RESIDUA" _
                            , "ICI_TOTALE_DETRAZIONE_STATALE_APPLICATA" _
                            , "ICI_TOTALE_DETRAZIONE_STATALE_CALCOLATA" _
                            , "ICI_TOTALE_DETRAZIONE_STATALE_RESIDUA" _
                            , "consistenza" _
                            , "AbitazionePrincipaleAttuale" _
                            , "COLTIVATOREDIRETTO" _
                            , "NUMEROFIGLI" _
                            , "ICI_VALORE_ALIQUOTA_STATALE" _
                            , "ICI_DOVUTA_ACCONTO_STATALE" _
                            , "ICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE" _
                            , "ICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE" _
                            , "ICI_TOTALE_DOVUTA_STATALE" _
                            , "ICI_TOTALE_DETRAZIONE_APPLICATA_STATALE" _
                            , "ICI_TOTALE_DETRAZIONE_RESIDUA_STATALE" _
                            , "ICI_DOVUTA_SALDO_STATALE" _
                            , "ICI_DOVUTA_DETRAZIONE_SALDO_STATALE" _
                            , "ICI_DOVUTA_DETRAZIONE_RESIDUA_STATALE" _
                            , "ICI_VALORE_ALIQUOTA" _
                            , "PERCENTCARICOFIGLI" _
                            , "ID_ALIQUOTA" _
                            , "CODTRIBUTO" _
                            , "IDTIPOUTILIZZO" _
                            , "IDTIPOPOSSESSO" _
                            , "TIPOTASI" _
                            , "IDCONTRIBUENTECALCOLO" _
                            , "IDCONTRIBUENTEDICH" _
                            , "OPERATORE" _
                            , "DATA_INSERIMENTO"
                        )
                        myDataView = ctx.GetDataView(sSQL, "TABELLA", ctx.GetParam("ID_SITUAZIONE_FINALE", lngID) _
                            , ctx.GetParam("ANNO", mySituazioneFinale.Anno) _
                            , ctx.GetParam("COD_ENTE", mySituazioneFinale.IdEnte) _
                            , ctx.GetParam("PROVENIENZA", mySituazioneFinale.Provenienza) _
                            , ctx.GetParam("CARATTERISTICA", mySituazioneFinale.Caratteristica) _
                            , ctx.GetParam("INDIRIZZO", mySituazioneFinale.Via & " " & mySituazioneFinale.NCivico) _
                            , ctx.GetParam("SEZIONE", mySituazioneFinale.Sezione) _
                            , ctx.GetParam("FOGLIO", mySituazioneFinale.Foglio) _
                            , ctx.GetParam("NUMERO", mySituazioneFinale.Numero) _
                            , ctx.GetParam("SUBALTERNO", mySituazioneFinale.Subalterno) _
                            , ctx.GetParam("CATEGORIA", mySituazioneFinale.Categoria) _
                            , ctx.GetParam("CLASSE", mySituazioneFinale.Classe) _
                            , ctx.GetParam("PROTOCOLLO", mySituazioneFinale.Protocollo) _
                            , ctx.GetParam("FLAG_STORICO", mySituazioneFinale.FlagStorico) _
                            , ctx.GetParam("VALORE", mySituazioneFinale.Valore) _
                            , ctx.GetParam("VALORE_REALE", mySituazioneFinale.ValoreReale) _
                            , ctx.GetParam("FLAG_PROVVISORIO", mySituazioneFinale.FlagProvvisorio) _
                            , ctx.GetParam("PERC_POSSESSO", mySituazioneFinale.PercPossesso) _
                            , ctx.GetParam("MESI_POSSESSO", mySituazioneFinale.MesiPossesso) _
                            , ctx.GetParam("MESI_ESCL_ESENZIONE", mySituazioneFinale.MesiEsenzione) _
                            , ctx.GetParam("MESI_RIDUZIONE", mySituazioneFinale.MesiRiduzione) _
                            , ctx.GetParam("IMPORTO_DETRAZIONE", mySituazioneFinale.ImpDetrazione) _
                            , ctx.GetParam("FLAG_POSSEDUTO", mySituazioneFinale.FlagPosseduto) _
                            , ctx.GetParam("FLAG_ESENTE", mySituazioneFinale.FlagEsente) _
                            , ctx.GetParam("FLAG_RIDUZIONE", mySituazioneFinale.FlagRiduzione) _
                            , ctx.GetParam("FLAG_PRINCIPALE", mySituazioneFinale.FlagPrincipale) _
                            , ctx.GetParam("COD_CONTRIBUENTE", mySituazioneFinale.IdContribuente) _
                            , ctx.GetParam("COD_IMMOBILE_PERTINENZA", mySituazioneFinale.IdImmobilePertinenza) _
                            , ctx.GetParam("COD_IMMOBILE", mySituazioneFinale.IdImmobile) _
                            , ctx.GetParam("DAL", mySituazioneFinale.Dal.ToString("yyyyMMdd")) _
                            , ctx.GetParam("AL", mySituazioneFinale.Al.ToString("yyyyMMdd")) _
                            , ctx.GetParam("NUMERO_MESI_ACCONTO", mySituazioneFinale.AccMesi) _
                            , ctx.GetParam("NUMERO_MESI_TOTALI", mySituazioneFinale.Mesi) _
                            , ctx.GetParam("NUMERO_UTILIZZATORI", mySituazioneFinale.NUtilizzatori) _
                            , ctx.GetParam("TIPO_RENDITA", mySituazioneFinale.TipoRendita) _
                            , ctx.GetParam("ICI_ACCONTO_SENZA_DETRAZIONE", mySituazioneFinale.AccSenzaDetrazione) _
                            , ctx.GetParam("ICI_ACCONTO_DETRAZIONE_APPLICATA", mySituazioneFinale.AccDetrazioneApplicata) _
                            , ctx.GetParam("ICI_DOVUTA_ACCONTO", mySituazioneFinale.AccDovuto) _
                            , ctx.GetParam("ICI_ACCONTO_DETRAZIONE_RESIDUA", mySituazioneFinale.AccDetrazioneResidua) _
                            , ctx.GetParam("ICI_TOTALE_SENZA_DETRAZIONE", mySituazioneFinale.TotSenzaDetrazione) _
                            , ctx.GetParam("ICI_TOTALE_DETRAZIONE_APPLICATA", mySituazioneFinale.TotDetrazioneApplicata) _
                            , ctx.GetParam("ICI_TOTALE_DOVUTA", mySituazioneFinale.TotDovuto) _
                            , ctx.GetParam("ICI_TOTALE_DETRAZIONE_RESIDUA", mySituazioneFinale.TotDetrazioneResidua) _
                            , ctx.GetParam("ICI_DOVUTA_SALDO", mySituazioneFinale.SalDovuto) _
                            , ctx.GetParam("ICI_DOVUTA_DETRAZIONE_SALDO", mySituazioneFinale.SalDetrazioneApplicata) _
                            , ctx.GetParam("ICI_DOVUTA_SENZA_DETRAZIONE", mySituazioneFinale.SalSenzaDetrazione) _
                            , ctx.GetParam("ICI_DOVUTA_DETRAZIONE_RESIDUA", mySituazioneFinale.SalDetrazioneResidua) _
                            , ctx.GetParam("RIDUZIONE", mySituazioneFinale.FlagRiduzione) _
                            , ctx.GetParam("MESE_INIZIO", mySituazioneFinale.MeseInizio) _
                            , ctx.GetParam("DATA_SCADENZA", mySituazioneFinale.DataScadenza) _
                            , ctx.GetParam("TIPO_OPERAZIONE", mySituazioneFinale.TipoOperazione) _
                            , ctx.GetParam("RITORNATA", myUtility.CToBit(False)) _
                            , ctx.GetParam("DATA_ELABORAZIONE", Date.Now.ToString("yyyyMMdd")) _
                            , ctx.GetParam("idtestata", DBNull.Value) _
                            , ctx.GetParam("PROGRESSIVO_ELABORAZIONE", nIDElaborazione) _
                            , ctx.GetParam("NUMERO_FABBRICATI", NumeroFabbricati) _
                            , ctx.GetParam("ICI_ACCONTO_DETRAZIONE_STATALE_APPLICATA", mySituazioneFinale.AccDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_ACCONTO_DETRAZIONE_STATALE_CALCOLATA", mySituazioneFinale.AccDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_ACCONTO_DETRAZIONE_STATALE_RESIDUA", mySituazioneFinale.AccDetrazioneResiduaStatale) _
                            , ctx.GetParam("ICI_SALDO_DETRAZIONE_STATALE_APPLICATA", mySituazioneFinale.SalDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_SALDO_DETRAZIONE_STATALE_CALCOLATA", mySituazioneFinale.SalDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_SALDO_DETRAZIONE_STATALE_RESIDUA", mySituazioneFinale.SalDetrazioneResiduaStatale) _
                            , ctx.GetParam("ICI_TOTALE_DETRAZIONE_STATALE_APPLICATA", mySituazioneFinale.TotDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_TOTALE_DETRAZIONE_STATALE_CALCOLATA", mySituazioneFinale.TotDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_TOTALE_DETRAZIONE_STATALE_RESIDUA", mySituazioneFinale.TotDetrazioneResiduaStatale) _
                            , ctx.GetParam("consistenza", mySituazioneFinale.Consistenza) _
                            , ctx.GetParam("AbitazionePrincipaleAttuale", mySituazioneFinale.AbitazionePrincipaleAttuale) _
                            , ctx.GetParam("COLTIVATOREDIRETTO", mySituazioneFinale.IsColtivatoreDiretto) _
                            , ctx.GetParam("NUMEROFIGLI", mySituazioneFinale.NumeroFigli) _
                            , ctx.GetParam("ICI_VALORE_ALIQUOTA_STATALE", mySituazioneFinale.AliquotaStatale) _
                            , ctx.GetParam("ICI_DOVUTA_ACCONTO_STATALE", mySituazioneFinale.AccDovutoStatale) _
                            , ctx.GetParam("ICI_ACCONTO_DETRAZIONE_APPLICATA_STATALE", mySituazioneFinale.AccDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_ACCONTO_DETRAZIONE_RESIDUA_STATALE", mySituazioneFinale.AccDetrazioneResiduaStatale) _
                            , ctx.GetParam("ICI_TOTALE_DOVUTA_STATALE", mySituazioneFinale.TotDovutoStatale) _
                            , ctx.GetParam("ICI_TOTALE_DETRAZIONE_APPLICATA_STATALE", mySituazioneFinale.TotDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_TOTALE_DETRAZIONE_RESIDUA_STATALE", mySituazioneFinale.TotDetrazioneResiduaStatale) _
                            , ctx.GetParam("ICI_DOVUTA_SALDO_STATALE", mySituazioneFinale.SalDovutoStatale) _
                            , ctx.GetParam("ICI_DOVUTA_DETRAZIONE_SALDO_STATALE", mySituazioneFinale.SalDetrazioneApplicataStatale) _
                            , ctx.GetParam("ICI_DOVUTA_DETRAZIONE_RESIDUA_STATALE", mySituazioneFinale.SalDetrazioneResiduaStatale) _
                            , ctx.GetParam("ICI_VALORE_ALIQUOTA", mySituazioneFinale.Aliquota) _
                            , ctx.GetParam("PERCENTCARICOFIGLI", mySituazioneFinale.PercentCaricoFigli) _
                            , ctx.GetParam("ID_ALIQUOTA", mySituazioneFinale.IdAliquota) _
                            , ctx.GetParam("CODTRIBUTO", mySituazioneFinale.Tributo) _
                            , ctx.GetParam("IDTIPOUTILIZZO", mySituazioneFinale.IdTipoUtilizzo) _
                            , ctx.GetParam("IDTIPOPOSSESSO", mySituazioneFinale.IdTipoPossesso) _
                            , ctx.GetParam("TIPOTASI", mySituazioneFinale.TipoTasi) _
                            , ctx.GetParam("IDCONTRIBUENTECALCOLO", mySituazioneFinale.IdContribuenteCalcolo) _
                            , ctx.GetParam("IDCONTRIBUENTEDICH", mySituazioneFinale.IdContribuenteDich) _
                            , ctx.GetParam("OPERATORE", Operatore) _
                            , ctx.GetParam("DATA_INSERIMENTO", DateTime.Now)
                        )
                    Catch ex As Exception
                        Log.Debug(IdEnte & " - ClsDBManager.Set_SITUAZIONE_FINALE.errore: ", ex)
                        nMyReturn = -1
                    Finally
                        ctx.Dispose()
                    End Try
                    For Each myRow As DataRowView In myDataView
                        nMyReturn = StringOperation.FormatInt(myRow("id"))
                    Next
                End Using
                If nMyReturn = Generale.INIT_VALUE_NUMBER Then
                    Log.Error(IdEnte & " - ClsDBManager.Set_SITUAZIONE_FINALE.errore:Inserimento in TP_SITUAZIONE_FINALE_ICI fallito")
                    Throw New Exception(IdEnte & " - ClsDBManager.Set_SITUAZIONE_FINALE.errore:Inserimento in TP_SITUAZIONE_FINALE_ICI fallito")
                End If
            Next
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.Set_SITUAZIONE_FINALE.errore: ", ex)
            Throw New Exception("ClsDBManager.Set_SITUAZIONE_FINALE.errore: " & ex.Message)
        End Try
        Return nMyReturn
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="ANNO"></param>
    ''' <param name="COD_CONTRIB"></param>
    ''' <returns></returns>
    Public Function getDati_TP_SITUAZIONE_FINALE_ICI_CALCOLO_MASSIVO(myStringConnection As String, IdEnte As String, ByVal ANNO As String, ByVal COD_CONTRIB As Long) As Integer
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GetPROGRESSIVO_ELABORAZIONE", "idEnte", "Anno", "idContribuente")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("idEnte", IdEnte) _
                            , ctx.GetParam("Anno", ANNO) _
                            , ctx.GetParam("idContribuente", COD_CONTRIB)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.getDati_TP_SITUAZIONE_FINALE_ICI_CALCOLO_MASSIVO.errore: ", ex)
                    nMyReturn = -1
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("PROGRESSIVO_ELABORAZIONE"))
                Next
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.getDati_TP_SITUAZIONE_FINALE_ICI_CALCOLO_MASSIVO.errore: ", ex)
            nMyReturn = -1
        End Try
        Return nMyReturn
    End Function

    Public Function Set_TP_CALCOLO_FINALE_ICI(myStringConnection As String, ByVal objDSfinale As objSituazioneFinale(), ByVal lngIDelaborazione As Long, ByVal blnCalcolaArrotondamento As Boolean) As Integer
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1
        Dim lngID As Long
        Dim TributoPrec As String = ""
        Dim ContribPrec As Integer = -1

        Try
            For Each mySituazioneFinale As objSituazioneFinale In objDSfinale
                '*** 20150430 - TASI Inquilino ***
                If TributoPrec <> mySituazioneFinale.Tributo Or ContribPrec <> mySituazioneFinale.IdContribuenteCalcolo Then
                    lngID = getNewID(myStringConnection, "TP_CALCOLO_FINALE_ICI")
                    Try
                        'Valorizzo la connessione
                        Using ctx As New DBModel(Generale.DBType, myStringConnection)
                            Try
                                sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_CALCOLO_FINALE_ICI_IU", "ID", "IdElaborazione", "IdEnte", "Anno", "Tributo", "IdContribuente")
                                myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("ID", lngID) _
                                    , ctx.GetParam("IdElaborazione", lngIDelaborazione) _
                                    , ctx.GetParam("IdEnte", mySituazioneFinale.IdEnte) _
                                    , ctx.GetParam("Anno", mySituazioneFinale.Anno) _
                                    , ctx.GetParam("Tributo", mySituazioneFinale.Tributo) _
                                    , ctx.GetParam("IdContribuente", mySituazioneFinale.IdContribuenteCalcolo)
                                )
                            Catch ex As Exception
                                Log.Debug(mySituazioneFinale.IdEnte & " - ClsDBManager.Set_TP_CALCOLO_FINALE_ICI.errore: ", ex)
                                nMyReturn = -1
                            Finally
                                ctx.Dispose()
                            End Try
                            For Each myRow As DataRowView In myDataView
                                nMyReturn = StringOperation.FormatInt(myRow("id"))
                            Next
                        End Using
                    Catch ex As Exception
                        Log.Debug(mySituazioneFinale.IdEnte & " - ClsDBManager.Set_TP_CALCOLO_FINALE_ICI.errore: ", ex)
                        nMyReturn = -1
                    End Try
                    If nMyReturn = Generale.INIT_VALUE_NUMBER Then
                        Log.Error(mySituazioneFinale.IdEnte & " - ClsDBManager.Set_TP_CALCOLO_FINALE_ICI.errore:Inserimento in TP_CALCOLO_FINALE_ICI fallito")
                        Throw New Exception(mySituazioneFinale.IdEnte & " - ClsDBManager.Set_TP_CALCOLO_FINALE_ICI.errore:Inserimento in TP_CALCOLO_FINALE_ICI fallito")
                    End If
                End If
                TributoPrec = mySituazioneFinale.Tributo
                ContribPrec = mySituazioneFinale.IdContribuenteCalcolo
            Next
        Catch ex As Exception
            Log.Debug("ClsDBManager.Set_SITUAZIONE_FINALE.errore: ", ex)
            nMyReturn = -1
        End Try
        Return nMyReturn
    End Function

    Public Function GetImportoVersatoPerCalcoloICI(myStringConnection As String, ByVal objDSfinale As objSituazioneFinale, ByVal IdEnte As String) As Double
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GetImportoVersatoPerCalcoloICI", "idEnte", "Anno", "idContribuente")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("idEnte", IdEnte) _
                            , ctx.GetParam("Anno", objDSfinale.Anno) _
                            , ctx.GetParam("idContribuente", objDSfinale.IdContribuente)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.GetImportoVersatoPerCalcoloICI.errore: ", ex)
                    nMyReturn = -1
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("IMP_AREE_FABB"))
                Next
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.GetImportoVersatoPerCalcoloICI.errore: ", ex)
            nMyReturn = -1
        End Try
        Return nMyReturn
    End Function

    Public Function Set_RibaltaVersatoNelDovuto(myStringConnection As String, ByVal objDSfinale As objSituazioneFinale, ByVal dblSumImpVersato As Double) As Integer
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_RibaltaVersatoNelDovuto", "ID", "IdEnte", "Anno", "ImpVersato", "IdContribuente")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("ID", 0) _
                            , ctx.GetParam("IdEnte", objDSfinale.IdEnte) _
                            , ctx.GetParam("Anno", objDSfinale.Anno) _
                            , ctx.GetParam("ImpVersato", Replace(dblSumImpVersato, ",", ".")) _
                            , ctx.GetParam("IdContribuente", objDSfinale.IdContribuente)
                        )
                Catch ex As Exception
                    Log.Debug(objDSfinale.IdEnte & " - ClsDBManager.Set_RibaltaVersatoNelDovuto.errore: ", ex)
                    nMyReturn = -1
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("id"))
                Next
            End Using
            If nMyReturn <= Generale.VALUE_NUMBER_ZERO Then
                Log.Error(objDSfinale.IdEnte & " - ClsDBManager.Set_RibaltaVersatoNelDovuto.errore:Inserimento fallito")
                Throw New Exception(objDSfinale.IdEnte & " - ClsDBManager.Set_RibaltaVersatoNelDovuto.errore:Inserimento fallito")
            End If
        Catch ex As Exception
            Log.Debug(objDSfinale.IdEnte & " - ClsDBManager.Set_RibaltaVersatoNelDovuto.errore: ", ex)
            Throw New Exception(objDSfinale.IdEnte & " - ClsDBManager.Set_RibaltaVersatoNelDovuto.errore:Inserimento fallito")
        End Try
        Return nMyReturn
    End Function

    Public Function Set_TP_TASK_REPOSITORY(myStringConnection As String, ByVal ID_TASK_REPOSITORY As Long, ByVal lngID_ELABORAZIONE As Long, ByVal TIPO_ELABORAZIONE As String, ByVal DESCRIZIONE As String, ByVal lngCOUNT_RECORD As Long, ByVal IdEnte As String, Anno As String, Username As String) As Integer
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_TASK_REPOSITORY_IU", "ID_TASK_REPOSITORY", "IdEnte", "Anno", "PROGRESSIVO", "TIPO_ELABORAZIONE", "DESCRIZIONE", "OPERATORE", "NUMERO_AGGIORNATI")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("ID_TASK_REPOSITORY", ID_TASK_REPOSITORY) _
                            , ctx.GetParam("IdEnte", IdEnte) _
                            , ctx.GetParam("Anno", Anno) _
                            , ctx.GetParam("PROGRESSIVO", lngID_ELABORAZIONE) _
                            , ctx.GetParam("TIPO_ELABORAZIONE", TIPO_ELABORAZIONE) _
                            , ctx.GetParam("DESCRIZIONE", myUtility.CToStr(DESCRIZIONE, True, False, False)) _
                            , ctx.GetParam("OPERATORE", Username) _
                            , ctx.GetParam("NUMERO_AGGIORNATI", lngCOUNT_RECORD)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.Set_TP_TASK_REPOSITORY.errore: ", ex)
                    nMyReturn = -1
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("id"))
                Next
            End Using
            If nMyReturn <= Generale.VALUE_NUMBER_ZERO Then
                Log.Error(IdEnte & " - ClsDBManager.Set_TP_TASK_REPOSITORY.errore:Inserimento fallito")
                Throw New Exception(IdEnte & " - ClsDBManager.Set_TP_TASK_REPOSITORY.errore:Inserimento fallito")
            End If
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.Set_TP_TASK_REPOSITORY.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.Set_TP_TASK_REPOSITORY.errore:Inserimento fallito")
        End Try
        Return nMyReturn
    End Function
    ''' <summary>
    ''' Prelevo da database l'elenco delle aliquote di detrazione configurate
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="Tributo"></param>
    ''' <returns></returns>
    Public Function getDetrazioni(myStringConnection As String, IdEnte As String, Tributo As String) As DataSet
        Dim sSQL As String = ""
        Dim myDataSet As New DataSet

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_ALIQUOTE_ICI_S", "ID_ALIQUOTA", "IdEnte", "Anno", "Tributo", "Tipo", "Default")
                    myDataSet = ctx.GetDataSet(sSQL, "TBL", ctx.GetParam("ID_ALIQUOTA", -1) _
                            , ctx.GetParam("IdEnte", IdEnte) _
                            , ctx.GetParam("Anno", "") _
                            , ctx.GetParam("Tributo", Tributo) _
                            , ctx.GetParam("Tipo", Generale.TipoAliquote_D) _
                            , ctx.GetParam("Default", 1)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.getDetrazioni.errore: ", ex)
                Finally
                    ctx.Dispose()
                End Try
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.getDetrazioni.errore: ", ex)
        End Try
        Return myDataSet
    End Function
    ''' <summary>
    ''' Svuoto la tabella di appoggio dei dati principali delle unità immobiliari per le quali calcolare
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="strCOD_CONTRIBUENTE"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="Tributo"></param>
    Public Sub DeleteFreezer(ByVal myStringConnection As String, ByVal strCOD_CONTRIBUENTE As Long, ByVal IdEnte As String, ByVal Tributo As String)
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_SITUAZIONE_VIRTUALE_DICHIARATO_D", "CodEnte", "Tributo", "CodContribuente")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("CodEnte", IdEnte) _
                            , ctx.GetParam("Tributo", Tributo) _
                            , ctx.GetParam("CodContribuente", strCOD_CONTRIBUENTE)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.DeleteFreezer.errore: ", ex)
                    nMyReturn = -1
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatInt(myRow("id"))
                Next
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.DeleteFreezer.errore: ", ex)
            nMyReturn = -1
        End Try
        If nMyReturn = Generale.INIT_VALUE_NUMBER Then
            Throw New Exception(IdEnte & " - ClsDBManager.DeleteFreezer.errore")
        End If
    End Sub

    Public Function getAnnoMinimoFreezer(myStringConnection As String, ByVal IdEnte As String, ByVal COD_CONTRIBUENTE As Long) As String
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim MyReturn As String = ""

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_getAnnoMinimoFreezer", "IDENTE", "IDCONTRIBUENTE")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("IDENTE", IdEnte) _
                            , ctx.GetParam("IDCONTRIBUENTE", COD_CONTRIBUENTE)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.getAnnoMinimoFreezer.errore: ", ex)
                    MyReturn = ""
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    MyReturn = StringOperation.FormatString(myRow("ANNO"))
                Next
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.getAnnoMinimoFreezer.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.getAnnoMinimoFreezer.errore: " & ex.Message)
        End Try
        Return MyReturn
    End Function
    '*** 20150430 - TASI Inquilino ***
    ''' <summary>
    ''' Prelevo da database l'elenco delle unità immobiliari da calcolare.
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="COD_CONTRIBUENTE"></param>
    ''' <param name="strAnno"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="TASIAProprietario"></param>
    ''' <param name="sTipoTASI"></param>
    ''' <returns></returns>
    Public Function getTutteDichiarazioni(myStringConnection As String, ByVal IdEnte As String, ByVal COD_CONTRIBUENTE As Long, ByVal strAnno As String, Tributo As String, TASIAProprietario As String, sTipoTASI As String) As DataSet
        Dim sSQL As String = ""
        Dim myDataSet As New DataSet

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GetTutteDichiarazioniFreezer", "IDENTE", "ANNO", "IDCONTRIBUENTE", "TASIAPROPRIETARIO", "TIPOTASI", "TRIBUTO")
                    myDataSet = ctx.GetDataSet(sSQL, "TBL", ctx.GetParam("IDENTE", IdEnte) _
                            , ctx.GetParam("ANNO", strAnno) _
                            , ctx.GetParam("IDCONTRIBUENTE", COD_CONTRIBUENTE) _
                            , ctx.GetParam("TASIAPROPRIETARIO", TASIAProprietario) _
                            , ctx.GetParam("TIPOTASI", sTipoTASI) _
                            , ctx.GetParam("TRIBUTO", Tributo)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.getTutteDichiarazioni.errore: ", ex)
                Finally
                    ctx.Dispose()
                End Try
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.getTutteDichiarazioni.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.getTutteDichiarazioni.errore: " & ex.Message)
        End Try
        Return myDataSet
    End Function
    '*** ***
    ''' <summary>
    ''' Salvo in una tabella di appoggio i dati principali delle unità immobiliari per le quali calcolare
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="objDSFreezerFINALE"></param>
    ''' <returns></returns>
    Public Function Set_TP_SITUAZIONE_VIRTUALE_DICHIARATO(myStringConnection As String, ByVal objDSFreezerFINALE As DataSet) As Integer
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Integer = -1

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                For Each myRow As DataRow In objDSFreezerFINALE.Tables(0).Rows
                    'se sono inquilino con tributo ici non inserisco
                    If Not (myRow.Item("TIPOTASI").ToString = Utility.Costanti.TIPOTASI_INQUILINO And myRow.Item("CODTRIBUTO").ToString = Utility.Costanti.TRIBUTO_ICI) Then
                        Try
                            sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_SITUAZIONE_VIRTUALE_DICHIARATO_IU", "Anno" _
                                , "Cod_Contribuente" _
                                , "ID_TESTATA" _
                                , "ID_IMMOBILE" _
                                , "Cod_Ente" _
                                , "CODTRIBUTO" _
                                , "COD_TIPO_PROCEDIMENTO" _
                                , "NUMERO_MESI_ACCONTO" _
                                , "NUMERO_MESI_TOTALI" _
                                , "NUMERO_UTILIZZATORI" _
                                , "FLAG_PRINCIPALE" _
                                , "PERC_POSSESSO" _
                                , "VALORE" _
                                , "RIDUZIONE" _
                                , "POSSESSO_FINE_ANNO" _
                                , "ESENTE_ESCLUSO" _
                                , "IDTIPOUTILIZZO" _
                                , "IDTIPOPOSSESSO" _
                                , "IMPORTO_DETRAZIONE" _
                                , "COD_IMMOBILE_PERTINENZA" _
                                , "COD_IMMOBILE_DA_ACCERTAMENTO" _
                                , "CONTITOLARE" _
                                , "TIPOTASI" _
                                , "IDCONTRIBUENTECALCOLO" _
                                , "IDCONTRIBUENTEDICH")
                            myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("Anno", StringOperation.FormatString(myRow.Item("ANNO"))) _
                                    , ctx.GetParam("Cod_Contribuente", StringOperation.FormatInt(myRow.Item("COD_CONTRIBUENTE"))) _
                                    , ctx.GetParam("ID_TESTATA", StringOperation.FormatString(myRow.Item("ID_TESTATA"))) _
                                    , ctx.GetParam("ID_IMMOBILE", StringOperation.FormatString(myRow.Item("ID_IMMOBILE"))) _
                                    , ctx.GetParam("Cod_Ente", StringOperation.FormatString(myRow.Item("COD_ENTE"))) _
                                    , ctx.GetParam("CODTRIBUTO", StringOperation.FormatString(myRow.Item("CODTRIBUTO"))) _
                                    , ctx.GetParam("COD_TIPO_PROCEDIMENTO", StringOperation.FormatString(myRow.Item("COD_TIPO_PROCEDIMENTO"))) _
                                    , ctx.GetParam("NUMERO_MESI_ACCONTO", StringOperation.FormatInt(myRow.Item("NUMERO_MESI_ACCONTO"))) _
                                    , ctx.GetParam("NUMERO_MESI_TOTALI", StringOperation.FormatInt(myRow.Item("NUMERO_MESI_TOTALI"))) _
                                    , ctx.GetParam("NUMERO_UTILIZZATORI", StringOperation.FormatInt(myRow.Item("NUMERO_UTILIZZATORI"))) _
                                    , ctx.GetParam("FLAG_PRINCIPALE", StringOperation.FormatInt(myRow.Item("FLAG_PRINCIPALE"))) _
                                    , ctx.GetParam("PERC_POSSESSO", StringOperation.FormatDouble(myRow.Item("PERC_POSSESSO"))) _
                                    , ctx.GetParam("VALORE", StringOperation.FormatDouble(myRow.Item("VALORE"))) _
                                    , ctx.GetParam("RIDUZIONE", StringOperation.FormatInt(myRow.Item("RIDUZIONE"))) _
                                    , ctx.GetParam("POSSESSO_FINE_ANNO", StringOperation.FormatInt(myRow.Item("POSSESSO_FINE_ANNO"))) _
                                    , ctx.GetParam("ESENTE_ESCLUSO", StringOperation.FormatInt(myRow.Item("ESENTE_ESCLUSO"))) _
                                    , ctx.GetParam("IDTIPOUTILIZZO", StringOperation.FormatInt(myRow.Item("IDTIPOUTILIZZO"))) _
                                    , ctx.GetParam("IDTIPOPOSSESSO", StringOperation.FormatInt(myRow.Item("IDTIPOPOSSESSO"))) _
                                    , ctx.GetParam("IMPORTO_DETRAZIONE", StringOperation.FormatDouble(myRow.Item("IMPORTO_DETRAZIONE"))) _
                                    , ctx.GetParam("COD_IMMOBILE_PERTINENZA", StringOperation.FormatString(myRow.Item("COD_IMMOBILE_PERTINENZA"))) _
                                    , ctx.GetParam("COD_IMMOBILE_DA_ACCERTAMENTO", StringOperation.FormatString(myRow.Item("COD_IMMOBILE_DA_ACCERTAMENTO"))) _
                                    , ctx.GetParam("CONTITOLARE", StringOperation.FormatInt(myRow.Item("CONTITOLARE"))) _
                                    , ctx.GetParam("TIPOTASI", StringOperation.FormatString(myRow.Item("TIPOTASI"))) _
                                    , ctx.GetParam("IDCONTRIBUENTECALCOLO", StringOperation.FormatInt(myRow.Item("IDCONTRIBUENTECALCOLO"))) _
                                    , ctx.GetParam("IDCONTRIBUENTEDICH", StringOperation.FormatInt(myRow.Item("IDCONTRIBUENTEDICH")))
                                )
                        Catch ex As Exception
                            Log.Debug(myRow.Item("COD_ENTE") & " - ClsDBManager.Set_TP_SITUAZIONE_VIRTUALE_DICHIARATO.errore: ", ex)
                            nMyReturn = -1
                        Finally
                            ctx.Dispose()
                        End Try
                        For Each myRowRet As DataRowView In myDataView
                            nMyReturn = StringOperation.FormatInt(myRowRet("id"))
                        Next
                        If nMyReturn = Generale.INIT_VALUE_NUMBER Then
                            Log.Debug("ClsDBManager.Set_TP_SITUAZIONE_VIRTUALE_DICHIARATO.errore: inserimento fallito")
                            Throw New Exception("ClsDBManager.Set_TP_SITUAZIONE_VIRTUALE_DICHIARATO.errore: inserimento fallito")
                        End If
                    End If
                Next
            End Using
        Catch ex As Exception
            Log.Debug("ClsDBManager.Set_TP_SITUAZIONE_VIRTUALE_DICHIARATO.errore: ", ex)
            nMyReturn = -1
        End Try
        Return nMyReturn
    End Function

    Public Function GetSituazioneVirtualeDichiarazioni(myStringConnection As String, ByVal objdsAnagrafica As DataSet, ByVal IdEnte As String, AnnoDa As String, AnnoA As String, ByVal strCOD_CONTRIBUENTE As String) As DataSet
        Dim sSQL As String = ""
        Dim myDataSet As New DataSet

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GETSITUAZIONEVIRTUALEDICHIARAZIONI", "IdEnte", "IdContribuente", "AnnoDa", "AnnoA")
                    myDataSet = ctx.GetDataSet(sSQL, "TBL", ctx.GetParam("IdEnte", IdEnte) _
                            , ctx.GetParam("IdContribuente", strCOD_CONTRIBUENTE) _
                            , ctx.GetParam("AnnoDa", AnnoDa) _
                            , ctx.GetParam("AnnoA", AnnoA)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.GetSituazioneVirtualeDichiarazioni.errore: ", ex)
                Finally
                    ctx.Dispose()
                End Try
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.GetSituazioneVirtualeDichiarazioni.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.GetSituazioneVirtualeDichiarazioni.errore: " & ex.Message)
        End Try
        Return myDataSet
    End Function
    ''' <summary>
    ''' Prelevo dalla tabella di appoggio i dati principali delle unità immobiliari per le quali calcolare.
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="AnnoDa"></param>
    ''' <param name="AnnoA"></param>
    ''' <param name="strCOD_CONTRIBUENTE"></param>
    ''' <returns></returns>
    Public Function GetSituazioneVirtualeImmobili(myStringConnection As String, ByVal IdEnte As String, AnnoDa As String, AnnoA As String, ByVal strCOD_CONTRIBUENTE As String) As DataSet
        Dim sSQL As String = ""
        Dim myDataSet As New DataSet

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GETSITUAZIONEVIRTUALEIMMOBILI", "IdEnte", "IdContribuente", "AnnoDa", "AnnoA")
                    myDataSet = ctx.GetDataSet(sSQL, "TBL", ctx.GetParam("IdEnte", IdEnte) _
                            , ctx.GetParam("IdContribuente", strCOD_CONTRIBUENTE) _
                            , ctx.GetParam("AnnoDa", AnnoDa) _
                            , ctx.GetParam("AnnoA", AnnoA)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.GetSituazioneVirtualeImmobili.errore: ", ex)
                Finally
                    ctx.Dispose()
                End Try
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.GetSituazioneVirtualeImmobili.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.GetSituazioneVirtualeImmobili.errore: " & ex.Message)
        End Try
        Return myDataSet
    End Function
    Public Function getCategorieDaEscludere(myStringConnection As String, ByVal IdEnte As String, Anno As String, ByVal sTipoAliquota As String, ByVal Tributo As String) As DataSet
        Dim sSQL As String = ""
        Dim myDataSet As New DataSet

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_getCategorieDaEscludere", "IdEnte", "Anno", "TipoAliquota", "Tributo")
                    myDataSet = ctx.GetDataSet(sSQL, "SITUAZIONE_VIRTUALE_IMMOBILI_ICI", ctx.GetParam("IdEnte", IdEnte) _
                            , ctx.GetParam("Anno", Anno) _
                            , ctx.GetParam("TipoAliquota", sTipoAliquota) _
                            , ctx.GetParam("Tributo", Tributo)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.getCategorieDaEscludere.errore: ", ex)
                Finally
                    ctx.Dispose()
                End Try
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.getCategorieDaEscludere.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.getCategorieDaEscludere.errore: " & ex.Message)
        End Try
        Return myDataSet
    End Function
    '*** 20150430 - TASI Inquilino ***
    Public Function getAliquote(myStringConnection As String, ByVal IdEnte As String, ByVal strAnno As String, ByVal strAliquota As String, ByVal Tributo As String, ByRef nValAliquotaStatale As Double, ByRef IdAliquota As Integer, ByRef nSogliaRendita As Double, ByRef sTipoSoglia As String, ByRef nPercInquilino As Double) As Double
        Dim sSQL As String = ""
        Dim myDataView As New DataView
        Dim nMyReturn As Double = 0

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TP_ALIQUOTE_ICI_S", "ID_ALIQUOTA", "IdEnte", "Anno", "Tributo", "Tipo", "Default")
                    myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("ID_ALIQUOTA", -1) _
                            , ctx.GetParam("IdEnte", IdEnte) _
                            , ctx.GetParam("Anno", strAnno) _
                            , ctx.GetParam("Tributo", Tributo) _
                            , ctx.GetParam("Tipo", strAliquota) _
                            , ctx.GetParam("Default", 1)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.getAliquote.errore: ", ex)
                    nMyReturn = 0
                Finally
                    ctx.Dispose()
                End Try
                For Each myRow As DataRowView In myDataView
                    nMyReturn = StringOperation.FormatDouble(myRow("valore"))
                    nValAliquotaStatale = StringOperation.FormatDouble(myRow("ALIQUOTA_STATALE"))
                    '*** 20130422 - aggiornamento IMU ***
                    IdAliquota = StringOperation.FormatDouble(myRow("id_aliquota"))
                    '*** ***
                    nSogliaRendita = StringOperation.FormatDouble(myRow("sogliarendita"))
                    sTipoSoglia = StringOperation.FormatString(myRow("tiposoglia"))
                    '*** 20150430 - TASI Inquilino ***
                    nPercInquilino = StringOperation.FormatDouble(myRow("PERCINQUILINO"))
                    '*** ***
                Next
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.getAliquote.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.getAliquote.errore: " & ex.Message)
        End Try
        Return nMyReturn
    End Function

    Public Function ViewCalcolo(myStringConnection As String, ByVal IdEnte As String, ByVal Anno As String, InCorso As Boolean) As DataSet
        Dim sSQL As String = ""
        Dim myDataSet As New DataSet

        Try
            'Valorizzo la connessione
            Using ctx As New DBModel(Generale.DBType, myStringConnection)
                Try
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GetCodaCalcolo", "IDENTE", "ANNO", "InCorso")
                    myDataSet = ctx.GetDataSet(sSQL, "TBL", ctx.GetParam("IDENTE", IdEnte) _
                            , ctx.GetParam("ANNO", Anno) _
                            , ctx.GetParam("InCorso", InCorso)
                        )
                Catch ex As Exception
                    Log.Debug(IdEnte & " - ClsDBManager.ViewCalcolo.errore: ", ex)
                Finally
                    ctx.Dispose()
                End Try
            End Using
        Catch ex As Exception
            Log.Debug(IdEnte & " - ClsDBManager.ViewCalcolo.errore: ", ex)
            Throw New Exception(IdEnte & " - ClsDBManager.ViewCalcolo.errore:  " & ex.Message)
        End Try
        Return myDataSet
    End Function
End Class
