Imports log4net
''' <summary>
''' Definizione oggetto calcolo
''' </summary>
''' <revisionHistory>
''' <revision date="20190910">
''' In caso di flag storico la riduzione deve essere applicata anche per l'abitazione principale
''' </revision>
''' </revisionHistory>
Public Class objCalcolo
    Private _dblValore As Double = 0
    Private _dblAliquota As Double = 0
    Private _dblPossesso As Double = 0
    Private _intMesi As Integer = 0
    Private _intMesiT As Integer = 0
    Private _dblAcconto As Double = 0
    Private _dblDetrazione As Double = 0
    Private _intUtilizzatori As Integer = 0
    Private _dblDetrazioneDichiarata As Double = 0
    Private _intAbitazionePrincipale As Integer = 0
    Private _intStorico As Integer = 0

    Private _dblIci_Teorica As Double = 0
    Private _dblIci_Dovuta As Double = 0
    Private _dblDetrazione_Applicata As Double = 0
    Private _dblDetrazione_Residua As Double = 0
    Private _blnRiduzione As Boolean = 0
    Private _intMese_Inizio As Integer = 0

    '*** 20120530 - IMU ***
    Private _AnnoCalcolo As Integer = 0
    Private _FigliACarico As Integer = 0
    Private _DetrazioneFigli As Double = 0
    Private _PercentCaricoFigli As Double = 0
    Private _AliquotaStatale As Double = 0
    Private _IciTeoricaStatale As Double = 0
    Private _IciDovutaStatale As Double = 0
    Private _DetrazioneResiduaStatale As Double = 0
    Private _DetrazioneApplicabileStatale As Double = 0
    '*** 20130422 - aggiornamento IMU ***
    Private _IdAliquota As Integer = 0
    '*** ***
    '*** 20150430 - TASI Inquilino ***
    Private _TipoTasi As String = Utility.Costanti.TIPOTASI_PROPRIETARIO
    Private _nPercInquilino As Double = 100
    '*** ***
    Private _TipoAliquota As String = ""
    Private _Detrazione_Residua_Standard As Double = 0

    Public Property AnnoCalcolo() As Integer
        Get
            Return _AnnoCalcolo
        End Get
        Set(ByVal Value As Integer)
            _AnnoCalcolo = Value
        End Set
    End Property
    Public Property FigliACarico() As Integer
        Get
            Return _FigliACarico
        End Get
        Set(ByVal Value As Integer)
            _FigliACarico = Value
        End Set
    End Property
    Public Property DetrazioneFigli() As Double
        Get
            Return _DetrazioneFigli
        End Get
        Set(ByVal Value As Double)
            _DetrazioneFigli = Value
        End Set
    End Property
    Public Property PercentCaricoFigli() As Double
        Get
            Return _PercentCaricoFigli
        End Get
        Set(ByVal Value As Double)
            _PercentCaricoFigli = Value
        End Set
    End Property
    Public Property AliquotaStatale() As Double
        Get
            Return _AliquotaStatale
        End Get
        Set(ByVal Value As Double)
            _AliquotaStatale = Value
        End Set
    End Property
    Public Property Ici_Teorica_Statale() As Double
        Get
            Return _IciTeoricaStatale
        End Get
        Set(ByVal Value As Double)
            _IciTeoricaStatale = Value
        End Set
    End Property
    Public Property Ici_Dovuta_Statale() As Double
        Get
            Return _IciDovutaStatale
        End Get
        Set(ByVal Value As Double)
            _IciDovutaStatale = Value
        End Set
    End Property
    Public Property Detrazione_Residua_Statale() As Double
        Get
            Return _DetrazioneResiduaStatale
        End Get
        Set(ByVal Value As Double)
            _DetrazioneResiduaStatale = Value
        End Set
    End Property
    Public Property Detrazione_Applicabile_Statale() As Double
        Get
            Return _DetrazioneApplicabileStatale
        End Get
        Set(ByVal Value As Double)
            _DetrazioneApplicabileStatale = Value
        End Set
    End Property
    Public Property Detrazione_Residua_Standard() As Double
        Get
            Return _Detrazione_Residua_Standard
        End Get
        Set(ByVal Value As Double)
            _Detrazione_Residua_Standard = Value
        End Set
    End Property
    '*** ***
    Public Property AbitazionePrincipale() As Integer
        Get
            AbitazionePrincipale = _intAbitazionePrincipale
        End Get
        Set(ByVal Value As Integer)
            _intAbitazionePrincipale = Value
        End Set
    End Property

    Public Property Mese_Inizio() As Integer
        Get
            Mese_Inizio = _intMese_Inizio
        End Get
        Set(ByVal Value As Integer)
            _intMese_Inizio = Value
        End Set
    End Property

    Public Property Riduzione() As Boolean
        Get
            Riduzione = _blnRiduzione
        End Get
        Set(ByVal Value As Boolean)
            _blnRiduzione = Value
        End Set
    End Property

    Public Property Detrazione_Residua() As Double
        Get
            Detrazione_Residua = _dblDetrazione_Residua
        End Get
        Set(ByVal Value As Double)
            _dblDetrazione_Residua = Value
        End Set
    End Property
    Public Property Detrazione_Applicabile() As Double
        Get
            Detrazione_Applicabile = _dblDetrazione_Applicata
        End Get
        Set(ByVal Value As Double)
            _dblDetrazione_Applicata = Value
        End Set
    End Property
    Public Property Ici_Dovuta() As Double
        Get
            Ici_Dovuta = _dblIci_Dovuta
        End Get
        Set(ByVal Value As Double)
            _dblIci_Dovuta = Value
        End Set
    End Property
    Public Property Ici_Teorica() As Double
        Get
            Ici_Teorica = _dblIci_Teorica
        End Get
        Set(ByVal Value As Double)
            _dblIci_Teorica = Value
        End Set
    End Property
    Public Property Utilizzatori() As Integer
        Get
            Utilizzatori = _intUtilizzatori
        End Get
        Set(ByVal Value As Integer)
            _intUtilizzatori = Value
        End Set
    End Property
    Public Property Detrazione() As Double
        Get
            Detrazione = _dblDetrazione
        End Get
        Set(ByVal Value As Double)
            _dblDetrazione = Value
        End Set
    End Property
    Public Property Acconto() As Double
        Get
            Acconto = _dblAcconto
        End Get
        Set(ByVal Value As Double)
            _dblAcconto = Value
        End Set
    End Property
    Public Property Mesi() As Integer
        Get
            Mesi = _intMesi
        End Get
        Set(ByVal Value As Integer)
            _intMesi = Value
        End Set
    End Property
    Public Property MesiT() As Integer
        Get
            MesiT = _intMesiT
        End Get
        Set(ByVal Value As Integer)
            _intMesiT = Value
        End Set
    End Property
    Public Property Possesso() As Double
        Get
            Possesso = _dblPossesso
        End Get
        Set(ByVal Value As Double)
            _dblPossesso = Value
        End Set
    End Property
    Public Property Aliquota() As Double
        Get
            Aliquota = _dblAliquota
        End Get
        Set(ByVal Value As Double)
            _dblAliquota = Value
        End Set
    End Property
    Public Property Valore() As Double
        Get
            Valore = _dblValore
        End Get
        Set(ByVal Value As Double)
            _dblValore = Value
        End Set
    End Property
    Public Property DetrazioneDichiarata() As Double
        Get
            DetrazioneDichiarata = _dblDetrazioneDichiarata
        End Get
        Set(ByVal Value As Double)
            _dblDetrazioneDichiarata = Value
        End Set
    End Property
    '*** 20130422 - aggiornamento IMU ***
    Public Property nIdAliquota() As Integer
        Get
            nIdAliquota = _IdAliquota
        End Get
        Set(ByVal Value As Integer)
            _IdAliquota = Value
        End Set
    End Property
    '*** ***
    '*** 20150430 - TASI Inquilino ***
    Public Property TipoTasi As String
        Get
            TipoTasi = _TipoTasi
        End Get
        Set(value As String)
            _TipoTasi = value
        End Set
    End Property
    Public Property nPercInquilino As Double
        Get
            nPercInquilino = _nPercInquilino
        End Get
        Set(value As Double)
            _nPercInquilino = value
        End Set
    End Property
    '*** ***
    Public Property TipoAliquota As String
        Get
            TipoAliquota = _TipoAliquota
        End Get
        Set(value As String)
            _TipoAliquota = value
        End Set
    End Property
    Public Property Storico() As Integer
        Get
            Storico = _intStorico
        End Get
        Set(ByVal Value As Integer)
            _intStorico = Value
        End Set
    End Property
End Class
'Public Class objCalcolo
'    Private _dblValore As Double = 0
'    Private _dblAliquota As Double = 0
'    Private _dblPossesso As Double = 0
'    Private _intMesi As Integer = 0
'    Private _intMesiT As Integer = 0
'    Private _dblAcconto As Double = 0
'    Private _dblDetrazione As Double = 0
'    Private _intUtilizzatori As Integer = 0
'    Private _dblDetrazioneDichiarata As Double = 0
'    Private _intAbitazionePrincipale As Integer = 0

'    Private _dblIci_Teorica As Double = 0
'    Private _dblIci_Dovuta As Double = 0
'    Private _dblDetrazione_Applicata As Double = 0
'    Private _dblDetrazione_Residua As Double = 0
'    Private _blnRiduzione As Boolean = 0
'    Private _intMese_Inizio As Integer = 0

'    '*** 20120530 - IMU ***
'    Private _AnnoCalcolo As Integer = 0
'    Private _FigliACarico As Integer = 0
'    Private _DetrazioneFigli As Double = 0
'    Private _PercentCaricoFigli As Double = 0
'    Private _AliquotaStatale As Double = 0
'    Private _IciTeoricaStatale As Double = 0
'    Private _IciDovutaStatale As Double = 0
'    Private _DetrazioneResiduaStatale As Double = 0
'    Private _DetrazioneApplicabileStatale As Double = 0
'    '*** 20130422 - aggiornamento IMU ***
'    Private _IdAliquota As Integer = 0
'    '*** ***
'    '*** 20150430 - TASI Inquilino ***
'    Private _TipoTasi As String = Utility.Costanti.TIPOTASI_PROPRIETARIO
'    Private _nPercInquilino As Double = 100
'    '*** ***
'    Private _TipoAliquota As String = ""
'    Private _Detrazione_Residua_Standard As Double = 0

'    Public Property AnnoCalcolo() As Integer
'        Get
'            Return _AnnoCalcolo
'        End Get
'        Set(ByVal Value As Integer)
'            _AnnoCalcolo = Value
'        End Set
'    End Property
'    Public Property FigliACarico() As Integer
'        Get
'            Return _FigliACarico
'        End Get
'        Set(ByVal Value As Integer)
'            _FigliACarico = Value
'        End Set
'    End Property
'    Public Property DetrazioneFigli() As Double
'        Get
'            Return _DetrazioneFigli
'        End Get
'        Set(ByVal Value As Double)
'            _DetrazioneFigli = Value
'        End Set
'    End Property
'    Public Property PercentCaricoFigli() As Double
'        Get
'            Return _PercentCaricoFigli
'        End Get
'        Set(ByVal Value As Double)
'            _PercentCaricoFigli = Value
'        End Set
'    End Property
'    Public Property AliquotaStatale() As Double
'        Get
'            Return _AliquotaStatale
'        End Get
'        Set(ByVal Value As Double)
'            _AliquotaStatale = Value
'        End Set
'    End Property
'    Public Property Ici_Teorica_Statale() As Double
'        Get
'            Return _IciTeoricaStatale
'        End Get
'        Set(ByVal Value As Double)
'            _IciTeoricaStatale = Value
'        End Set
'    End Property
'    Public Property Ici_Dovuta_Statale() As Double
'        Get
'            Return _IciDovutaStatale
'        End Get
'        Set(ByVal Value As Double)
'            _IciDovutaStatale = Value
'        End Set
'    End Property
'    Public Property Detrazione_Residua_Statale() As Double
'        Get
'            Return _DetrazioneResiduaStatale
'        End Get
'        Set(ByVal Value As Double)
'            _DetrazioneResiduaStatale = Value
'        End Set
'    End Property
'    Public Property Detrazione_Applicabile_Statale() As Double
'        Get
'            Return _DetrazioneApplicabileStatale
'        End Get
'        Set(ByVal Value As Double)
'            _DetrazioneApplicabileStatale = Value
'        End Set
'    End Property
'    Public Property Detrazione_Residua_Standard() As Double
'        Get
'            Return _Detrazione_Residua_Standard
'        End Get
'        Set(ByVal Value As Double)
'            _Detrazione_Residua_Standard = Value
'        End Set
'    End Property
'    '*** ***
'    Public Property AbitazionePrincipale() As Integer
'        Get
'            AbitazionePrincipale = _intAbitazionePrincipale
'        End Get
'        Set(ByVal Value As Integer)
'            _intAbitazionePrincipale = Value
'        End Set
'    End Property

'    Public Property Mese_Inizio() As Integer
'        Get
'            Mese_Inizio = _intMese_Inizio
'        End Get
'        Set(ByVal Value As Integer)
'            _intMese_Inizio = Value
'        End Set
'    End Property

'    Public Property Riduzione() As Boolean
'        Get
'            Riduzione = _blnRiduzione
'        End Get
'        Set(ByVal Value As Boolean)
'            _blnRiduzione = Value
'        End Set
'    End Property

'    Public Property Detrazione_Residua() As Double
'        Get
'            Detrazione_Residua = _dblDetrazione_Residua
'        End Get
'        Set(ByVal Value As Double)
'            _dblDetrazione_Residua = Value
'        End Set
'    End Property
'    Public Property Detrazione_Applicabile() As Double
'        Get
'            Detrazione_Applicabile = _dblDetrazione_Applicata
'        End Get
'        Set(ByVal Value As Double)
'            _dblDetrazione_Applicata = Value
'        End Set
'    End Property
'    Public Property Ici_Dovuta() As Double
'        Get
'            Ici_Dovuta = _dblIci_Dovuta
'        End Get
'        Set(ByVal Value As Double)
'            _dblIci_Dovuta = Value
'        End Set
'    End Property
'    Public Property Ici_Teorica() As Double
'        Get
'            Ici_Teorica = _dblIci_Teorica
'        End Get
'        Set(ByVal Value As Double)
'            _dblIci_Teorica = Value
'        End Set
'    End Property
'    Public Property Utilizzatori() As Integer
'        Get
'            Utilizzatori = _intUtilizzatori
'        End Get
'        Set(ByVal Value As Integer)
'            _intUtilizzatori = Value
'        End Set
'    End Property
'    Public Property Detrazione() As Double
'        Get
'            Detrazione = _dblDetrazione
'        End Get
'        Set(ByVal Value As Double)
'            _dblDetrazione = Value
'        End Set
'    End Property
'    Public Property Acconto() As Double
'        Get
'            Acconto = _dblAcconto
'        End Get
'        Set(ByVal Value As Double)
'            _dblAcconto = Value
'        End Set
'    End Property
'    Public Property Mesi() As Integer
'        Get
'            Mesi = _intMesi
'        End Get
'        Set(ByVal Value As Integer)
'            _intMesi = Value
'        End Set
'    End Property
'    Public Property MesiT() As Integer
'        Get
'            MesiT = _intMesiT
'        End Get
'        Set(ByVal Value As Integer)
'            _intMesiT = Value
'        End Set
'    End Property
'    Public Property Possesso() As Double
'        Get
'            Possesso = _dblPossesso
'        End Get
'        Set(ByVal Value As Double)
'            _dblPossesso = Value
'        End Set
'    End Property
'    Public Property Aliquota() As Double
'        Get
'            Aliquota = _dblAliquota
'        End Get
'        Set(ByVal Value As Double)
'            _dblAliquota = Value
'        End Set
'    End Property
'    Public Property Valore() As Double
'        Get
'            Valore = _dblValore
'        End Get
'        Set(ByVal Value As Double)
'            _dblValore = Value
'        End Set
'    End Property
'    Public Property DetrazioneDichiarata() As Double
'        Get
'            DetrazioneDichiarata = _dblDetrazioneDichiarata
'        End Get
'        Set(ByVal Value As Double)
'            _dblDetrazioneDichiarata = Value
'        End Set
'    End Property
'    '*** 20130422 - aggiornamento IMU ***
'    Public Property nIdAliquota() As Integer
'        Get
'            nIdAliquota = _IdAliquota
'        End Get
'        Set(ByVal Value As Integer)
'            _IdAliquota = Value
'        End Set
'    End Property
'    '*** ***
'    '*** 20150430 - TASI Inquilino ***
'    Public Property TipoTasi As String
'        Get
'            TipoTasi = _TipoTasi
'        End Get
'        Set(value As String)
'            _TipoTasi = value
'        End Set
'    End Property
'    Public Property nPercInquilino As Double
'        Get
'            nPercInquilino = _nPercInquilino
'        End Get
'        Set(value As Double)
'            _nPercInquilino = value
'        End Set
'    End Property
'    '*** ***
'    Public Property TipoAliquota As String
'        Get
'            TipoAliquota = _TipoAliquota
'        End Get
'        Set(value As String)
'            _TipoAliquota = value
'        End Set
'    End Property
'End Class
''' <summary>
''' Classe per il calcolo
''' </summary>
Public Class CALCOLO_ICI
    Private Shared Log As ILog = LogManager.GetLogger(GetType(CALCOLO_ICI))

    'Public Function getCALCOLO_ICI_ACCONTO_TOTALE_PREmodificheIMU() As Double

    '    Dim dblCalcoloIci As Double
    '    Dim dblCalcoloDetrazione As Double

    '    Const dblRIDUZIONE = 0.5

    '    _dblIci_Teorica = 0
    '    _dblDetrazione_Applicata = 0
    '    _dblIci_Dovuta = 0
    '    _dblDetrazione_Residua = 0


    '    '****************************************************************************************************
    '    'GESTIONE CONTROLLI VALORI PASSATI
    '    '****************************************************************************************************
    '    If _dblAliquota = -1 Then
    '        Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::ALIQUTA ERRATA")
    '    End If
    '    If _dblDetrazione = -1 Then
    '        Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::DETRAZIONE ERRATA")
    '    End If

    '    'If m_intUtilizzatori < 1 And m_intUtilizzatori > 1000 Then
    '    '    Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Il numero di UTILIZZATORI deve essere compreso tra 1 e 1000")
    '    'End If

    '    If _intMesi < 1 And _intMesi > 12 Then
    '        Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Il numero di MESI deve essere compreso tra 1 e 12")
    '    End If

    '    If _dblAcconto < 1 And _dblAcconto > 100 Then
    '        Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::La percentuale di ACCONTO deve essere compresa tra 1 e 100")
    '    End If

    '    If _dblPossesso < 0.1 And _dblPossesso > 100 Then
    '        Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::La percentuale di POSSESSO deve essere compresa tra 0.1 e 100")
    '    End If
    '    '****************************************************************************************************
    '    'FINE GESTIONE CONTROLLI VALORI PASSATI
    '    '****************************************************************************************************


    '    Try
    '        If _blnRiduzione Then
    '            dblCalcoloIci = (((_dblValore * (_dblAliquota / 1000)) / 12) * (_dblPossesso / 100) * _intMesi * (_dblAcconto / 100)) * dblRIDUZIONE
    '        Else
    '            dblCalcoloIci = ((_dblValore * (_dblAliquota / 1000)) / 12) * (_dblPossesso / 100) * _intMesi * (_dblAcconto / 100)
    '        End If

    '        'GIULIA 20060620
    '        'TESTO LA DETRAZIONE DICHIARATA PER IL CALCOLO DELLA DETRAZIONE DA APPLICARE
    '        'If _DBLDetrazioneDichiarata <> 0 Then
    '        '    dblCalcoloDetrazione = ((_DBLDetrazione / m_intMesiT) * m_intMesi * (_DBLAcconto / 100))
    '        'Else
    '        '    dblCalcoloDetrazione = (_DBLDetrazione * (_DBLAcconto / 100))
    '        'End If
    '        If _intAbitazionePrincipale = 1 Then


    '            If _intUtilizzatori > 0 Then
    '                dblCalcoloDetrazione = ((_dblDetrazione / 12) / _intUtilizzatori) * _intMesi * (_dblAcconto / 100)
    '                'Else
    '                '    dblCalcoloDetrazione = 0 '(_DBLDetrazione * (_DBLAcconto / 100))
    '            End If
    '        End If

    '        getCALCOLO_ICI_ACCONTO_TOTALE_PREmodificheIMU = dblCalcoloIci - dblCalcoloDetrazione

    '        _dblIci_Teorica = dblCalcoloIci
    '        _dblDetrazione_Applicata = dblCalcoloDetrazione

    '        If dblCalcoloIci - dblCalcoloDetrazione = 0 Then
    '            _dblIci_Dovuta = 0
    '            _dblDetrazione_Residua = 0
    '        End If

    '        If dblCalcoloIci - dblCalcoloDetrazione > 0 Then
    '            _dblIci_Dovuta = dblCalcoloIci - dblCalcoloDetrazione
    '            _dblDetrazione_Residua = 0
    '        End If

    '        If dblCalcoloIci - dblCalcoloDetrazione < 0 Then
    '            _dblIci_Dovuta = 0
    '            _dblDetrazione_Residua = dblCalcoloIci - dblCalcoloDetrazione
    '        End If

    '    Catch ex As Exception
    '        Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Errore durante la fase di calcolo dell'ICI")
    '    End Try
    'End Function
    ''' <summary>
    ''' Funzione di preparazione parametri per il calcolo effettivo.
    ''' Controllo valori passati e richiamo della funzione di calcolo per gli importi comunali e per gli importi statali
    ''' </summary>
    ''' <param name="oParamCalcolo"></param>
    ''' <param name="bIsEsenzione"></param>
    ''' <param name="nMesiEsenzione"></param>
    ''' <param name="impDetrazioneResidua"></param>
    ''' <param name="impDetrazResiduaStandard"></param>
    ''' <returns></returns>
    ''' <revisionHistory>
    ''' <revision date="20190910">
    ''' In caso di flag storico la riduzione deve essere applicata anche per l'abitazione principale
    ''' </revision>
    ''' </revisionHistory>
    Public Function CalcolaICI(ByVal oParamCalcolo As objCalcolo, ByVal bIsEsenzione As Boolean, ByVal nMesiEsenzione As Integer, ByRef impDetrazioneResidua As Double, ByRef impDetrazResiduaStandard As Double) As objCalcolo
        Dim oMyCalcolo As New objCalcolo

        Try
            Dim impCalcoloIci As Double = 0
            Dim impDetrazione As Double = 0
            Dim nPercCalcolo As Double = 0

            Const dblRIDUZIONE As Double = 0.5

            oMyCalcolo = oParamCalcolo
            oMyCalcolo.Ici_Teorica = 0
            oMyCalcolo.Ici_Dovuta = 0
            oMyCalcolo.Detrazione_Applicabile = 0
            oMyCalcolo.Detrazione_Residua = impDetrazioneResidua
            oMyCalcolo.Detrazione_Residua_Standard = impDetrazResiduaStandard
            oMyCalcolo.storico = oParamCalcolo.storico

            If oMyCalcolo.Aliquota = -1 Then
                Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::ALIQUTA ERRATA")
                Return Nothing
            End If
            If oMyCalcolo.Detrazione = -1 Then
                Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::DETRAZIONE ERRATA")
                Return Nothing
            End If

            If oMyCalcolo.Mesi < 1 And oMyCalcolo.Mesi > 12 Then
                Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Il numero di MESI deve essere compreso tra 1 e 12")
                Return Nothing
            End If

            If oMyCalcolo.Acconto < 1 And oMyCalcolo.Acconto > 100 Then
                Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::La percentuale di ACCONTO deve essere compresa tra 1 e 100")
                Return Nothing
            End If

            If oMyCalcolo.Possesso < 0.1 And oMyCalcolo.Possesso > 100 Then
                Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::La percentuale di POSSESSO deve essere compresa tra 0.1 e 100")
                Return Nothing
            End If
            '*** 20150430 - TASI Inquilino ***
            If oMyCalcolo.TipoTasi = "I" Then
                nPercCalcolo = oMyCalcolo.nPercInquilino / 100
            Else
                nPercCalcolo = (100 - oMyCalcolo.nPercInquilino) / 100
            End If
            'calcolo l'ici ordinaria/comunale
            If CalcolaImporti(oMyCalcolo.AnnoCalcolo, oMyCalcolo.Valore, oMyCalcolo.Aliquota, oMyCalcolo.Detrazione, oMyCalcolo.AbitazionePrincipale, oMyCalcolo.Possesso, oMyCalcolo.Utilizzatori, oMyCalcolo.Mesi, oMyCalcolo.Acconto, oMyCalcolo.Riduzione, dblRIDUZIONE, oMyCalcolo.FigliACarico, oMyCalcolo.PercentCaricoFigli, oMyCalcolo.DetrazioneFigli, impCalcoloIci, impDetrazione, impDetrazioneResidua, impDetrazResiduaStandard, bIsEsenzione, nMesiEsenzione, nPercCalcolo, oMyCalcolo.TipoAliquota, oMyCalcolo.storico) = False Then
                Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Errore durante la fase di calcolo dell'ICI")
                Return Nothing
            Else
                Log.Debug("getCALCOLO_ICI_ACCONTO_TOTALE::ho calcolato ici=" & impCalcoloIci.ToString)
                oMyCalcolo.Ici_Teorica = impCalcoloIci

                oMyCalcolo.Detrazione_Applicabile = impDetrazione
                oMyCalcolo.Ici_Dovuta = impCalcoloIci
                oMyCalcolo.Detrazione_Residua = impDetrazioneResidua
                oMyCalcolo.Detrazione_Residua_Standard = impDetrazResiduaStandard
            End If
            '*** 20120530 - IMU ***
            'calcolo l'ici statale
            oMyCalcolo.Ici_Teorica_Statale = 0
            oMyCalcolo.Ici_Dovuta_Statale = 0
            oMyCalcolo.Detrazione_Applicabile_Statale = 0
            oMyCalcolo.Detrazione_Residua_Statale = 0
            oMyCalcolo.Detrazione_Residua_Standard = 0

            If oMyCalcolo.AliquotaStatale > 0 Then
                If CalcolaImporti(oMyCalcolo.AnnoCalcolo, oMyCalcolo.Valore, oMyCalcolo.AliquotaStatale, oMyCalcolo.Detrazione, oMyCalcolo.AbitazionePrincipale, oMyCalcolo.Possesso, oMyCalcolo.Utilizzatori, oMyCalcolo.Mesi, oMyCalcolo.Acconto, oMyCalcolo.Riduzione, dblRIDUZIONE, oMyCalcolo.FigliACarico, oMyCalcolo.PercentCaricoFigli, oMyCalcolo.DetrazioneFigli, impCalcoloIci, impDetrazione, impDetrazioneResidua, oMyCalcolo.Detrazione_Residua_Standard, bIsEsenzione, nMesiEsenzione, nPercCalcolo, oMyCalcolo.TipoAliquota, oMyCalcolo.storico) = False Then
                    Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Errore durante la fase di calcolo dell'ICI")
                    Return Nothing
                Else
                    oMyCalcolo.Ici_Teorica_Statale = impCalcoloIci
                    oMyCalcolo.Detrazione_Applicabile_Statale = impDetrazione
                    oMyCalcolo.Ici_Dovuta_Statale = impCalcoloIci
                    oMyCalcolo.Detrazione_Residua_Statale = impDetrazioneResidua
                    oMyCalcolo.Detrazione_Residua_Standard = impDetrazResiduaStandard
                End If
            End If
            '*** ***
            '*** ***
            Return oMyCalcolo
        Catch ex As Exception
            Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Errore durante la fase di calcolo dell'ICI")
            Return Nothing
        End Try
    End Function
    'Public Function CalcolaICI(ByVal oParamCalcolo As objCalcolo, ByVal bIsEsenzione As Boolean, ByVal nMesiEsenzione As Integer, ByRef impDetrazioneResidua As Double, ByRef impDetrazResiduaStandard As Double) As objCalcolo
    '    Dim oMyCalcolo As New objCalcolo

    '    Try
    '        Dim impCalcoloIci As Double = 0
    '        Dim impDetrazione As Double = 0
    '        Dim nPercCalcolo As Double = 0

    '        Const dblRIDUZIONE As Double = 0.5

    '        oMyCalcolo = oParamCalcolo
    '        oMyCalcolo.Ici_Teorica = 0
    '        oMyCalcolo.Ici_Dovuta = 0
    '        oMyCalcolo.Detrazione_Applicabile = 0
    '        oMyCalcolo.Detrazione_Residua = impDetrazioneResidua
    '        oMyCalcolo.Detrazione_Residua_Standard = impDetrazResiduaStandard

    '        '****************************************************************************************************
    '        'GESTIONE CONTROLLI VALORI PASSATI
    '        '****************************************************************************************************
    '        If oMyCalcolo.Aliquota = -1 Then
    '            Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::ALIQUTA ERRATA")
    '            Return Nothing
    '        End If
    '        If oMyCalcolo.Detrazione = -1 Then
    '            Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::DETRAZIONE ERRATA")
    '            Return Nothing
    '        End If

    '        If oMyCalcolo.Mesi < 1 And oMyCalcolo.Mesi > 12 Then
    '            Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Il numero di MESI deve essere compreso tra 1 e 12")
    '            Return Nothing
    '        End If

    '        If oMyCalcolo.Acconto < 1 And oMyCalcolo.Acconto > 100 Then
    '            Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::La percentuale di ACCONTO deve essere compresa tra 1 e 100")
    '            Return Nothing
    '        End If

    '        If oMyCalcolo.Possesso < 0.1 And oMyCalcolo.Possesso > 100 Then
    '            Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::La percentuale di POSSESSO deve essere compresa tra 0.1 e 100")
    '            Return Nothing
    '        End If
    '        '****************************************************************************************************
    '        'FINE GESTIONE CONTROLLI VALORI PASSATI
    '        '****************************************************************************************************
    '        '*** 20150430 - TASI Inquilino ***
    '        If oMyCalcolo.TipoTasi = "I" Then
    '            nPercCalcolo = oMyCalcolo.nPercInquilino / 100
    '        Else
    '            nPercCalcolo = (100 - oMyCalcolo.nPercInquilino) / 100
    '        End If
    '        'calcolo l'ici ordinaria/comunale
    '        If CalcolaImporti(oMyCalcolo.AnnoCalcolo, oMyCalcolo.Valore, oMyCalcolo.Aliquota, oMyCalcolo.Detrazione, oMyCalcolo.AbitazionePrincipale, oMyCalcolo.Possesso, oMyCalcolo.Utilizzatori, oMyCalcolo.Mesi, oMyCalcolo.Acconto, oMyCalcolo.Riduzione, dblRIDUZIONE, oMyCalcolo.FigliACarico, oMyCalcolo.PercentCaricoFigli, oMyCalcolo.DetrazioneFigli, impCalcoloIci, impDetrazione, impDetrazioneResidua, impDetrazResiduaStandard, bIsEsenzione, nMesiEsenzione, nPercCalcolo, oMyCalcolo.TipoAliquota) = False Then
    '            'If CalcolaImportiICI(oMyCalcolo.AnnoCalcolo, oMyCalcolo.Valore, oMyCalcolo.Aliquota, oMyCalcolo.Detrazione, oMyCalcolo.AbitazionePrincipale, oMyCalcolo.Possesso, oMyCalcolo.Utilizzatori, oMyCalcolo.Mesi, oMyCalcolo.Acconto, oMyCalcolo.Riduzione, dblRIDUZIONE, oMyCalcolo.FigliACarico, oMyCalcolo.PercentCaricoFigli, oMyCalcolo.DetrazioneFigli, dblCalcoloIci, dblCalcoloDetrazione, bIsEsenzione, nMesiEsenzione) = False Then
    '            Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Errore durante la fase di calcolo dell'ICI")
    '            Return Nothing
    '        Else
    '            Log.Debug("getCALCOLO_ICI_ACCONTO_TOTALE::ho calcolato ici=" & impCalcoloIci.ToString)
    '            oMyCalcolo.Ici_Teorica = impCalcoloIci

    '            oMyCalcolo.Detrazione_Applicabile = impDetrazione
    '            oMyCalcolo.Ici_Dovuta = impCalcoloIci
    '            oMyCalcolo.Detrazione_Residua = impDetrazioneResidua
    '            oMyCalcolo.Detrazione_Residua_Standard = impDetrazResiduaStandard
    '        End If
    '        '*** 20120530 - IMU ***
    '        'calcolo l'ici statale
    '        oMyCalcolo.Ici_Teorica_Statale = 0
    '        oMyCalcolo.Ici_Dovuta_Statale = 0
    '        oMyCalcolo.Detrazione_Applicabile_Statale = 0
    '        oMyCalcolo.Detrazione_Residua_Statale = 0
    '        oMyCalcolo.Detrazione_Residua_Standard = 0

    '        If oMyCalcolo.AliquotaStatale > 0 Then
    '            If CalcolaImporti(oMyCalcolo.AnnoCalcolo, oMyCalcolo.Valore, oMyCalcolo.AliquotaStatale, oMyCalcolo.Detrazione, oMyCalcolo.AbitazionePrincipale, oMyCalcolo.Possesso, oMyCalcolo.Utilizzatori, oMyCalcolo.Mesi, oMyCalcolo.Acconto, oMyCalcolo.Riduzione, dblRIDUZIONE, oMyCalcolo.FigliACarico, oMyCalcolo.PercentCaricoFigli, oMyCalcolo.DetrazioneFigli, impCalcoloIci, impDetrazione, impDetrazioneResidua, oMyCalcolo.Detrazione_Residua_Standard, bIsEsenzione, nMesiEsenzione, nPercCalcolo, oMyCalcolo.TipoAliquota) = False Then
    '                'If CalcolaImportiICI(oMyCalcolo.AnnoCalcolo, oMyCalcolo.Valore, oMyCalcolo.AliquotaStatale, oMyCalcolo.Detrazione, oMyCalcolo.AbitazionePrincipale, oMyCalcolo.Possesso, oMyCalcolo.Utilizzatori, oMyCalcolo.Mesi, oMyCalcolo.Acconto, oMyCalcolo.Riduzione, dblRIDUZIONE, oMyCalcolo.FigliACarico, oMyCalcolo.PercentCaricoFigli, oMyCalcolo.DetrazioneFigli, dblCalcoloIci, dblCalcoloDetrazione, bIsEsenzione, nMesiEsenzione) = False Then
    '                Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Errore durante la fase di calcolo dell'ICI")
    '                Return Nothing
    '            Else
    '                oMyCalcolo.Ici_Teorica_Statale = impCalcoloIci
    '                oMyCalcolo.Detrazione_Applicabile_Statale = impDetrazione
    '                oMyCalcolo.Ici_Dovuta_Statale = impCalcoloIci
    '                oMyCalcolo.Detrazione_Residua_Statale = impDetrazioneResidua
    '                oMyCalcolo.Detrazione_Residua_Standard = impDetrazResiduaStandard
    '            End If
    '        End If
    '        '*** ***
    '        '*** ***
    '        Return oMyCalcolo
    '    Catch ex As Exception
    '        Throw New Exception("GESTIONE_CALCOLO_ICI::CALCOLO_ICI.CalcolaICI::Errore durante la fase di calcolo dell'ICI")
    '        Return Nothing
    '    End Try
    'End Function
    '*** 20150430 - TASI Inquilino ***
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="nAnnoCalcolo"></param>
    ''' <param name="nValoreUI"></param>
    ''' <param name="nAliquota"></param>
    ''' <param name="nDetrazione"></param>
    ''' <param name="IsAbiPrin"></param>
    ''' <param name="nPercPossesso"></param>
    ''' <param name="nUtilizzatori"></param>
    ''' <param name="nMesi"></param>
    ''' <param name="nPercAcconto"></param>
    ''' <param name="bIsRiduzione"></param>
    ''' <param name="nRiduzione"></param>
    ''' <param name="nFigliACarico"></param>
    ''' <param name="nPercentCaricoFigli"></param>
    ''' <param name="nDetrazioneFigli"></param>
    ''' <param name="impICI"></param>
    ''' <param name="impDetrazione"></param>
    ''' <param name="impDetrazioneResidua"></param>
    ''' <param name="impDetrazioneResiduaStandard"></param>
    ''' <param name="bIsEsenzione"></param>
    ''' <param name="nMesiEsenzione"></param>
    ''' <param name="nPerc"></param>
    ''' <param name="TipoAliquota"></param>
    ''' <param name="IsStorico"></param>
    ''' <returns></returns>
    ''' <revisionHistory>
    ''' <revision date="201804">
    ''' <strong>MINIIMU</strong>
    ''' se nel 2013 l'aliquota per abitazione principale e/o pertinenza è maggiore di 4 devo:
    ''' - calcolare con aliquota configurata
    ''' - calcolare con aliquota 4
    ''' - prendere il 40% della differenza
    ''' </revision>
    ''' </revisionHistory>
    ''' <revisionHistory>
    ''' <revision date="20190910">
    ''' In caso di flag storico la riduzione deve essere applicata anche per l'abitazione principale
    ''' </revision>
    ''' </revisionHistory>
    Private Function CalcolaImporti(ByVal nAnnoCalcolo As Integer, ByVal nValoreUI As Double, ByVal nAliquota As Double, ByVal nDetrazione As Double, ByVal IsAbiPrin As Integer, ByVal nPercPossesso As Double, ByVal nUtilizzatori As Integer, ByVal nMesi As Integer, ByVal nPercAcconto As Double, ByVal bIsRiduzione As Boolean, ByVal nRiduzione As Double, ByVal nFigliACarico As Integer, ByVal nPercentCaricoFigli As Double, ByVal nDetrazioneFigli As Double, ByRef impICI As Double, ByRef impDetrazione As Double, ByRef impDetrazioneResidua As Double, ByRef impDetrazioneResiduaStandard As Double, ByVal bIsEsenzione As Boolean, ByVal nMesiEsenzione As Integer, nPerc As Double, TipoAliquota As String, IsStorico As Integer) As Boolean
        Try
            'se sono abitazione principale non posso avere la riduzione
            If (IsAbiPrin = 1 Or IsAbiPrin = 2) And bIsRiduzione Then
                'se flag storico si
                If IsStorico = 0 Then
                    bIsRiduzione = False
                    nRiduzione = 0
                End If
            End If
            Log.Debug("CalcolaImporti.impICI = ((((nValoreUI * (nAliquota / 1000)) / 12) * (nPercPossesso / 100) * nMesi * (nPercAcconto / 100)) * nRiduzione)*nPerc -> valore=" & CStr(nValoreUI) & ",aliquota=" & CStr(nAliquota) & ",possesso=" & CStr(nPercPossesso) & ",n mesi=" & nMesi.ToString & ",% acconto=" & CStr(nPercAcconto) & ",nriduzione=" & nRiduzione.ToString & ",nPerc=" & nPerc.ToString & ",bIsRiduzione=" & bIsRiduzione.ToString)
            If bIsRiduzione Then
                impICI = ((((nValoreUI * (nAliquota / 1000)) / 12) * (nPercPossesso / 100) * nMesi * (nPercAcconto / 100)) * nRiduzione) * nPerc
            Else
                impICI = ((nValoreUI * (nAliquota / 1000)) / 12) * (nPercPossesso / 100) * nMesi * (nPercAcconto / 100) * nPerc
            End If

            If IsAbiPrin = 1 Then
                'se sono abitazione principale devo SEMPRE avere il numero di utilizzatori, quindi meglio calcolare con 1 che non calcolare la detrazione
                If nUtilizzatori <= 0 Then
                    nUtilizzatori = 1
                End If
                impDetrazione = ((nDetrazione / 12) / nUtilizzatori) * nMesi * (nPercAcconto / 100)
                '*** 20120530 - IMU ***
                'dal 2012 se ho dei figli a carico applico la riduzione per i figli altrimenti faccio il calcolo solito
                If nAnnoCalcolo >= 2012 And nFigliACarico > 0 Then
                    '*** 20120629 - IMU ***
                    'impDetrazione += ((((nDetrazioneFigli * nFigliACarico) * nPercentCaricoFigli) / 100 * nMesi) / 12) * (nPercAcconto / 100)
                    impDetrazione += ((nDetrazioneFigli * (nPercentCaricoFigli / 100) * nMesi) / 12) * (nPercAcconto / 100)
                    '*** ***
                End If
                '*** ***
            End If

            If bIsEsenzione = True And nMesi > 0 Then
                Log.Debug("CalcolaImporti.Esenzione.impICI = impICI * (nMesi - nMesiEsenzione) / nMesi  -> impICI=" & CStr(impICI) & ",nMesi=" & CStr(nMesi) & ",nMesiEsenzione=" & CStr(nMesiEsenzione))
                'i mesi di esenzione devono essere rapportati ai mesi di possesso
                If nMesiEsenzione > nMesi Then
                    nMesiEsenzione = nMesi
                End If
                impICI = impICI * (nMesi - nMesiEsenzione) / nMesi
            End If
            'tolgo la detrazione
            Log.Debug("CalcolaImporti.Detrazione.impICI = impICI - impDetrazione -> impICI=" & CStr(impICI) & ",impDetrazione=" & CStr(impDetrazione))
            If nAnnoCalcolo = 2013 And (IsAbiPrin = 1 Or IsAbiPrin = 2) And nAliquota = 4 Then
                If impDetrazioneResiduaStandard > 0 Then
                    impICI -= impDetrazioneResiduaStandard
                Else
                    impICI -= impDetrazione
                End If
            ElseIf IsAbiPrin = 1 Or IsAbiPrin = 2 Then
                If impDetrazioneResidua > 0 Then
                    impICI -= impDetrazioneResidua
                Else
                    impICI -= impDetrazione
                End If
            End If
            If impICI < 0 Then
                impDetrazioneResidua = impICI * -1
                impICI = 0
            Else
                impDetrazioneResidua = 0
            End If
            'controllo mini imu
            If nAnnoCalcolo = 2013 And (IsAbiPrin = 1 Or IsAbiPrin = 2) And nAliquota > 4 And TipoAliquota <> Generale.TipoAliquote_AS Then
                Dim impAliqStandard As Double = 0
                Dim impMiniIMU As Double = 0
                Log.Debug("GESTIONE_CALCOLO_ICI.CALCOLO_ICI.CalcolaImporti devo calcolare MINIIMU")
                If CalcolaImporti(nAnnoCalcolo, nValoreUI, 4, nDetrazione, IsAbiPrin, nPercPossesso, nUtilizzatori, nMesi, nPercAcconto, bIsRiduzione, nRiduzione, nFigliACarico, nPercentCaricoFigli, nDetrazioneFigli, impAliqStandard, impDetrazione, impDetrazioneResiduaStandard, impDetrazioneResiduaStandard, bIsEsenzione, nMesiEsenzione, nPerc, TipoAliquota, IsStorico) = False Then
                    Log.Debug("GESTIONE_CALCOLO_ICI.CALCOLO_ICI.CalcolaImporti.errore MINIIMU")
                    Return False
                Else
                    impMiniIMU = (((impICI - impAliqStandard) / 100) * 40)
                End If
                impICI = impMiniIMU
            End If
            If nAnnoCalcolo = 2013 And (IsAbiPrin = 1 Or IsAbiPrin = 2) And nAliquota = 4 Then
                impDetrazioneResiduaStandard = impDetrazioneResidua
            End If
            '*** ***
            Log.Debug("CalcolaImporti calcolato impICI=" & CStr(impICI))

            Return True
        Catch ex As Exception
            Throw New Exception("GESTIONE_CALCOLO_ICI:: CALCOLO_ICI::CalcolaImporti::Errore::" & ex.Message)
            Return False
        End Try
    End Function
    ''' <summary>
    '''**********************************************************************************************
    '''ATTENZIONE ATTUALMENTE NON VENGONO GESTITE LE MULTI ALIQUOTE
    '''**********************************************************************************************
    ''' DETERMINO L'ANNO PER L'ACCESSO ALLA TABELLA ALIQUOTE
    '''*** 20120509 - IMU**********************************************************************************
    '''NUOVA GESTIONE PER IMU E TASI
    '''ABITAZIONE PRINCIPALE:
    '''su IMU A/1, A/8, A/9 si paga ma senza detrazione dei figli, sulle altre categorie non si paga
    '''su TASI A/1, A/8, A/9 si paga con un'aliquota, sulle altre categorie ci sarà un'altra aliquota
    '''le detrazioni...
    '''TRIBUTO: 
    '''   Le aliquote sono prelevate in base al tributo che sto calcolando.
    '''SOGLIA RENDITA: 
    '''   Nella configurazione delle aliquote sarà possibile inserire una quota massima di rendita assoggettabile alla prima casa se "Uso gratuito ai famigliari".
    '''   Se sulla dichiarazione IMU il flag "abitazione principale" è settato e per quanto riguarda il  "tipo di utilizzo" è selezionato "Uso gratuito ai famigliari" allora il calcolo dell'IMU considererà solo la quota di rendita (comprensiva delle pertinenze) eccedente la quota massima di rendita assoggettabile alla prima casa a cui applicherà l'aliquota definita in relazione ad "Uso gratuito ai famigliari"
    '''   Es. soglia minima configurata per comodato d’uso = 500 
    '''       Rendita Abitazione Principale = 450 
    '''       Calcolo dovuto su rendita 450-500 = -50 quindi non si calcola
    '''       Rendita Pertinenza = 80
    '''       Calcolo dovuto su rendita 80-50(residuo da abitazione principale) = 30 quindi il dovuto utilizza la rendita = 30 
    '''ALIQUOTA C/2 C/6:
    '''   Nuova aliquota specifica per le categorie catastali C/2 e C/6 usata nei seguenti casi: 
    '''       - UI non identificato come pertinenza o fabbricato rurale;
    '''       - UI non identificato come abitazione principale;
    '''       - Tipo di utilizzo diverso dalle opzioni per le quali è prevista un’aliquota specifica ;
    '''       - Categoria catastale uguale a c/2 o c/6.
    '''TIPO UTILIZZO:
    '''   Attualmente il campo “tipo possesso” assume significati che possono essere identificati come tipo di possesso o come tipo di utilizzo. 
    '''   Per la corretta gestione del dato, lo stesso campo viene sdoppiato in due, e precisamente: “tipo utilizzo” e “tipo possesso”. 
    '''   Quest ultimo sarà usato nel calcolo ad ognuna delle voci sarà possibile attribuire un'aliquota
    '''   Nel campo tipo utilizzo saranno identificate le seguenti informazioni: 
    '''       - abitazione
    '''       - comodato d’uso 
    '''       - a disposizione / sfitto 
    '''       - locato 
    '''       - altri
    '''*** 20120530 - IMU**********************************************************************************
    '''NUOVA GESTIONE IMU IN BASE A FINANZIARIA 2012
    '''Il test per il reperimento aliquote all'interno del motore di calcolo ICI seguirà i seguenti passi:
    '''dal 2012 in poi
    '''prima dei passi per la finanziaria 2008 verifico se sono su coltivatore diretto per prelevarne la specifica aliquota
    '''le aliquote sono tutte configurate per parte statale e parte comunale
    '''le detrazioni per i figli sono calcolate con il seguente metodo:
    '''(((((detrazione*numero figli)*percentuale di carico dei figli)/100)*mesi possesso)/12)
    '''**********************************************************************************************
    '''NUOVA GESTIONE ICI IN BASE A FINANZIARIA 2008
    '''DIPE 25/03/2009
    '''Il test per il reperimento aliquote all'interno del motore di calcolo ICI seguirà i seguenti passi:
    '''- TIPO IMMOBILE C1 o C3 (Per Pomarance)
    '''	Reperimento della relativa aliquota e dell'eventuale detrazione configurata. 
    '''   Nel caso in cui l'immobile sia del tipo indicato non si devono cercare altre aliquote
    '''-Uso Gratuito 
    '''	Reperimento della relativa aliquota e dell'eventuale detrazione configurata. 
    '''	-Se è una categoria che paga (permettere configurazione categorie escluse dal pagamento o no), 
    '''    in tal caso utilizzare l'aliquota configurata 
    '''	-Se è una categoria che NON paga il dovuto sarà uguale a 0 (immobile esente) 
    '''-AIRE 
    '''	-Reperimento della relativa aliquota e dell'eventuale detrazione configurata.
    '''-Immobile Locato 
    '''	-Reperimento della relativa aliquota e dell'eventuale detrazione configurata.
    '''-Flag Abitazione Principale NON attivo
    '''	Controllo in ordine le seguenti tipologie di immobili per il reperimento della relativa aliquota. 
    '''		-Flag immobile sfitto 
    '''		-Flag immobile affitto convenzionato 
    '''		-Tipo fabbricato Area Fabbricabile 
    '''		-Tipo fabbricato Terreno Agricolo 
    '''		-Tipo Rendita RE/RP/RPM (Tipo fabbricato generico) 
    '''		-Tipo Rendita Libri Contabili 
    '''	In caso l'immobile non rientri nelle categorie citate viene utilizzata quella per altri fabbricati 
    '''-Flag Abitazione Principale è attivo
    '''	Reperimento della relativa aliquota e dell'eventuale detrazione configurata 
    '''		-Se è una categoria che paga verrà utilizzata l'aliquota configurata 
    '''		-se è una categoria che NON paga il dovuto sarà uguale a 0 (immobile esente) 
    '''-Se flag pertinenza è attivo
    '''	Il sistema reperisce l'eventuale aliquota configurata 
    '''       -calcolerà l'ici se la pertinenza è legata ad una abitazione principale/uso gratuito 
    '''         con categoria che paga 
    '''		-non calcolerà l'ici se è legata ad una abitazione principale con categoria che NON paga 
    '''**********************************************************************************************
    '''VERIFICA DEL TIPO ABITAZIONE: ---> PRINCIPALE O PERTINENZA
    '''DIPE 25/03/2009
    '''verifico il tipo di possesso dell'immobile e reperisco la relativa aliquota
    '''se 
    '''tipo possesso = USO GRATUITO 1° 2° 3° GRADO 
    '''tipo possesso = AIRE
    '''tipo possesso = LOCATO
    '''se tipo possesso non appartiene alle precedenti casistiche controllo il tipo abitazione
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="structPARAMETRI_ICI"></param>
    ''' <param name="Tributo"></param>
    ''' <returns></returns>
    ''' <revisionHistory><revision date="11/06/2021">Nuove Tipologie di Utilizzo3</revision></revisionHistory>
    Public Function getAliquotaDetrazione(myStringConnection As String, IdEnte As String, ByVal structPARAMETRI_ICI As Generale.PARAMETRI_ICI, ByVal Tributo As String) As ListALIQUOTA_DETRAZIONE
        Dim strANNO_CALCOLO_Aliquota As String = Now.Year.ToString
        Dim objDSResults As New DataSet
        Dim blnResultsOK As Boolean = False
        Dim objUtility As New Generale
        Dim oMyAliquote As New ListALIQUOTA_DETRAZIONE

        Try
            strANNO_CALCOLO_Aliquota = getCalcoloICI_Anno(structPARAMETRI_ICI.intANNO_CALCOLO, structPARAMETRI_ICI.strACCONTO_TOTALE)
            If strANNO_CALCOLO_Aliquota = "" Then
                Throw New Exception("getAliquotaDetrazione::errore in recupero Anno")
            End If
            Dim objDBOPENgovProvvedimentiSelect As New ClsDBManager

            Dim sTipoAliquota As String = String.Empty
            Dim sTipoDetrazione As String = String.Empty
            Dim blnFindTP As Boolean = False

            'If strANNO_CALCOLO_Aliquota >= 2012 And structPARAMETRI_ICI.IsColtivatoreDiretto = True Then
            '    '*** 20130422 - aggiornamento IMU ***
            '    If structPARAMETRI_ICI.strCATEGORIA = "D/10" Then
            '        sTipoAliquota = Generale.TipoAliquote_CDD10
            '        blnFindTP = True
            '        sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_CDD10
            '    Else
            '        sTipoAliquota = Generale.TipoAliquote_CD
            '        blnFindTP = True
            '        sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_CD
            '    End If
            '    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, strANNO_CALCOLO_Aliquota, sTipoAliquota, sTipoDetrazione, Tributo)
            '    If oMyAliquote.nIdAliquota <= 0 Then
            '        sTipoAliquota = Generale.TipoAliquote_A
            '        sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
            '        oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, strANNO_CALCOLO_Aliquota, sTipoAliquota, sTipoDetrazione, Tributo)
            '    End If
            '    '*** ***
            'ElseIf strANNO_CALCOLO_Aliquota < 2012 And structPARAMETRI_ICI.strCATEGORIA = "D/10" Then
            If strANNO_CALCOLO_Aliquota < 2012 And structPARAMETRI_ICI.strCATEGORIA = "D/10" Then
                    '*** 20120906 - IMU se sono prima del 2012 i D/10 non devono pagare
                    sTipoAliquota = ""
                    sTipoDetrazione = ""
                    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, strANNO_CALCOLO_Aliquota, sTipoAliquota, sTipoDetrazione, Tributo)
                    '*** ***
                ElseIf IdEnte = "050027" And (structPARAMETRI_ICI.strCATEGORIA = "C/1" Or structPARAMETRI_ICI.strCATEGORIA = "C/3") Then
                    sTipoAliquota = Generale.TipoAliquote_BO
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_BO
                    '*** 20130422 - aggiornamento IMU ***
                    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, strANNO_CALCOLO_Aliquota, sTipoAliquota, sTipoDetrazione, Tributo)
                    If oMyAliquote.nIdAliquota <= 0 Then
                        sTipoAliquota = Generale.TipoAliquote_A
                        sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
                        oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, strANNO_CALCOLO_Aliquota, sTipoAliquota, sTipoDetrazione, Tributo)
                    End If
                    '*** ***
                Else
                    oMyAliquote = PrelevaAliquoteVSTipoRendita(myStringConnection, IdEnte, strANNO_CALCOLO_Aliquota, structPARAMETRI_ICI, Tributo)
            End If
        Catch ex As Exception
            Log.Debug("getAliquotaDetrazione:.si è verificato il seguente errore::", ex)
        End Try
        Return oMyAliquote
    End Function

    Private Function getCalcoloICI_Anno(ByVal AnnoCalcolo As Integer, ByVal sTipoCalcolo As String) As String
        Dim myRet As String = ""
        Dim AnnoRiferimento As Integer = 0
        Dim objUtility As New Generale

        Try
            '*** 20120530 - IMU ***
            If AnnoCalcolo < 2012 Then
                AnnoRiferimento = AnnoCalcolo
                If sTipoCalcolo.CompareTo("ACCONTO") = 0 Then
                    Select Case AnnoRiferimento
                        Case Is < Generale.ANNO_CALCOLO
                            myRet = objUtility.CToStr(AnnoRiferimento, True, False, False)
                        Case Is >= Generale.ANNO_CALCOLO
                            AnnoRiferimento = AnnoRiferimento - 1
                            myRet = objUtility.CToStr(AnnoRiferimento, True, False, False)
                    End Select
                End If
            Else
                AnnoRiferimento = AnnoCalcolo
                myRet = AnnoCalcolo
            End If
            '*** ***
            If sTipoCalcolo.CompareTo(Generale.TOTALE) = 0 Then
                myRet = objUtility.CToStr(AnnoRiferimento, True, False, False)
            End If

            Return myRet
        Catch ex As Exception
            Log.Debug("getCalcoloICI_Anno::si è verificato il seguente errore::", ex)
            Return ""
        End Try
    End Function

    Private Function PrelevaAliquoteVSTipoRendita(myStringConnection As String, IdEnte As String, ByVal sAnno As String, ByVal myParamCalcolo As Freezer.Generale.PARAMETRI_ICI, ByVal Tributo As String) As ListALIQUOTA_DETRAZIONE
        Dim sTipoAliquota As String = String.Empty
        Dim sTipoDetrazione As String = String.Empty
        Dim blnFindTP As Boolean = False
        Dim oMyAliquote As New ListALIQUOTA_DETRAZIONE

        Try
            'altrimenti testo il tipo rendita
            Select Case myParamCalcolo.strTIPO_RENDITA
                Case Generale.TipoRendita_AF                                      'AREE EDIFICABILI
                    sTipoAliquota = Generale.TipoAliquote_AAF
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AAF
                Case Generale.TipoRendita_TA                                        'TERRENI AGRICOLI
                    sTipoAliquota = Generale.TipoAliquote_TTAA
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_TTAA
                Case Else                                  'RENDITA EFFETTIVA,RENDITA PRESUNTA,RENDITA PRESUNTA MODIFICATA,LIBRI CONTABILI
                    oMyAliquote = PrelevaAliquoteVSTipoUtilizzo(myStringConnection, IdEnte, sAnno, myParamCalcolo, Tributo)
            End Select
            '*** ***
            If blnFindTP = True Then
                '*** 20130422 - aggiornamento IMU ***
                oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
                '*** ***
            End If
            Return oMyAliquote
        Catch ex As Exception
            Return New ListALIQUOTA_DETRAZIONE
        End Try
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="sAnno"></param>
    ''' <param name="myParamCalcolo"></param>
    ''' <param name="Tributo"></param>
    ''' <returns></returns>
    ''' <revisionHistory><revision date="11/06/2021">Nuove Tipologie di Utilizzo</revision></revisionHistory>
    Private Function PrelevaAliquoteVSTipoUtilizzo(myStringConnection As String, IdEnte As String, ByVal sAnno As String, ByVal myParamCalcolo As Freezer.Generale.PARAMETRI_ICI, ByVal Tributo As String) As ListALIQUOTA_DETRAZIONE
        Dim sTipoAliquota As String = String.Empty
        Dim sTipoDetrazione As String = String.Empty
        Dim blnFindTP As Boolean = False
        Dim oMyAliquote As New ListALIQUOTA_DETRAZIONE
        Dim objDBOPENgovProvvedimentiSelect As New ClsDBManager
        Dim dsCatDaEscludere As DataSet

        Try
            Select Case myParamCalcolo.IdTipoUtilizzo
                Case Generale.TitoloPossesso_UG1  'USO GRATUITO 1° GRADO
                    sTipoAliquota = Generale.TipoAliquote_AUG1
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AUG1
                Case Generale.TitoloPossesso_UG2  'USO GRATUITO 2° GRADO
                    sTipoAliquota = Generale.TipoAliquote_AUG2
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AUG2
                Case Generale.TitoloPossesso_UG3  'USO GRATUITO 3° GRADO
                    sTipoAliquota = Generale.TipoAliquote_AUG3
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AUG3
                Case Generale.TitoloPossesso_AIRE  'AIRE
                    sTipoAliquota = Generale.TipoAliquote_AAIRE
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AAIRE
                Case Generale.TitoloPossesso_LOC  'IMMOBILE LOCATO
                    sTipoAliquota = Generale.TipoAliquote_AL
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AL
                Case Generale.TitoloPossesso_APEX  'Detrazione Ex 104/92
                    sTipoAliquota = Generale.TipoAliquote_AAPEX
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AAPEX
                Case Generale.TitoloPossesso_SAD  'SFITTI/A DISPOSIZIONE
                    sTipoAliquota = Generale.TipoAliquote_S
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_S
                Case Generale.TitoloPossesso_AFC  'AFFITTI CONVENZIONATI
                    sTipoAliquota = Generale.TipoAliquote_AC
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AC
                Case Generale.TitoloPossesso_STO  'Storico	
                    sTipoAliquota = Generale.TipoAliquote_STO
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_STO
                Case Generale.TitoloPossesso_RUR  'Rurale	
                    sTipoAliquota = Generale.TipoAliquote_RUR
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_RUR
                Case Generale.TitoloPossesso_IACP  'IMU per l'Agenzia Territoriale per la Casa (ex IACP)	
                    sTipoAliquota = Generale.TipoAliquote_IACP
                    blnFindTP = True
                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_IACP
            End Select
            If blnFindTP = True Then
                '*** 20130422 - aggiornamento IMU ***
                oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
                If oMyAliquote.nIdAliquota <= 0 Then
                    oMyAliquote = PrelevaAliquoteVSTipoAbitazione(myStringConnection, IdEnte, sAnno, myParamCalcolo, Tributo)
                End If
                dsCatDaEscludere = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(myStringConnection, IdEnte, sAnno, sTipoAliquota, Tributo)
                If dsCatDaEscludere.Tables(0).Select("COD_CAT='" & myParamCalcolo.strCATEGORIA & "'").Length > 0 Then
                    oMyAliquote.p_ESENTE = 1
                End If
                '*** ***
                'la soglia rendita da applicare solo se abitazione principale
                If myParamCalcolo.intTIPO_ABITAZIONE <> Generale.ABITAZIONE_PRINCIPALE_PERTINENZA.ABITAZIONE_PRINCIPALE Then
                    oMyAliquote.nSogliaRendita = 0
                End If
            Else 'If blnFindTP = False Then
                oMyAliquote = PrelevaAliquoteVSTipoAbitazione(myStringConnection, IdEnte, sAnno, myParamCalcolo, Tributo)
            End If
            Return oMyAliquote
        Catch ex As Exception
            Return New ListALIQUOTA_DETRAZIONE
        End Try
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="myStringConnection"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="sAnno"></param>
    ''' <param name="myParamCalcolo"></param>
    ''' <param name="Tributo"></param>
    ''' <returns></returns>
    ''' <revisionHistory><revision date="07/09/2021">le tariffe specifiche sono per D/5 e D/1 anzichè D/8</revision></revisionHistory>
    Private Function PrelevaAliquoteVSTipoAbitazione(myStringConnection As String, IdEnte As String, ByVal sAnno As String, ByVal myParamCalcolo As Freezer.Generale.PARAMETRI_ICI, ByVal Tributo As String) As ListALIQUOTA_DETRAZIONE
        Dim sTipoAliquota As String = String.Empty
        Dim sTipoDetrazione As String = String.Empty
        Dim blnFindTP As Boolean = False
        Dim dsCatDaEscludere As DataSet
        Dim oMyAliquote As New ListALIQUOTA_DETRAZIONE
        Dim objDBOPENgovProvvedimentiSelect As New ClsDBManager
        Dim IdAliquota As Integer
        Dim nAliquotaStatale As Double
        Dim nPercInquilino As Double = 0

        Try
            Select Case myParamCalcolo.intTIPO_ABITAZIONE
                Case Generale.ABITAZIONE_PRINCIPALE_PERTINENZA.ABITAZIONE_PRINCIPALE
                    '*****************************************************************************************
                    'SE E' UNA ABITAZIONE PRINCIPALE SI ACCEDE ALLA TABELLA ALIQUOTE 
                    'UTILIZZANDO COME PARAMETRO (TIPO ALIQUOTA 'AAP') L'ANNO=strANNO_CALCOLO_Aliquota E UTILIZZANDO COME PARAMETRO (TIPO ALIQUOTA 'D') L'ANNO=strANNO_CALCOLO_Aliquota
                    'LA FUNZIONE CHIAMATA DOVRA' RESTITUIRE IL VALORE DELL'ALIQUOTA E DELLA DETRAZIONE
                    '*****************************************************************************************
                    'Log.Debug("ABITAZIONE_PRINCIPALE")
                    '*** 20130422 - aggiornamento IMU ***
                    Select Case myParamCalcolo.strCATEGORIA
                        Case "A/1", "A/8", "A/9"
                            sTipoAliquota = Generale.TipoAliquote_AS
                        Case Else
                            sTipoAliquota = Generale.TipoAliquote_AAP
                    End Select
                    sTipoDetrazione = Generale.TipoAliquote_D & sTipoAliquota
                    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
                    If sTipoAliquota = Generale.TipoAliquote_AS And oMyAliquote.nIdAliquota <= 0 Then
                        sTipoAliquota = Generale.TipoAliquote_AAP
                        sTipoDetrazione = Generale.TipoAliquote_D & sTipoAliquota
                        oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
                    End If
                    If sAnno >= 2012 Then
                        '*** 20150430 - TASI Inquilino ***
                        oMyAliquote.nDetrazioneFigli = objDBOPENgovProvvedimentiSelect.getAliquote(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_DFAAP, Tributo, nAliquotaStatale, IdAliquota, 0, "", nPercInquilino)
                        '*** ***
                    End If
                    dsCatDaEscludere = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_AAP, Tributo)
                    If dsCatDaEscludere.Tables(0).Select("COD_CAT='" & myParamCalcolo.strCATEGORIA & "'").Length > 0 Then
                        oMyAliquote.p_ESENTE = 1
                    End If
                    '*** ***
                Case Generale.ABITAZIONE_PRINCIPALE_PERTINENZA.ABITAZIONE_PERTINENZA
                    '*****************************************************************************************
                    'SE E' UNA PERTINENZA SI ACCEDE ALLA TABELLA ALIQUOTE 
                    'UTILIZZANDO COME PARAMETRO (TIPO ALIQUOTA 'P') L'ANNO=strANNO_CALCOLO_Aliquota
                    'LA FUNZIONE CHIAMATA DOVRA' RESTITUIRE IL VALORE DELL'ALIQUOTA
                    '*****************************************************************************************
                    '*** 20130422 - aggiornamento IMU ***
                    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_P, Generale.TipoAliquote_D & Generale.TipoAliquote_P, Tributo)
                    If sAnno >= 2012 Then
                        '*** 20150430 - TASI Inquilino ***
                        oMyAliquote.nDetrazioneFigli = objDBOPENgovProvvedimentiSelect.getAliquote(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_DFAAP, Tributo, nAliquotaStatale, IdAliquota, 0, "", nPercInquilino)
                        '*** ***
                    End If
                    '*** 201805 - se la pertinenza è riferita ad una principale esclusa devo esludere anche lei ***
                    dsCatDaEscludere = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_AAP, Tributo)
                    '*** ***
                    If dsCatDaEscludere.Tables(0).Select("COD_CAT='" & myParamCalcolo.Categoria_AAP & "'").Length > 0 Then
                        oMyAliquote.p_ESENTE = 1
                    End If
                    '*** ***
                    'se non è nè abitazione principale nè pertinenza
                Case Else 'Utility.ABITAZIONE_PRINCIPALE_PERTINENZA.NO_ABITAZIONE_PERTINENZA
                    '*** 20130422 - aggiornamento IMU ****
                    If myParamCalcolo.strCATEGORIA.StartsWith("D") = True Then
                        '*** 201801 - aliquote specifiche sui D ***
                        ''BD 1/10/2021 Ripristinato il D8 per VIGLIANO NON FUNZIONERA' PIU POMARANCE CON D1 e D5
                        ''If myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper() = Generale.TipoAliquote_D5 Or myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper() = Generale.TipoAliquote_D1 Then
                        ''BD 1/10/2021 OCCORRERA' GESTIRE ENTRAMBI I CASI
                        ''
                        If myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper() = Generale.TipoAliquote_D8 Then
                            sTipoAliquota = myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper()
                            blnFindTP = True
                            sTipoDetrazione = Generale.TipoAliquote_D & myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper()
                            oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
                            If oMyAliquote.nIdAliquota <= 0 Then
                                sTipoAliquota = Generale.TipoAliquote_AFD                                     'altri fabbricati categoria D
                                blnFindTP = True
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AFD
                            End If
                        Else
                            sTipoAliquota = Generale.TipoAliquote_AFD                                     'altri fabbricati categoria D
                            blnFindTP = True
                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AFD
                        End If
                        '*** ***
                    Else
                        'altrimenti testo il tipo rendita
                        Select Case myParamCalcolo.strTIPO_RENDITA
                            Case Generale.TipoRendita_AF                                      'AREE EDIFICABILI
                                sTipoAliquota = Generale.TipoAliquote_AAF
                                blnFindTP = True
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AAF
                            Case Generale.TipoRendita_TA                                        'TERRENI AGRICOLI
                                sTipoAliquota = Generale.TipoAliquote_TTAA
                                blnFindTP = True
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_TTAA
                            Case Generale.TipoRendita_RE                                        'RENDITA EFFETTIVA
                                sTipoAliquota = Generale.TipoAliquote_A
                                blnFindTP = True
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
                            Case Generale.TipoRendita_RP                                        'RENDITA PRESUNTA
                                sTipoAliquota = Generale.TipoAliquote_A
                                blnFindTP = True
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
                            Case Generale.TipoRendita_RPM                                         'RENDITA PRESUNTA MODIFICATA
                                sTipoAliquota = Generale.TipoAliquote_A
                                blnFindTP = True
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
                            Case Generale.TipoRendita_LC                                        'LIBRI CONTABILI
                                sTipoAliquota = Generale.TipoAliquote_A
                                blnFindTP = True
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
                        End Select
                    End If
                    '*** ***
                    If blnFindTP = True Then
                        Select Case myParamCalcolo.strCATEGORIA
                            Case "C/2"
                                sTipoAliquota = Generale.TipoAliquote_C2
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_C2
                            Case "C/6"
                                sTipoAliquota = Generale.TipoAliquote_C6
                                sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_C6
                        End Select
                        '*** 20130422 - aggiornamento IMU ***
                        oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
                        '*** ***
                    End If
                    If oMyAliquote.nIdAliquota <= 0 Then
                        sTipoAliquota = Generale.TipoAliquote_A
                        sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
                        oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
                    End If
            End Select
            Return oMyAliquote
        Catch ex As Exception
            Return New ListALIQUOTA_DETRAZIONE
        End Try
    End Function
    'Private Function PrelevaAliquoteVSTipoAbitazione(myStringConnection As String, IdEnte As String, ByVal sAnno As String, ByVal myParamCalcolo As Freezer.Generale.PARAMETRI_ICI, ByVal Tributo As String) As ListALIQUOTA_DETRAZIONE
    '    Dim sTipoAliquota As String = String.Empty
    '    Dim sTipoDetrazione As String = String.Empty
    '    Dim blnFindTP As Boolean = False
    '    Dim dsCatDaEscludere As DataSet
    '    Dim oMyAliquote As New ListALIQUOTA_DETRAZIONE
    '    Dim objDBOPENgovProvvedimentiSelect As New ClsDBManager
    '    Dim IdAliquota As Integer
    '    Dim nAliquotaStatale As Double
    '    Dim nPercInquilino As Double = 0

    '    Try
    '        Select Case myParamCalcolo.intTIPO_ABITAZIONE
    '            Case Generale.ABITAZIONE_PRINCIPALE_PERTINENZA.ABITAZIONE_PRINCIPALE
    '                '*****************************************************************************************
    '                'SE E' UNA ABITAZIONE PRINCIPALE SI ACCEDE ALLA TABELLA ALIQUOTE 
    '                'UTILIZZANDO COME PARAMETRO (TIPO ALIQUOTA 'AAP') L'ANNO=strANNO_CALCOLO_Aliquota E UTILIZZANDO COME PARAMETRO (TIPO ALIQUOTA 'D') L'ANNO=strANNO_CALCOLO_Aliquota
    '                'LA FUNZIONE CHIAMATA DOVRA' RESTITUIRE IL VALORE DELL'ALIQUOTA E DELLA DETRAZIONE
    '                '*****************************************************************************************
    '                'Log.Debug("ABITAZIONE_PRINCIPALE")
    '                '*** 20130422 - aggiornamento IMU ***
    '                Select Case myParamCalcolo.strCATEGORIA
    '                    Case "A/1", "A/8", "A/9"
    '                        sTipoAliquota = Generale.TipoAliquote_AS
    '                    Case Else
    '                        sTipoAliquota = Generale.TipoAliquote_AAP
    '                End Select
    '                sTipoDetrazione = Generale.TipoAliquote_D & sTipoAliquota
    '                oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
    '                If sTipoAliquota = Generale.TipoAliquote_AS And oMyAliquote.nIdAliquota <= 0 Then
    '                    sTipoAliquota = Generale.TipoAliquote_AAP
    '                    sTipoDetrazione = Generale.TipoAliquote_D & sTipoAliquota
    '                    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
    '                End If
    '                If sAnno >= 2012 Then
    '                    '*** 20150430 - TASI Inquilino ***
    '                    oMyAliquote.nDetrazioneFigli = objDBOPENgovProvvedimentiSelect.getAliquote(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_DFAAP, Tributo, nAliquotaStatale, IdAliquota, 0, "", nPercInquilino)
    '                    '*** ***
    '                End If
    '                dsCatDaEscludere = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_AAP, Tributo)
    '                If dsCatDaEscludere.Tables(0).Select("COD_CAT='" & myParamCalcolo.strCATEGORIA & "'").Length > 0 Then
    '                    oMyAliquote.p_ESENTE = 1
    '                End If
    '                '*** ***
    '            Case Generale.ABITAZIONE_PRINCIPALE_PERTINENZA.ABITAZIONE_PERTINENZA
    '                '*****************************************************************************************
    '                'SE E' UNA PERTINENZA SI ACCEDE ALLA TABELLA ALIQUOTE 
    '                'UTILIZZANDO COME PARAMETRO (TIPO ALIQUOTA 'P') L'ANNO=strANNO_CALCOLO_Aliquota
    '                'LA FUNZIONE CHIAMATA DOVRA' RESTITUIRE IL VALORE DELL'ALIQUOTA
    '                '*****************************************************************************************
    '                '*** 20130422 - aggiornamento IMU ***
    '                oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_P, Generale.TipoAliquote_D & Generale.TipoAliquote_P, Tributo)
    '                If sAnno >= 2012 Then
    '                    '*** 20150430 - TASI Inquilino ***
    '                    oMyAliquote.nDetrazioneFigli = objDBOPENgovProvvedimentiSelect.getAliquote(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_DFAAP, Tributo, nAliquotaStatale, IdAliquota, 0, "", nPercInquilino)
    '                    '*** ***
    '                End If
    '                '*** 201805 - se la pertinenza è riferita ad una principale esclusa devo esludere anche lei ***
    '                dsCatDaEscludere = objDBOPENgovProvvedimentiSelect.getCategorieDaEscludere(myStringConnection, IdEnte, sAnno, Generale.TipoAliquote_AAP, Tributo)
    '                '*** ***
    '                If dsCatDaEscludere.Tables(0).Select("COD_CAT='" & myParamCalcolo.Categoria_AAP & "'").Length > 0 Then
    '                    oMyAliquote.p_ESENTE = 1
    '                End If
    '                '*** ***
    '                'se non è nè abitazione principale nè pertinenza
    '            Case Else 'Utility.ABITAZIONE_PRINCIPALE_PERTINENZA.NO_ABITAZIONE_PERTINENZA
    '                '*** 20130422 - aggiornamento IMU ****
    '                If myParamCalcolo.strCATEGORIA.StartsWith("D") = True Then
    '                    '*** 201801 - aliquote specifiche sui D ***
    '                    If myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper() = Generale.TipoAliquote_D5 Or myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper() = Generale.TipoAliquote_D8 Then
    '                        sTipoAliquota = myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper()
    '                        blnFindTP = True
    '                        sTipoDetrazione = Generale.TipoAliquote_D & myParamCalcolo.strCATEGORIA.Replace("/", "").ToUpper()
    '                        oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
    '                        If oMyAliquote.nIdAliquota <= 0 Then
    '                            sTipoAliquota = Generale.TipoAliquote_AFD                                     'altri fabbricati categoria D
    '                            blnFindTP = True
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AFD
    '                        End If
    '                    Else
    '                        sTipoAliquota = Generale.TipoAliquote_AFD                                     'altri fabbricati categoria D
    '                        blnFindTP = True
    '                        sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AFD
    '                    End If
    '                    '*** ***
    '                Else
    '                    'altrimenti testo il tipo rendita
    '                    Select Case myParamCalcolo.strTIPO_RENDITA
    '                        Case Generale.TipoRendita_AF                                      'AREE EDIFICABILI
    '                            sTipoAliquota = Generale.TipoAliquote_AAF
    '                            blnFindTP = True
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_AAF
    '                        Case Generale.TipoRendita_TA                                        'TERRENI AGRICOLI
    '                            sTipoAliquota = Generale.TipoAliquote_TTAA
    '                            blnFindTP = True
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_TTAA
    '                        Case Generale.TipoRendita_RE                                        'RENDITA EFFETTIVA
    '                            sTipoAliquota = Generale.TipoAliquote_A
    '                            blnFindTP = True
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
    '                        Case Generale.TipoRendita_RP                                        'RENDITA PRESUNTA
    '                            sTipoAliquota = Generale.TipoAliquote_A
    '                            blnFindTP = True
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
    '                        Case Generale.TipoRendita_RPM                                         'RENDITA PRESUNTA MODIFICATA
    '                            sTipoAliquota = Generale.TipoAliquote_A
    '                            blnFindTP = True
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
    '                        Case Generale.TipoRendita_LC                                        'LIBRI CONTABILI
    '                            sTipoAliquota = Generale.TipoAliquote_A
    '                            blnFindTP = True
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
    '                    End Select
    '                End If
    '                '*** ***
    '                If blnFindTP = True Then
    '                    Select Case myParamCalcolo.strCATEGORIA
    '                        Case "C/2"
    '                            sTipoAliquota = Generale.TipoAliquote_C2
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_C2
    '                        Case "C/6"
    '                            sTipoAliquota = Generale.TipoAliquote_C6
    '                            sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_C6
    '                    End Select
    '                    '*** 20130422 - aggiornamento IMU ***
    '                    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
    '                    '*** ***
    '                End If
    '                If oMyAliquote.nIdAliquota <= 0 Then
    '                    sTipoAliquota = Generale.TipoAliquote_A
    '                    sTipoDetrazione = Generale.TipoAliquote_D & Generale.TipoAliquote_A
    '                    oMyAliquote = PrelevaAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, sTipoDetrazione, Tributo)
    '                End If
    '        End Select
    '        Return oMyAliquote
    '    Catch ex As Exception
    '        Return New ListALIQUOTA_DETRAZIONE
    '    End Try
    'End Function
    '*** 20150430 - TASI Inquilino ***
    Private Function PrelevaAliquote(myStringConnection As String, ByVal IdEnte As String, ByVal sAnno As String, ByVal sTipoAliquota As String, ByVal sTipoDetrazione As String, ByVal Tributo As String) As ListALIQUOTA_DETRAZIONE
        Dim oMyAliquote As New ListALIQUOTA_DETRAZIONE
        Dim IdAliquota As Integer = 0
        Dim nAliquotaStatale As Double = 0
        Dim nSogliaRendita As Double = 0
        Dim sTipoSoglia As String = ">"
        Dim nPercInquilino As Double = 0
        Dim objDBOPENgovProvvedimentiSelect As New ClsDBManager

        Try
            'MODIFICA GIULIA
            'QUANDO REPERISCO L'ALIQUOTA USO ANCHE LA CATEGORIA PER LA GESTIONE PIU' PUNTUALE
            'DELLE MULTIALIQUOTE (QUINDI IL CAMPO DEFAULT DIVENTA INUTILE)
            'IN QUESTO MODO POSSO AVERE ALIQUOTE DIFFERENZIATE PER TIPOLOGIA DI CATEGORIA OLTRE CHE PER IMMOBILE
            oMyAliquote.p_dblVALORE_ALIQUOTA = objDBOPENgovProvvedimentiSelect.getAliquote(myStringConnection, IdEnte, sAnno, sTipoAliquota, Tributo, nAliquotaStatale, IdAliquota, nSogliaRendita, sTipoSoglia, nPercInquilino)
            oMyAliquote.AliquotaStatale = nAliquotaStatale
            oMyAliquote.nIdAliquota = IdAliquota
            oMyAliquote.nPercInquilino = nPercInquilino
            oMyAliquote.p_dblVALORE_DETRAZIONE = objDBOPENgovProvvedimentiSelect.getAliquote(myStringConnection, IdEnte, sAnno, sTipoDetrazione, Tributo, nAliquotaStatale, IdAliquota, nSogliaRendita, sTipoSoglia, nPercInquilino)
            oMyAliquote.nSogliaRendita = nSogliaRendita
            oMyAliquote.sTipoSoglia = sTipoSoglia
            oMyAliquote.TipoAliquota = sTipoAliquota
            Return oMyAliquote
        Catch ex As Exception
            Log.Debug("PrelevaAliquote::si è verificato il seguente errore::", ex)
            Return New ListALIQUOTA_DETRAZIONE
        End Try
    End Function
    '*** ***
End Class
