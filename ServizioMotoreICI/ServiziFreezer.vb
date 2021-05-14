Imports log4net
Imports ComPlusInterface
Imports Freezer
''' <summary>
''' Classe rende disponibili le interfacce di calcolo 
''' </summary>
Public Class ServiziFreezer
    Inherits MarshalByRefObject : Implements IFreezer
    Private Shared ReadOnly Log As ILog = LogManager.GetLogger(GetType(ServiziFreezer))
    ''' <summary>
    ''' Interfaccia per il calcolo degli importi dovuti dei soggetti in ingresso
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="nIdContribuente"></param>
    ''' <param name="TributoCalcolo"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="AnnoDa"></param>
    ''' <param name="AnnoA"></param>
    ''' <param name="IsMassivo"></param>
    ''' <param name="blnConfigurazioneDich"></param>
    ''' <param name="blnRibaltaVersatoSuDovuto"></param>
    ''' <param name="blnCalcolaArrotondamento"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <param name="TipoTASI"></param>
    ''' <param name="TASIAProprietario"></param>
    ''' <param name="TipoOperazione"></param>
    ''' <param name="Operatore"></param>
    ''' <param name="ListSituazioneFinale"></param>
    ''' <returns></returns>
    Public Function CalcoloFromSoggetto(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, nIdContribuente As Integer, TributoCalcolo As String, Tributo As String, AnnoDa As String, AnnoA As String, IsMassivo As Boolean, ByVal blnConfigurazioneDich As Boolean, ByVal blnRibaltaVersatoSuDovuto As Boolean, ByVal blnCalcolaArrotondamento As Boolean, ByVal TipoCalcolo As Integer, TipoTASI As String, TASIAProprietario As String, TipoOperazione As String, Operatore As String, ByRef ListSituazioneFinale() As objSituazioneFinale) As Boolean Implements ComPlusInterface.IFreezer.CalcoloFromSoggetto
        Try
            Return New ClsFreezer().CalcoloICIcompletoAsync(StringConnectionGOV, StringConnectionICI, IdEnte, nIdContribuente, TributoCalcolo, Tributo, AnnoDa, AnnoA, IsMassivo, blnConfigurazioneDich, blnRibaltaVersatoSuDovuto, blnCalcolaArrotondamento, TipoCalcolo, TipoTASI, TASIAProprietario, TipoOperazione, Operatore, ListSituazioneFinale)
        Catch ex As Exception
            Throw New Exception(ex.Message & "::" & ex.StackTrace)
        End Try
    End Function
    ''' <summary>
    ''' Interfaccia per il calcolo degli importi dovuti su una lista di immobili in ingresso
    ''' </summary>
    ''' <param name="StringConnectionGOV"></param>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="ListUI"></param>
    ''' <param name="TipoCalcolo"></param>
    ''' <param name="ListSituazioneFinale"></param>
    ''' <returns></returns>
    Public Function CalcoloFromUI(StringConnectionGOV As String, StringConnectionICI As String, IdEnte As String, ListUI As ArrayList, ByVal TipoCalcolo As Integer, ByRef ListSituazioneFinale() As objSituazioneFinale) As Boolean Implements ComPlusInterface.IFreezer.CalcoloFromUI
        Try
            Log.Debug("entrata")
            ListSituazioneFinale = New ClsFreezer().CalcoloICISingle(StringConnectionGOV, StringConnectionICI, IdEnte, ListUI, TipoCalcolo)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message & "::" & ex.StackTrace)
        End Try
    End Function
    ''' <summary>
    ''' Interfaccia per la determinazione dei mesi IMU applicabili
    ''' </summary>
    ''' <param name="dInizio"></param>
    ''' <param name="dFine"></param>
    ''' <param name="nAnno"></param>
    ''' <returns></returns>
    Public Function CalcolaMesi(dInizio As Date, dFine As Date, nAnno As Integer) As Integer Implements ComPlusInterface.IFreezer.CalcolaMesi
        Try
            Return New Generale().mesi_possesso(dInizio, dFine, nAnno, 0, 0)
        Catch ex As Exception
            Throw New Exception(ex.Message & "::" & ex.StackTrace)
        End Try
    End Function
    ''' <summary>
    ''' Interfaccia per il salvataggio del calcolo
    ''' </summary>
    ''' <param name="myConnectionString"></param>
    ''' <param name="IdEnte"></param>
    ''' <param name="objICI"></param>
    ''' <param name="nIDElaborazione"></param>
    ''' <param name="Operatore"></param>
    ''' <returns></returns>
    Public Function SetSituazioneFinale(myConnectionString As String, IdEnte As String, ByVal objICI As objSituazioneFinale(), ByVal nIDElaborazione As Long, Operatore As String) As Long Implements ComPlusInterface.IFreezer.SetSituazioneFinale
        Try
            Return New ClsDBManager().Set_SITUAZIONE_FINALE(myConnectionString, IdEnte, objICI, nIDElaborazione, Operatore)
        Catch ex As Exception
            Throw New Exception(ex.Message & "::" & ex.StackTrace)
        End Try
    End Function
    ''' <summary>
    ''' Interfaccia per la visualizzazione della progressione del calcolo massivo   ''' 
    ''' </summary>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <returns></returns>
    Public Function ViewCodaCalcoloICIMassivo(StringConnectionICI As String, IdEnte As String) As String Implements ComPlusInterface.IFreezer.ViewCodaCalcoloICIMassivo
        Try
            Dim objCOMPLUSBaseCalcoloICI As New ClsFreezer
            Return objCOMPLUSBaseCalcoloICI.ViewCodaCalcolo(StringConnectionICI, IdEnte)
        Catch ex As Exception
            Throw New Exception(ex.Message & "::" & ex.StackTrace)
        End Try
    End Function
    ''' <summary>
    ''' Interfaccia per la visualizzazione dei calcoli massivi pregressi    
    ''' </summary>
    ''' <param name="StringConnectionICI"></param>
    ''' <param name="IdEnte"></param>
    ''' <returns></returns>
    Public Function getDATI_TASK_REPOSITORY_CALCOLO_ICI(StringConnectionICI As String, IdEnte As String) As DataSet Implements IFreezer.getDATI_TASK_REPOSITORY_CALCOLO_ICI
        Dim objCOMPLUSBaseCalcoloICI As New ClsFreezer
        Return objCOMPLUSBaseCalcoloICI.ViewCalcolo(StringConnectionICI, IdEnte)
    End Function
End Class
