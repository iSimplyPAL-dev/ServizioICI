Imports System
Imports System.Runtime.Remoting
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Http
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Collections
Imports log4net
Imports log4net.Config
Imports System.IO
Imports System.Configuration
Imports System.ServiceProcess
''' <summary>
''' Classe di iniziazione del servizio.
''' 
''' Il servizio si occupa di calcolare il dovuto IMU/TASI.
''' </summary>
Public Class ServiceMotoreICI
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ServiceMotoreICI))
    'true --> quando si deve buildare il servizio
    'false --> quando si vuole lanciare in console per il debug
    Private Shared _runService As Boolean = True

    Private chan As TcpChannel
    Private httpChan As HttpChannel

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Dim pathfileinfo As String = ConfigurationSettings.AppSettings("pathfileconflog4net").ToString()
        Dim fileconfiglog4net As New FileInfo(pathfileinfo)
        XmlConfigurator.ConfigureAndWatch(fileconfiglog4net)

        RegisterService()
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        ChannelServices.UnregisterChannel(chan)
    End Sub

    Private Shared Sub RegisterService()

        ' Use the configuration file. 
        RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)

        ' Check to see if we have full errors. 

        'string s = "Errore eccezioni"; 
        If RemotingConfiguration.CustomErrorsEnabled(False) = True Then
        End If

        Console.WriteLine("Inizializzazione Servizio Remoto")
        Dim clientProvider As New BinaryClientFormatterSinkProvider
        Dim serverProvider As New BinaryServerFormatterSinkProvider
        serverProvider.TypeFilterLevel = System.Runtime.Serialization.Formatters.TypeFilterLevel.Full
        Dim props As IDictionary = New Hashtable
        props("port") = ConfigurationSettings.AppSettings("TCP_PORT").ToString()
        '50010; 
        'props["typeFilterLevel"] = TypeFilterLevel.Full; 
        Dim chan As New TcpChannel(props, clientProvider, serverProvider)

        props("port") = ConfigurationSettings.AppSettings("HTTP_PORT").ToString()
        ' 50011; 
        Dim clientProviderSoap As New SoapClientFormatterSinkProvider
        Dim serverProviderSoap As New SoapServerFormatterSinkProvider
        serverProviderSoap.TypeFilterLevel = System.Runtime.Serialization.Formatters.TypeFilterLevel.Full


        Dim httpChan As New HttpChannel(props, Nothing, Nothing)

        log.Debug("Registrazione Canale")
        ChannelServices.RegisterChannel(chan)
        ChannelServices.RegisterChannel(httpChan)
        'ChannelServices.RegisterChannel(httpChan2); 

        RemotingConfiguration.RegisterWellKnownServiceType(GetType(ServiziFreezer), "COMPlusFreezer.rem", WellKnownObjectMode.SingleCall)
        log.Debug("Registrato COMPlusFreezer")
    End Sub
End Class
