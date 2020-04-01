Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System

<OptionText()>
 Public Class Configuracion
    Private vDirInstalacion As String

    Private vMdb As String

    Private vDirRepositorioLocal As String

    Private vDirRepositorioLocalHabilitado As Boolean

    Private vDirRepositorioExtraible As String

    Private vDirRepositorioExtraibleHabilitado As Boolean

    Private vDirRepositorioInternet As String

    Private vDirRepositorioInternetHabilitado As Boolean

    Private vDirRepositorio As String

    Private vIdiomaActual As String

    Private vFuenteIndiceDiccionarioDeEntrada As String

    Private vFuenteIndiceDiccionarioDeEntradaTamano As Integer

    Private vFuenteIndiceGrupoDeDiccionarios As String

    Private vFuenteIndiceGrupoDeDiccionariosTamano As Integer

    Private vConexion As String

    Private vDiccionarioPredeterminado As String

    Private vZoomPorcentual As Boolean

    Private vZoomRenderizado As Boolean

    Private vPestanasMaximaCantidad As Integer

    Private vPestanasMaximaCantidadAutomatica As Boolean

    Public Property conexion As String
        Get
            Return Me.vConexion
        End Get
        Set(ByVal value As String)
            Me.vConexion = value
        End Set
    End Property

    Public Property diccionarioPredeterminado As String
        Get
            Return Me.vDiccionarioPredeterminado
        End Get
        Set(ByVal value As String)
            Me.vDiccionarioPredeterminado = value
        End Set
    End Property

    Public Property dirInstalacion As String
        Get
            Return Me.vDirInstalacion
        End Get
        Set(ByVal value As String)
            Me.vDirInstalacion = value
        End Set
    End Property

    Public Property dirMdb As String
        Get
            Return Me.vMdb
        End Get
        Set(ByVal value As String)
            Me.vMdb = value
        End Set
    End Property

    Public Property dirRepositorio As String
        Get
            Return Me.vDirRepositorio
        End Get
        Set(ByVal value As String)
            Me.vDirRepositorio = value
        End Set
    End Property

    Public Property dirRepositorioExtraible As String
        Get
            Return Me.vDirRepositorioExtraible
        End Get
        Set(ByVal value As String)
            Me.vDirRepositorioExtraible = value
        End Set
    End Property

    Public Property dirRepositorioExtraibleHabilitado As Boolean
        Get
            Return Me.vDirRepositorioExtraibleHabilitado
        End Get
        Set(ByVal value As Boolean)
            Me.vDirRepositorioExtraibleHabilitado = value
        End Set
    End Property

    Public Property dirRepositorioInternet As String
        Get
            Return Me.vDirRepositorioInternet
        End Get
        Set(ByVal value As String)
            Me.vDirRepositorioInternet = value
        End Set
    End Property

    Public Property dirRepositorioInternetHabilitado As Boolean
        Get
            Return Me.vDirRepositorioInternetHabilitado
        End Get
        Set(ByVal value As Boolean)
            Me.vDirRepositorioInternetHabilitado = value
        End Set
    End Property

    Public Property dirRepositorioLocal As String
        Get
            Return Me.vDirRepositorioLocal
        End Get
        Set(ByVal value As String)
            Me.vDirRepositorioLocal = value
        End Set
    End Property

    Public Property dirRepositorioLocalHabilitado As Boolean
        Get
            Return Me.vDirRepositorioLocalHabilitado
        End Get
        Set(ByVal value As Boolean)
            Me.vDirRepositorioLocalHabilitado = value
        End Set
    End Property

    Public Property fuenteIndiceDiccionarioDeEntrada As String
        Get
            Return Me.vFuenteIndiceDiccionarioDeEntrada
        End Get
        Set(ByVal value As String)
            Me.vFuenteIndiceDiccionarioDeEntrada = value
        End Set
    End Property

    Public Property fuenteIndiceDiccionarioDeEntradaTamano As Integer
        Get
            Return Me.vFuenteIndiceDiccionarioDeEntradaTamano
        End Get
        Set(ByVal value As Integer)
            Me.vFuenteIndiceDiccionarioDeEntradaTamano = value
        End Set
    End Property

    Public Property fuenteIndiceGrupoDeDiccionarios As String
        Get
            Return Me.vFuenteIndiceGrupoDeDiccionarios
        End Get
        Set(ByVal value As String)
            Me.vFuenteIndiceGrupoDeDiccionarios = value
        End Set
    End Property

    Public Property fuenteIndiceGrupoDeDiccionariosTamano As Integer
        Get
            Return Me.vFuenteIndiceGrupoDeDiccionariosTamano
        End Get
        Set(ByVal value As Integer)
            Me.vFuenteIndiceGrupoDeDiccionariosTamano = value
        End Set
    End Property

    Public Property idiomaActual As String
        Get
            Return Me.vIdiomaActual
        End Get
        Set(ByVal value As String)
            Me.vIdiomaActual = value
        End Set
    End Property

    Public Property pestanasMaximaCantidad As Integer
        Get
            Return Me.vPestanasMaximaCantidad
        End Get
        Set(ByVal value As Integer)
            Me.vPestanasMaximaCantidad = value
        End Set
    End Property

    Public Property pestanasMaximaCantidadAutomatica As Boolean
        Get
            Return Me.vPestanasMaximaCantidadAutomatica
        End Get
        Set(ByVal value As Boolean)
            Me.vPestanasMaximaCantidadAutomatica = value
        End Set
    End Property

    Public Property zoomPorcentual As Boolean
        Get
            Return Me.vZoomPorcentual
        End Get
        Set(ByVal value As Boolean)
            Me.vZoomPorcentual = value
        End Set
    End Property

    Public Property zoomRenderizado As Boolean
        Get
            Return Me.vZoomRenderizado
        End Get
        Set(ByVal value As Boolean)
            Me.vZoomRenderizado = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        Me.vMdb = "diccionarios.mdb"
        Me.vDirRepositorioLocal = String.Concat(Me.vDirInstalacion, "\Diccionarios")
        Me.vDirRepositorioLocalHabilitado = True
        Me.vDirRepositorioExtraibleHabilitado = False
        Me.vDirRepositorioInternetHabilitado = False
        Me.vIdiomaActual = "castellano"
        Me.vFuenteIndiceDiccionarioDeEntrada = "Times New Roman"
        Me.vFuenteIndiceDiccionarioDeEntradaTamano = 10
        Me.vFuenteIndiceGrupoDeDiccionarios = "Times New Roman"
        Me.vFuenteIndiceGrupoDeDiccionariosTamano = 10
        Me.vZoomPorcentual = True
        Me.vZoomRenderizado = False
        Me.vPestanasMaximaCantidad = 6
        Me.vPestanasMaximaCantidadAutomatica = False
    End Sub

    Private Sub cargarIni()
        Dim cIniClass As CIniClass = New CIniClass() With
        {
         .Archivo = "Diccionarios.ini"
        }
        Interaction.MsgBox(cIniClass.LeeIni("Hex Workshop", "Path"), MsgBoxStyle.OkOnly, Nothing)
        cIniClass.GrabaIni("prueba_net", "mensaje", "hola mundo")
    End Sub
End Class
