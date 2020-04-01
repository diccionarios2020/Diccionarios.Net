'--Imports Diccionarios.My
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.ApplicationServices
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.IO
Imports System.Resources
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

<OptionText()>
Public Class frmPrincipal
    Inherits Form

    ' Ctrl-M, O to collapse all regions; Ctrl-M, P to open them all

    ''' <summary>
    ''' Private variables
    ''' </summary>
    Private config As Configuracion
    Private cargadorDeDics As cargadorDeDiccionarios
    Private diccionarios As Collection  '-- TODO add a type for this
    Private m_PanStartPoint As Point
    Private vZoomActual As Single
    Private vAnchoDePagina As Boolean
    Private vAutoCentrado As Boolean
    Private vID_GrupoActual As Integer
    Private PrintPageSettings As PageSettings

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        MyBase.New()
        Me.config = New Configuracion()
        Me.cargadorDeDics = New cargadorDeDiccionarios()
        Me.diccionarios = New Collection()
        Me.m_PanStartPoint = New Point(0, 0)
        Me.vZoomActual = 100.0!
        Me.vAnchoDePagina = True
        Me.vAutoCentrado = False
        Me.vID_GrupoActual = 1
        Me.PrintPageSettings = New PageSettings()
        Me.InitializeComponent()
    End Sub

#Region "frmPrincipal events - load, resize, etc"

    Private Sub frmPrincipal_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Me.cargarConfiguracion()
        Dim cargadorDeDiccionario As cargadorDeDiccionarios = Me.cargadorDeDics
        cargadorDeDiccionario.conexion = Me.config.conexion
        Dim vIDGrupoActual As Integer = Me.vID_GrupoActual
        Dim comboBox As ComboBox = Me.cbDiccionario
        cargadorDeDiccionario.cargarDiccionarios(vIDGrupoActual, comboBox, Me.diccionarios)
        Me.cbDiccionario = comboBox
        Dim [integer] As Integer = Conversions.ToInteger(Me.cbDiccionario.SelectedValue)
        Dim dataGridView As DataGridView = Me.dgvIndice
        cargadorDeDiccionario.cargarHojasDeUnDiccionario([integer], dataGridView)
        Me.dgvIndice = dataGridView
        Me.lblDatosBibliograficos.Text = Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).referencia
        Me.tabDiccionarioActual.Text = Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).ruta
        cargadorDeDiccionario = Nothing
        Me.panelImagen.AutoScroll = True
        Me.pb1.SizeMode = PictureBoxSizeMode.AutoSize

        '-- The .ico file needs to be specified with Properties->Build Type->Embedded Resource or it will not be found.
        Me.pb1.Cursor = New Cursor(Me.GetType(), "Diccionarios.Manito.ico")

        '-- If panelImagen is set to Autosize=true, this will cause it to grow hugely.  Set it to false
        Me.zoomAlAncho()
        Me.tstbZoom.Text = Strings.Format(Me.vZoomActual, "#.00")
        Me.darFormatoDGVIndice()

    End Sub

    Private Sub frmPrincipal_MouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel

        ' Maneja el scroll vertical si no se presiona otra tecla de control
        If (e.Delta < 0) Then
            Me.zoomImagen(Me.zoomAnterior(Me.vZoomActual))
        End If
        If (e.Delta > 0) Then
            Me.zoomImagen(Me.zoomPosterior(Me.vZoomActual))
        End If

        ' Maneja el scroll horizontal si se manteiene apretada la tecla SHIFT

        If (Control.ModifierKeys = Keys.Shift OrElse (e.Delta < 0 OrElse e.Delta > 0)) Then
        End If

        ' Maneja el zoom si se mantiene apretada la letra CTRL
        If (Control.ModifierKeys = Keys.Control) Then
            If (e.Delta < 0) Then
                Me.zoomImagen(Me.zoomAnterior(Me.vZoomActual))
            End If
            If (e.Delta > 0) Then
                Me.zoomImagen(Me.zoomPosterior(Me.vZoomActual))
            End If
        End If

    End Sub

    Private Sub frmPrincipal_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        If (Me.vAnchoDePagina) Then
            Me.zoomAlAncho()
        End If
    End Sub
#End Region

    '------------------------------------------------------------------------------------------------------------

#Region "Utility  methods"

    ''' <summary>
    ''' Utility method - Show help
    ''' </summary>
    Private Sub cargarAyuda(Optional ByRef key As String = "")
        Help.ShowHelp(Me, Me.hpAyuda.HelpNamespace, key)
    End Sub

    ''' <summary>
    ''' Utility method - Load configuration
    ''' </summary>
    Private Sub cargarConfiguracion()

        Dim configuracion As Configuracion = Me.config

        '-- The Real code was this
        'If (Not Directory.Exists(String.Concat(My.Application.Info.DirectoryPath, "\Diccionarios\"))) Then
        'MsgBox(String.Concat("No se encontró la carpeta 'Diccionarios' dentro de la carpeta: ", My.Application.Info.DirectoryPath), MsgBoxStyle.OkOnly, "Diccionarios no encontrados")
        'Me.Dispose()
        'Else
        'configuracion.dirRepositorio = String.Concat(My.Application.Info.DirectoryPath, "\Diccionarios\")
        'End If

        '-- But this fix from an earlier version allows dev to work without copying stuff into Release and Debug dirs
        If Directory.Exists(My.Application.Info.DirectoryPath & "\Diccionarios\") Then
            configuracion.dirRepositorio = My.Application.Info.DirectoryPath & "\Diccionarios\"
        ElseIf Directory.Exists("D:\Proyecto Libros\Diccionarios\") Then
            configuracion.dirRepositorio = "D:\Proyecto Libros\Diccionarios\"
        Else
            configuracion.dirRepositorio = "D:\Proyecto Libros\Release\Diccionarios\"
        End If

        configuracion.dirMdb = "diccionarios.mdb"
        configuracion.dirInstalacion = String.Concat(My.Application.Info.DirectoryPath, "\")
        configuracion.conexion = String.Concat("Provider=Microsoft.Jet.OLEDB.4.0; Data Source =", configuracion.dirInstalacion, configuracion.dirMdb)
        configuracion = Nothing
    End Sub

    ''' <summary>
    ''' Utility method - Load dictionary
    ''' </summary>
    Private Sub cargarDiccionario()

        Dim buscador As Buscador = New Buscador()
        Dim selVal As Integer = Conversions.ToInteger(Me.cbDiccionario.SelectedValue)

        Me.cargadorDeDics.cargarHojasDeUnDiccionario(selVal, Me.dgvIndice)

        Me.lblDatosBibliograficos.Text = Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).referencia
        Me.tabDiccionarioActual.Text = Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).ruta

        buscador.buscarPalabra(Me.tbBuscar.Text, Me.dgvIndice)
    End Sub

    ''' <summary>
    ''' Utility method - Load the page
    ''' </summary>
    Private Sub cargarPagina()

        Dim rowIndex As Integer

        Dim pagina As Paginas = New Paginas()

        panelImagen.AutoScrollPosition = New Point(0, 0)

        ' Trata de cargar la imagen de la página
        ' Try to load the page image
        Try
            rowIndex = Me.dgvIndice.SelectedCells(0).RowIndex

            pagina.dirRepositorio = Me.config.dirRepositorio
            pagina.nombreDirDiccionario = Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).ruta
            pagina.nombreArchivoImagen = Conversions.ToString(Me.dgvIndice("Archivo", rowIndex).Value)
            Me.pb1.Image = New Bitmap(pagina.urlImagen)

            Me.lblNumeroDePagina.Text = Conversions.ToString(Me.dgvIndice.SelectedCells(1).Value)

        Catch exception As System.Exception

            Debug.Print("cargarPagina:" & exception.Message)
            ProjectData.SetProjectError(exception)
            rowIndex = 0

            '-- Display error image "image not available" / "imagen non diponable"
            Me.pb1.Image = New Bitmap(String.Concat(My.Application.Info.DirectoryPath, "\ind.png"))
            ProjectData.ClearProjectError()
        End Try

        ' Mueve la imagen al vértice superior izquiero
        ' Move the image to the top left corner
        panelImagen.AutoScrollPosition = New Point(0, 0)
    End Sub

    ''' <summary>
    ''' Utility - Convert keystrokes to Greek if language is Greek
    ''' </summary>
    Private Function castGriego(ByVal unaLetra As Char) As Char

        Dim chr As Char
        Select Case unaLetra
            Case "A"c
            Case "a"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B1"))
                Exit Select
            Case "B"c
            Case "b"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B2"))
                Exit Select
            Case "C"c
            Case "c"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C7"))
                Exit Select
            Case "D"c
            Case "d"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B4"))
                Exit Select
            Case "E"c
            Case "e"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B5"))
                Exit Select
            Case "F"c
            Case "f"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C6"))
                Exit Select
            Case "G"c
            Case "g"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B3"))
                Exit Select
            Case "H"c
            Case "h"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B7"))
                Exit Select
            Case "I"c
            Case "i"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B9"))
                Exit Select
            Case "J"c
            Case "j"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C3"))
                Exit Select
            Case "K"c
            Case "k"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03BA"))
                Exit Select
            Case "L"c
            Case "l"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03BB"))
                Exit Select
            Case "M"c
            Case "m"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03BC"))
                Exit Select
            Case "N"c
            Case "n"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03BD"))
                Exit Select
            Case "O"c
            Case "o"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03BF"))
                Exit Select
            Case "P"c
            Case "p"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C0"))
                Exit Select
            Case "Q"c
            Case "q"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B8"))
                Exit Select
            Case "R"c
            Case "r"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C1"))
                Exit Select
            Case "S"c
            Case "s"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C3"))
                Exit Select
            Case "T"c
            Case "t"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C4"))
                Exit Select
            Case "U"c
            Case "u"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C5"))
                Exit Select
            Case "V"c
            Case "["c
            Case "\"c
            Case "]"c
            Case Strings.ChrW(94)
            Case Strings.ChrW(95)
            Case Strings.ChrW(96)
            Case "v"c
Label0:
                chr = Strings.ChrW(0)
                Exit Select
            Case "W"c
            Case "w"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C9"))
                Exit Select
            Case "X"c
            Case "x"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03BE"))
                Exit Select
            Case "Y"c
            Case "y"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03C8"))
                Exit Select
            Case "Z"c
            Case "z"c
                chr = Strings.ChrW(Conversions.ToInteger("&H03B6"))
                Exit Select
            Case Else
                GoTo Label0
        End Select
        Return chr
    End Function

    ''' <summary>
    ''' Utility - Filter non-ascii characters for Latin
    ''' </summary>
    Private Function castLatin(ByVal unaLetra As Char) As Char
        Return If(unaLetra <> Strings.ChrW(0), unaLetra, Strings.ChrW(0))
    End Function

    Private Sub centrarImagen()
        Dim num As Integer = CInt(Math.Round(CDbl(Me.panelImagen.Height) / 2 - CDbl(Me.pb1.Height) / 2))
        Dim point As Point = New Point()
        Dim num1 As Integer = CInt(Math.Round(CDbl(Me.panelImagen.Width) / 2 - 20 - CDbl(Me.pb1.Width) / 2))
        If (num1 <= 0) Then
            point.X = 0
        Else
            point.X = num1
        End If
        If (num <= 0) Then
            point.Y = 0
        Else
            point.Y = num
        End If
        Me.pb1.Location = point
        Me.Text = String.Concat("DeltaX = ", Conversions.ToString(num1), " ; DeltaY = ", Conversions.ToString(num))
    End Sub

    Private Sub centrarImagenA(ByVal x As Integer, ByVal y As Integer)
        Me.panelImagen.AutoScrollPosition = New Point(x, y)
    End Sub

    ''' <summary>
    ''' Utility - copy image to the clipboard
    ''' </summary>
    Private Sub copiarImagenAlPortapapeles()
        Clipboard.SetDataObject(Me.pb1.Image, True)
    End Sub

    ''' <summary>
    ''' Utility - copy bibliographic reference
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub copiarReferenciaBibliográfica()
        Clipboard.SetText(Me.lblDatosBibliograficos.Text)
    End Sub

    ''' <summary>
    ''' Load - set up the DataGridView index column at the left
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub darFormatoDGVIndice()
        Me.dgvIndice.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Me.dgvIndice.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
        Me.dgvIndice.Columns(2).Visible = False
        Me.dgvIndice.Columns(3).Visible = False
        Me.dgvIndice.Columns(4).Visible = False
        Me.dgvIndice.Columns(0).HeaderText = "Encabezado"
        Me.dgvIndice.Columns(1).HeaderText = "Pág."
        Me.dgvIndice.Columns(0).DefaultCellStyle.Font = New Font("Times New Roman", 10.0!, FontStyle.Regular)
        Me.dgvIndice.Columns(1).DefaultCellStyle.Font = New Font("Times New Roman", 10.0!, FontStyle.Regular)
        Me.cbDiccionario.Font = New Font("Times New Roman", 10.0!, FontStyle.Regular)
    End Sub

    ''' <summary>
    ''' Utility - Return a dictionary by ID
    ''' </summary>
    Private Function devolverDiccionarioPorID(ByVal unaID_Libro As Integer) As Diccionario

        Dim enumerator As IEnumerator = Nothing
        Dim diccionario As Diccionario = New Diccionario()
        Try
            enumerator = Me.diccionarios.GetEnumerator()
            While enumerator.MoveNext()
                Dim objectValue As Object = RuntimeHelpers.GetObjectValue(enumerator.Current)
                If (Not Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, Nothing, "id_libro", New Object(-1) {}, Nothing, Nothing, Nothing), unaID_Libro, True)) Then
                    Continue While
                End If
                diccionario = DirectCast(objectValue, Diccionario)
                Return diccionario
            End While
        Finally
            If (TypeOf enumerator Is IDisposable) Then
                TryCast(enumerator, IDisposable).Dispose()
            End If
        End Try
        Return diccionario
    End Function

    ''' <summary>
    ''' Utility - Save As... - save a page to an image file
    ''' </summary>
    Private Sub guardarComo()

        Me.guardarComoDialog.Filter = "Imagen Jpeg(*.jpg)|*.jpg|Imagen Bmp(*.bmp)|*.bmp|Imagen Gif(*.gif)|*.gif|Imagen Tiff(*.tif)|*.tif"
        Me.guardarComoDialog.Title = "Seleccione donde quiere guardar la imagen de la página actual"
        Me.guardarComoDialog.InitialDirectory = String.Concat("C:\Documents and Settings\", My.User.Name, "\Escritorio")
        Me.guardarComoDialog.RestoreDirectory = True
        Me.guardarComoDialog.FileName = ""

        'Se utiliza el save file y se especifica el tipo de dato a guardar

        If Me.guardarComoDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then ' 'Si el pulsamos aceptar en la ventanita

            If Operators.CompareString(Me.guardarComoDialog.FileName, "", True) <> 0 Then

                'Si la ruta del archivo del OpenFileDialog es diferente a nada, es decir, si tiene un nombre 
                'será que hemos cargado una foto, de lo contrario nos dejaría guardar una foto que realmente no tenemos.
                If (Me.guardarComoDialog.FilterIndex = 1) Then
                    Me.pb1.Image.Save(Me.guardarComoDialog.FileName.ToString(), ImageFormat.Jpeg)
                    'MsgBox("jpeg")
                ElseIf (Me.guardarComoDialog.FilterIndex = 2) Then
                    Me.pb1.Image.Save(Me.guardarComoDialog.FileName.ToString(), ImageFormat.Bmp)
                    'MsgBox("bmp")
                ElseIf (Me.guardarComoDialog.FilterIndex = 3) Then
                    Me.pb1.Image.Save(Me.guardarComoDialog.FileName.ToString(), ImageFormat.Gif)
                    'MsgBox("gif")
                ElseIf (Me.guardarComoDialog.FilterIndex = 4) Then
                    Me.pb1.Image.Save(Me.guardarComoDialog.FileName.ToString(), ImageFormat.Tiff)
                    'MsgBox("tiff")
                    'Con los demás es exactamente igual, pero cambiando el formato.
                End If
            End If
        End If

    End Sub

    ''' <summary>
    ''' Utility - Show / Hide the left index pane
    ''' </summary>
    Private Sub ocultarMostrarIndice()

        Dim splitContainer As System.Windows.Forms.SplitContainer = Me.scDiccionario
        If (Not splitContainer.Panel1Collapsed) Then
            splitContainer.Panel1Collapsed = True
        Else
            splitContainer.Panel1Collapsed = False
        End If
        splitContainer = Nothing

        If (Me.vAnchoDePagina) Then
            Me.zoomAlAncho()
        End If
    End Sub

    Private Sub PrintGraphic(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        ev.Graphics.DrawImage(Me.pb1.Image, ev.Graphics.VisibleClipBounds)
        ev.HasMorePages = False
    End Sub

    ''' <summary>
    ''' Utility - Display the Group Dialog
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub seleccionarGrupo()
        frmGrupos.conexion = Me.config.conexion
        frmGrupos.idGrupoActual = Me.vID_GrupoActual
        frmGrupos.ShowDialog()
        If (frmGrupos.hayCambios) Then
            Me.tbBuscar.Clear()
            If (frmGrupos.idGrupoActual <> 0) Then
                Me.vID_GrupoActual = frmGrupos.idGrupoActual
            End If
            Dim cargadorDeDiccionario As cargadorDeDiccionarios = Me.cargadorDeDics
            Dim vIDGrupoActual As Integer = Me.vID_GrupoActual
            Dim comboBox As System.Windows.Forms.ComboBox = Me.cbDiccionario
            cargadorDeDiccionario.cargarDiccionarios(vIDGrupoActual, comboBox, Me.diccionarios)
            Me.cbDiccionario = comboBox
            Dim cargadorDeDiccionario1 As cargadorDeDiccionarios = Me.cargadorDeDics
            Dim [integer] As Integer = Conversions.ToInteger(Me.cbDiccionario.SelectedValue)
            Dim dataGridView As System.Windows.Forms.DataGridView = Me.dgvIndice
            cargadorDeDiccionario1.cargarHojasDeUnDiccionario([integer], dataGridView)
            Me.dgvIndice = dataGridView
            Me.lblDatosBibliograficos.Text = Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).referencia
            Me.tabDiccionarioActual.Text = Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).ruta
            Me.zoomAlAncho()
        End If
        frmGrupos.Dispose()
    End Sub

    ''' <summary>
    ''' Utility - Print
    ''' </summary>
    Private Sub imprimir()
        Try
            Me.imprimirImagen.DefaultPageSettings = Me.PrintPageSettings
            Me.PrintDialog1.Document = Me.imprimirImagen
            If (Me.PrintDialog1.ShowDialog() = DialogResult.OK) Then
                Dim _frmPrincipal As frmPrincipal = Me
                AddHandler Me.imprimirImagen.PrintPage, New PrintPageEventHandler(AddressOf Me.PrintGraphic)
                'imprimirImagen.PrinterSettings.PrinterName = "Adobe PDF"
                Me.imprimirImagen.Print()
            End If
        Catch exception As System.Exception
            ProjectData.SetProjectError(exception)
            MessageBox.Show(exception.Message)
            ProjectData.ClearProjectError()
        End Try
    End Sub

#End Region

#Region "Utility - zoom methods"

    Private Sub tstbZoom_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        If (e.KeyChar = Strings.ChrW(13)) Then
            e.Handled = True
            If (Conversions.ToDouble(Me.tstbZoom.Text) > 0 And Conversions.ToDouble(Me.tstbZoom.Text) < 201) Then
                Me.vZoomActual = Conversions.ToSingle(Me.tstbZoom.Text)
                Me.zoomImagen(Me.vZoomActual)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Utility - zoom to width
    ''' </summary>
    Private Sub zoomAlAncho()
        Try
            Dim width As Single = CSng((CDbl(((Me.panelImagen.Width - 20) * 100)) / CDbl(Me.pb1.Image.Width)))
            Me.zoomImagen(width)
            Me.vAnchoDePagina = True
        Catch exception As System.Exception
            ProjectData.SetProjectError(exception)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    ''' <summary>
    ''' Zoom - make smaller
    ''' </summary>
    Private Function zoomAnterior(ByVal actual As Single) As Single
        Dim [single] As Single = 0.0!
        If (actual >= 1.0! And CDbl(actual) <= 6.25) Then
            [single] = 1.0!
        ElseIf (CDbl(actual) > 6.25 And CDbl(actual) <= 12.5) Then
            [single] = 6.25!
        ElseIf (CDbl(actual) > 12.5 And actual <= 25.0!) Then
            [single] = 12.5!
        ElseIf (actual > 25.0! And CDbl(actual) <= 33.331) Then
            [single] = 25.0!
        ElseIf (CDbl(actual) > 33.33 And actual <= 50.0!) Then
            [single] = 33.33!
        ElseIf (actual > 50.0! And actual <= 75.0!) Then
            [single] = 50.0!
        ElseIf (actual > 75.0! And actual <= 100.0!) Then
            [single] = 75.0!
        ElseIf (actual > 100.0! And actual <= 125.0!) Then
            [single] = 100.0!
        ElseIf (actual > 125.0! And actual <= 150.0!) Then
            [single] = 125.0!
        ElseIf (actual > 150.0! And actual <= 200.0!) Then
            [single] = 150.0!
        End If
        Return [single]
    End Function

    ''' <summary>
    ''' Utility - Zoom the image
    ''' </summary>
    Private Sub zoomImagen(Optional ByVal zoom As Single = 100.0!)
        Me.vZoomActual = zoom
        Dim pictureBox As System.Windows.Forms.PictureBox = Me.pb1
        pictureBox.SizeMode = PictureBoxSizeMode.StretchImage
        pictureBox.Width = CInt(Math.Round(CDbl(CSng((CSng(pictureBox.Image.Width) * (zoom / 100.0!))))))
        pictureBox.Height = CInt(Math.Round(CDbl(CSng((CSng(pictureBox.Image.Height) * (zoom / 100.0!))))))
        Me.tstbZoom.Text = Strings.Format(Me.vZoomActual, "#.00")
        pictureBox = Nothing
        If (Me.vAutoCentrado) Then
            Me.centrarImagen()
        End If
        Me.vAnchoDePagina = False
    End Sub

    ''' <summary>
    ''' Utility - zoom - show whole page
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub zoomPaginaCompleta()
        Me.zoomImagen(CSng((CDbl(((Me.panelImagen.Height - 24) * 100)) / CDbl(Me.pb1.Image.Height))))
    End Sub

    ''' <summary>
    ''' Zoom - zoom bigger
    ''' </summary>
    ''' <param name="actual"></param>
    Private Function zoomPosterior(ByVal actual As Single) As Single
        Dim [single] As Single = 0.0!
        If (actual >= 1.0! And CDbl(actual) < 6.25) Then
            [single] = 6.25!
        ElseIf (CDbl(actual) >= 6.25 And CDbl(actual) < 12.5) Then
            [single] = 12.5!
        ElseIf (CDbl(actual) >= 12.5 And actual < 25.0!) Then
            [single] = 25.0!
        ElseIf (actual >= 25.0! And CDbl(actual) < 33.33) Then
            [single] = 33.33!
        ElseIf (CDbl(actual) >= 33.33 And actual < 50.0!) Then
            [single] = 50.0!
        ElseIf (actual >= 50.0! And actual < 75.0!) Then
            [single] = 75.0!
        ElseIf (actual >= 75.0! And actual < 100.0!) Then
            [single] = 100.0!
        ElseIf (actual >= 100.0! And actual < 125.0!) Then
            [single] = 125.0!
        ElseIf (actual >= 125.0! And actual < 150.0!) Then
            [single] = 150.0!
        ElseIf (actual >= 150.0!) Then
            [single] = 200.0!
        End If
        Return [single]
    End Function
#End Region

    '------------------------------------------------------------------------------------------------------------

#Region "Main Menu Items"

    ''' <summary>
    ''' Menu - Choose language group
    ''' </summary>
    Private Sub SeleccionarGrupo_Click(sender As System.Object, e As System.EventArgs) Handles seleccionarGrupoToolStripMenuItem.Click
        seleccionarGrupo()
    End Sub

    ''' <summary>
    ''' Menu - Exit
    ''' </summary>
    Private Sub SalirToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Dispose()
    End Sub

    ''' <summary>
    ''' Menu - Save As...
    ''' </summary>
    Private Sub GuardarComoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GuardarComoToolStripMenuItem.Click
        guardarComo()
    End Sub

    ''' <summary>
    ''' Menu - Print
    ''' </summary>
    Private Sub ImprimirToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ImprimirToolStripMenuItem.Click
        imprimir()
    End Sub

    ''' <summary>
    ''' Menu - Page Setup
    ''' </summary>
    Private Sub ConfigurarPáginaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ConfigurarPáginaToolStripMenuItem.Click
        Try
            Me.PageSetupDialog1.PageSettings = Me.PrintPageSettings
            Me.PageSetupDialog1.ShowDialog()
        Catch exception As System.Exception
            ProjectData.SetProjectError(exception)
            MessageBox.Show(exception.Message)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    ''' <summary>
    ''' Menu - Print preview
    ''' </summary>
    Private Sub VistaPreviaDeImpresiónToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VistaPreviaDeImpresiónToolStripMenuItem.Click
        Try
            Me.imprimirImagen.DefaultPageSettings = Me.PrintPageSettings
            Dim _frmPrincipal As frmPrincipal = Me
            AddHandler Me.imprimirImagen.PrintPage, New PrintPageEventHandler(AddressOf _frmPrincipal.PrintGraphic)
            Me.PrintPreviewDialog1.Document = Me.imprimirImagen
            Me.PrintPreviewDialog1.ShowDialog()
        Catch exception As System.Exception
            ProjectData.SetProjectError(exception)
            MessageBox.Show(exception.Message)
            ProjectData.ClearProjectError()
        End Try

    End Sub

    ''' <summary>
    ''' Menu - Copy Page
    ''' </summary>
    Private Sub CopiarPáginaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopiarPáginaToolStripMenuItem.Click
        Me.copiarImagenAlPortapapeles()
    End Sub

    ''' <summary>
    ''' Menu - Copy bibliographic reference
    ''' </summary>
    Private Sub CopiarDatosBibliográficosToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopiarDatosBibliográficosToolStripMenuItem.Click
        Me.copiarReferenciaBibliográfica()
    End Sub

    ''' <summary>
    ''' Menu - zoom out.  Shortcut key is Ctrl-.  Set the key to Ctrl+OemMinus, 
    ''' then set the ShortcutKeyDisplayString property to override ugly display
    ''' </summary>
    Private Sub AlejarToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AlejarToolStripMenuItem.Click
        Me.zoomImagen(Me.zoomAnterior(Me.vZoomActual))
    End Sub

    ''' <summary>
    ''' Menu - zoom in.    Shortcut key is Ctrl+.  Set the key to Ctrl+OemPlus, 
    ''' then set the ShortcutKeyDisplayString property to override ugly display
    ''' </summary>
    Private Sub AcercarToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AcercarToolStripMenuItem.Click
        Me.zoomImagen(Me.zoomPosterior(Me.vZoomActual))
    End Sub

    ''' <summary>
    ''' Menu - zoom 25%
    ''' </summary>
    Private Sub Perc25ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc25ToolStripMenuItem.Click
        Me.zoomImagen(25.0!)
    End Sub

    ''' <summary>
    ''' Menu - zoom 50%
    ''' </summary>
    Private Sub Perc50ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc50ToolStripMenuItem.Click
        Me.zoomImagen(50.0!)
    End Sub

    ''' <summary>
    ''' Menu - zoom 75%
    ''' </summary>
    Private Sub Perc75ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc75ToolStripMenuItem.Click
        Me.zoomImagen(75.0!)
    End Sub

    ''' <summary>
    ''' Menu - zoom 100%
    ''' </summary>
    Private Sub Perc100ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc100ToolStripMenuItem.Click
        Me.zoomImagen(100.0!)
    End Sub

    ''' <summary>
    ''' Menu - zoom 150%
    ''' </summary>
    Private Sub Perc150ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc150ToolStripMenuItem.Click
        Me.zoomImagen(150.0!)
    End Sub

    ''' <summary>
    ''' Menu - zoom 200%
    ''' </summary>
    Private Sub Perc200ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc200ToolStripMenuItem.Click
        Me.zoomImagen(200.0!)
    End Sub

    ''' <summary>
    ''' Menu - zoom 300%
    ''' </summary>
    Private Sub Perc300ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc300ToolStripMenuItem.Click
        Me.zoomImagen(300.0!)
    End Sub

    ''' <summary>
    ''' Menu - real size = 100%
    ''' </summary>
    Private Sub VerTamañoRealToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VerTamañoRealToolStripMenuItem.Click
        Me.zoomImagen(100.0!)
    End Sub

    ''' <summary>
    ''' Menu - zoom page to width
    ''' </summary>
    Private Sub VerAlAnchoDePáginaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VerAlAnchoDePáginaToolStripMenuItem.Click
        Me.zoomAlAncho()
    End Sub

    ''' <summary>
    ''' Menu - view whole page
    ''' </summary>
    Private Sub VerPáginaCompletaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VerPáginaCompletaToolStripMenuItem.Click
        Me.zoomPaginaCompleta()
    End Sub

    ''' <summary>
    ''' Menu - Auto centre.  Greyed out in original, so probably does not work
    ''' </summary>
    Private Sub AutocentradoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AutocentradoToolStripMenuItem.Click

        If (Not Me.AutocentradoToolStripMenuItem.Checked) Then
            Me.AutocentradoToolStripMenuItem.Checked = True
            Me.vAutoCentrado = True
        Else
            Me.AutocentradoToolStripMenuItem.Checked = False
            Me.vAutoCentrado = False
        End If

    End Sub

    ''' <summary>
    ''' Menu - Show / Hide list of pages
    ''' </summary>
    Private Sub OccultaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OccultaToolStripMenuItem.Click
        Me.ocultarMostrarIndice()
    End Sub

    ''' <summary>
    ''' Menu - Make the app the topmost app
    ''' </summary>
    Private Sub SiempreVisibleToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles SiempreVisibleToolStripMenuItem.Click
        If (Not Me.SiempreVisibleToolStripMenuItem.Checked) Then
            Me.SiempreVisibleToolStripMenuItem.Checked = True
            Me.TopMost = True
        Else
            Me.SiempreVisibleToolStripMenuItem.Checked = False
            Me.TopMost = False
        End If
    End Sub

    ''' <summary>
    ''' Menu - Show Greek keyboard image
    ''' </summary>
    Private Sub TecladaGriegoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TecladaGriegoToolStripMenuItem.Click
        Dim form As Form = New Form()
        Dim size As Form = form
        size.Name = "frmTeclado"
        size.FormBorderStyle = FormBorderStyle.FixedToolWindow
        size.Size = New System.Drawing.Size(606, 224)
        size.StartPosition = FormStartPosition.CenterScreen
        size.BackgroundImage = New Bitmap(Me.[GetType](), "keyboard.png")
        size = Nothing
        form.ShowDialog()
    End Sub

    ''' <summary>
    ''' Menu - Show Options - not enabled
    ''' </summary>
    Private Sub OpcionesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OpcionesToolStripMenuItem.Click
        frmOpciones.ShowDialog()
    End Sub

    ''' <summary>
    ''' Menu - Show help
    ''' </summary>
    Private Sub ÍndiceToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles ÍndiceToolStripMenuItem.Click
        Dim str As String = ""
        Me.cargarAyuda(str)
        'System.Windows.Forms.Help.ShowHelp(Me, "diccionarios.chm")
    End Sub

    ''' <summary>
    ''' Menu - Show about box
    ''' </summary>
    Private Sub AcercaDeToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles AcercaDeToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub
#End Region

#Region "Toolbar Buttons"

    '-- Toolbar - find group
    Private Sub tsbSeleccionarGrupo_Click(sender As System.Object, e As System.EventArgs) Handles tsbSeleccionarGrupo.Click
        Me.seleccionarGrupo()
    End Sub

    '-- Toolbar - Save
    Private Sub tsbGuardarComo_Click(sender As System.Object, e As System.EventArgs) Handles tsbGuardarComo.Click
        Me.guardarComo()
    End Sub

    '-- Toolbar - print
    Private Sub tsbImprimir_Click(sender As System.Object, e As System.EventArgs) Handles tsbImprimir.Click
        Me.imprimir()
    End Sub

    '-- Toolbar - print preview
    Private Sub tsbPreview_Click(sender As System.Object, e As System.EventArgs) Handles tsbPreview.Click
        Try
            imprimirImagen.DefaultPageSettings = PrintPageSettings
            AddHandler imprimirImagen.PrintPage, AddressOf Me.PrintGraphic
            PrintPreviewDialog1.Document = imprimirImagen
            PrintPreviewDialog1.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    '-- Toolbar - select something
    Private Sub tsbSeleccionar_Click(sender As System.Object, e As System.EventArgs) Handles tsbSeleccionar.Click
        MsgBox("Unimplemented")
    End Sub

    '-- Toolbar - deselect something
    Private Sub tsbDeseleccionar_Click(sender As System.Object, e As System.EventArgs) Handles tsbDeseleccionar.Click
        MsgBox("Unimplemented")
    End Sub

    '-- Toolbar - copy
    Private Sub tsbCopiar_Click(sender As System.Object, e As System.EventArgs) Handles tsbCopiar.Click
        Me.copiarImagenAlPortapapeles()
    End Sub

    '-- Toolbar - show/hide sidebar
    Private Sub tbSidebar_Click(sender As System.Object, e As System.EventArgs) Handles tbSidebar.Click
        Me.ocultarMostrarIndice()
    End Sub

    'Tool bar - zoom out  - written by RP
    Private Sub tsbAcercar_Click(sender As System.Object, e As System.EventArgs) Handles tsbAcercar.Click
        Me.zoomImagen(Me.zoomPosterior(Me.vZoomActual))
    End Sub

    'Tool bar - zoom in - written by RP
    Private Sub tsbAlejar_Click(sender As System.Object, e As System.EventArgs) Handles tsbAlejar.Click
        Me.zoomImagen(Me.zoomAnterior(Me.vZoomActual))
    End Sub


    ''' <summary>
    ''' Toolbar - Show help
    ''' </summary>
    Private Sub tsbAyuda_Click(sender As System.Object, e As System.EventArgs) Handles tsbAyuda.Click
        Me.cargarAyuda("")
    End Sub

    ''' <summary>
    ''' Display options dialog - not implemented
    ''' </summary>
    Private Sub tsbOpciones_Click(sender As System.Object, e As System.EventArgs) Handles tsbOpciones.Click
        frmOpciones.ShowDialog()
    End Sub

#End Region

#Region "Toolbar - Zoom text box events"

    ''' <summary>
    ''' Toolbar text box - key press 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub tstbZoom_KeyPress1(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles tstbZoom.KeyPress
        If (e.KeyChar = Strings.ChrW(13)) Then
            e.Handled = True
            If (Conversions.ToDouble(Me.tstbZoom.Text) > 0 And Conversions.ToDouble(Me.tstbZoom.Text) < 601) Then
                Me.zoomImagen(Conversions.ToSingle(Me.tstbZoom.Text))
            End If
        End If
    End Sub

    Private Sub tstbZoom_Leave(sender As Object, e As System.EventArgs) Handles tstbZoom.Leave
        Me.tstbZoom.Text = Strings.Format(Me.vZoomActual, "#.00")
    End Sub

#End Region

#Region "Toolbar - Zoom Drop-down Menu"

    Private Sub zoom25_Click(sender As System.Object, e As System.EventArgs) Handles zoom25.Click
        Me.zoomImagen(25.0!)
    End Sub

    Private Sub zoom50_Click(sender As System.Object, e As System.EventArgs) Handles zoom50.Click
        Me.zoomImagen(50.0!)
    End Sub

    Private Sub zoom75_Click(sender As System.Object, e As System.EventArgs) Handles zoom75.Click
        Me.zoomImagen(75.0!)
    End Sub

    Private Sub zoom100_Click(sender As System.Object, e As System.EventArgs) Handles zoom100.Click
        Me.zoomImagen(100.0!)
    End Sub

    Private Sub zoom150_Click(sender As System.Object, e As System.EventArgs) Handles zoom150.Click
        Me.zoomImagen(150.0!)
    End Sub

    Private Sub zoom200_Click(sender As System.Object, e As System.EventArgs) Handles zoom200.Click
        Me.zoomImagen(200.0!)
    End Sub

    Private Sub zoom300_Click(sender As System.Object, e As System.EventArgs) Handles zoom300.Click
        Me.zoomImagen(300.0!)
    End Sub

    '-- Real size
    Private Sub zoomTamanoReal_Click(sender As System.Object, e As System.EventArgs) Handles zoomTamanoReal.Click
        Me.zoomImagen(100.0!)
    End Sub

    '-- Width of the window
    Private Sub zoomAncho_Click(sender As System.Object, e As System.EventArgs) Handles zoomAncho.Click
        Me.zoomAlAncho()
    End Sub

    '-- View whole page
    Private Sub zoom_PaginaCompleta_Click(sender As System.Object, e As System.EventArgs) Handles zoom_PaginaCompleta.Click
        zoomPaginaCompleta()
    End Sub

#End Region

    '------------------------------------------------------------------------------------------------------------

#Region "Left panel - Combo box drop down to change dictionary"

    ''' <summary>
    ''' Change current dictionary using the combo box drop down
    ''' </summary>
    Private Sub cbDiccionario_SelectionChangeCommitted(sender As Object, e As System.EventArgs) _
        Handles cbDiccionario.SelectionChangeCommitted
        Me.cargarDiccionario()
        Me.dgvIndice.Focus()
    End Sub
#End Region

#Region "Left panel - DataGridView events"

    ''' <summary>
    ''' DataGridView - indice, key handler
    ''' </summary>
    Private Sub dgvIndice_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles dgvIndice.KeyPress


        If (Char.IsLetter(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Not Char.IsSeparator(e.KeyChar)) Then
            e.Handled = True
        Else
            e.Handled = False
        End If
        Me.tbBuscar.Focus()
        Me.tbBuscar.Clear()

        If (e.KeyChar = Strings.ChrW(8)) Then                              ' controla la tecla BACKSPACE
            e.KeyChar = Strings.ChrW(8)
        ElseIf (e.KeyChar = Strings.ChrW(27)) Then                         ' hace que la tecla ESC borre el texto
            Me.tbBuscar.Clear()
        ElseIf (Operators.CompareString(Conversions.ToString(e.KeyChar), "" & vbCrLf & "", True) = 0) Then
            Me.tbBuscar.AutoCompleteCustomSource.Add(Me.tbBuscar.Text)
        ElseIf (Operators.CompareString(Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).idioma, "Latín", True) <> 0) Then
            e.KeyChar = Me.castGriego(e.KeyChar)
        Else
            e.KeyChar = Me.castLatin(e.KeyChar)
        End If

        Me.tbBuscar.AppendText(Conversions.ToString(e.KeyChar))
    End Sub

    '-- List of pages - Selection changed
    Private Sub dgvIndice_SelectionChanged1(sender As Object, e As System.EventArgs) Handles dgvIndice.SelectionChanged
        If (Me.dgvIndice.SelectedRows.Count > 0) Then
            Me.cargarPagina()
        End If
    End Sub
#End Region

#Region "Left panel - Search box"

    ''' <summary>
    ''' Clear the search box using the button, btnBorrar
    ''' </summary>
    Private Sub btnBorrarTBBuscar_Click(sender As System.Object, e As System.EventArgs) Handles btnBorrarTBBuscar.Click
        Me.tbBuscar.Text = ""
    End Sub

    Private Sub tbBuscar_Click(sender As Object, e As System.EventArgs) Handles tbBuscar.Click
        Me.tbBuscar.SelectAll()
    End Sub

    ''' <summary>
    ''' Search for page starting with letters as we type into text box tbBuscar
    ''' </summary>
    Private Sub tbBuscar_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles tbBuscar.KeyPress
        If (Char.IsLetter(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Not Char.IsSeparator(e.KeyChar)) Then
            e.Handled = True
        Else
            e.Handled = True
        End If
        If (e.KeyChar = Strings.ChrW(8)) Then
            e.KeyChar = Strings.ChrW(8)
        ElseIf (e.KeyChar = Strings.ChrW(27)) Then
            Me.tbBuscar.Clear()
        ElseIf (Operators.CompareString(Conversions.ToString(e.KeyChar), "" & vbCrLf & "", True) = 0) Then
            Me.tbBuscar.AutoCompleteCustomSource.Add(Me.tbBuscar.Text)
        ElseIf (Operators.CompareString(Me.devolverDiccionarioPorID(Conversions.ToInteger(Me.cbDiccionario.SelectedValue)).idioma, "Latín", True) <> 0) Then
            e.KeyChar = Me.castGriego(e.KeyChar)
        Else
            e.KeyChar = Me.castLatin(e.KeyChar)
        End If
    End Sub

    Private Sub tbBuscar_TextChanged(sender As Object, e As System.EventArgs) Handles tbBuscar.TextChanged
        Dim buscador As Buscador = New Buscador()
        buscador.buscarPalabra(Me.tbBuscar.Text, Me.dgvIndice)
    End Sub

#End Region

    '------------------------------------------------------------------------------------------------------------

#Region "Middle - Splitter events"

    ''' <summary>
    ''' If the splitter moves, then repaint the right window.  But the splitter does not move, so I think this may be dead code.
    ''' </summary>
    Private Sub scDiccionario_SplitterMoved(sender As Object, e As System.Windows.Forms.SplitterEventArgs) Handles scDiccionario.SplitterMoved
        Me.dgvIndice.Focus()
        If (Me.vAnchoDePagina) Then
            Me.zoomAlAncho()
        End If
    End Sub
#End Region

    '------------------------------------------------------------------------------------------------------------

#Region "Right-panel - PictureBox PB1 events"

    ''' <summary>
    ''' PictureBox1 - mouse click sets focus
    ''' </summary>
    Private Sub pb1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseClick
        Me.pb1.Focus()
    End Sub

    ''' <summary>
    ''' PictureBox1 - mouse down changes the cursor 
    ''' </summary>
    Private Sub pb1_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseDown
        Me.m_PanStartPoint = New Point(e.X, e.Y)
        Me.pb1.Cursor = New Cursor(Me.GetType(), "Diccionarios.Manito2.ico")
    End Sub

    ''' <summary>
    ''' PictureBox1 - dragging the mouse drags the picture around the panel
    ''' </summary>
    Private Sub pb1_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseMove
        If (e.Button = System.Windows.Forms.MouseButtons.Left) Then
            Dim x As Integer = Me.m_PanStartPoint.X - e.X
            Dim y As Integer = Me.m_PanStartPoint.Y - e.Y
            Dim panel As System.Windows.Forms.Panel = Me.panelImagen
            Dim num As Integer = x - Me.panelImagen.AutoScrollPosition.X
            Dim autoScrollPosition As Point = Me.panelImagen.AutoScrollPosition
            Dim point As Point = New Point(num, y - autoScrollPosition.Y)
            panel.AutoScrollPosition = point
        End If

    End Sub

    ''' <summary>
    ''' PictureBox1 - mouse up changes the cursor back to the normal hand
    ''' </summary>
    Private Sub pb1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseUp

        Me.pb1.Cursor = New Cursor(Me.GetType(), "Diccionarios.Manito.ico")
        If (Control.ModifierKeys = Keys.Control) Then
            Dim point As Point = New Point() With
            {
             .X = e.X - 25,
             .Y = e.Y - 25
            }
            Me.centrarImagenA(point.X, point.Y)
        End If

    End Sub

#End Region

#Region "cmsPB1 context menu strip for picture box"

    ''' <summary>
    ''' cmsPB1 menu - Hand tool.  TODO does not seem to do anything
    ''' </summary>
    Private Sub HerramientaManoCmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HerramientaManoCmsPB1ToolStripMenuItem.Click
        MsgBox("Not implemented")
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Copy page to clipboard
    ''' </summary>
    Private Sub CopiarCmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopiarCmsPB1ToolStripMenuItem.Click
        Me.copiarImagenAlPortapapeles()
    End Sub

#End Region

#Region "cmsPB1 context menu strip for picture box - Zoom"

    ''' <summary>
    ''' cmsPB1 menu - Zoom in
    ''' </summary>
    Private Sub AlejarCmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AlejarCmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(Me.zoomAnterior(Me.vZoomActual))
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom out
    ''' </summary>
    Private Sub AcercarCmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AcercarCmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(Me.zoomPosterior(Me.vZoomActual))
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - 25%
    ''' </summary>
    Private Sub Perc25CmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc25CmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(25.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - 50%
    ''' </summary>
    Private Sub Perc50CmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc50CmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(50.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - 75%
    ''' </summary>
    Private Sub Perc75CmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc75CmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(75.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - 100%
    ''' </summary>
    Private Sub Perc100CmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc100CmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(100.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - 150%
    ''' </summary>
    Private Sub Perc150CmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc150CmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(150.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - 200%
    ''' </summary>
    Private Sub Perc200CmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc200CmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(200.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - 300%
    ''' </summary>
    Private Sub Perc300CmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Perc300CmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(300.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - View normal size, i.e. 100%
    ''' </summary>
    Private Sub VerTamañoRealCmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VerTamañoRealCmsPB1ToolStripMenuItem.Click
        Me.zoomImagen(100.0!)
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - Width of page
    ''' </summary>
    Private Sub AnchoDePáginaCmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AnchoDePáginaCmsPB1ToolStripMenuItem.Click
        Me.zoomAlAncho()
    End Sub

    ''' <summary>
    ''' cmsPB1 menu - Zoom - Whole page
    ''' </summary>
    Private Sub VerPáginaCompletaCmsPB1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VerPáginaCompletaCmsPB1ToolStripMenuItem.Click
        zoomPaginaCompleta()
    End Sub

#End Region

    '------------------------------------------------------------------------------------------------------------

#Region "Bottom right - Prev/Next buttons"

    ''' <summary>
    ''' Prev page
    ''' </summary>
    Private Sub btnPagAnterior_Click(sender As System.Object, e As System.EventArgs) Handles btnPagAnterior.Click
        Dim index As Integer = Me.dgvIndice.SelectedRows(0).Index
        If (index > 0) Then
            Me.dgvIndice.Rows(index - 1).Selected = True
        End If
    End Sub

    ''' <summary>
    ''' Next page
    ''' </summary>
    Private Sub btnPagSiguiente_Click(sender As System.Object, e As System.EventArgs) Handles btnPagSiguiente.Click
        Dim index As Integer = Me.dgvIndice.SelectedRows(0).Index
        If (index < Me.dgvIndice.Rows.Count - 1) Then
            Me.dgvIndice.Rows(index + 1).Selected = True
        End If
    End Sub
#End Region

End Class