Imports System.Data
Imports DevExpress.XtraBars
Imports DevExpress.XtraGrid.Views.Base
Imports NTSInformatica.CLN__STD

Public Class FRMZCIMAN

    Public WithEvents tlbRecordCancella As NTSInformatica.NTSBarButtonItem
    Public WithEvents pnTop As NTSPanel
    Public oCleIMAN As CLEZCIMAN
    Public oCallParams As CLE__CLDP
    Public dsIMAN As DataSet
    Public dcIMAN As BindingSource = New BindingSource()

    Private components As System.ComponentModel.IContainer
    Public WithEvents NtsBarManager1 As NTSInformatica.NTSBarManager
    Public WithEvents tlbMain As NTSInformatica.NTSBar
    Public WithEvents tlbNuovo As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbCancella As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbZoom As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbStampa As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbStampaVideo As NTSInformatica.NTSBarButtonItem
    Public WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Public WithEvents tlbGuida As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbEsci As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbStrumenti As NTSInformatica.NTSBarSubItem
    Public WithEvents tlbImpostaStampante As NTSInformatica.NTSBarButtonItem

    Public oCleBoll As CLEVEBOLL


    Public Overridable Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FRMZCIMAN))
        Me.NtsBarManager1 = New NTSInformatica.NTSBarManager()
        Me.tlbMain = New NTSInformatica.NTSBar()
        Me.tlbZoom = New NTSInformatica.NTSBarButtonItem()
        Me.tlbElabora = New NTSInformatica.NTSBarButtonItem()
        Me.tlbStrumenti = New NTSInformatica.NTSBarSubItem()
        Me.tlbImpostaStampante = New NTSInformatica.NTSBarButtonItem()
        Me.tlbGuida = New NTSInformatica.NTSBarButtonItem()
        Me.tlbEsci = New NTSInformatica.NTSBarButtonItem()
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.tlbNuovo = New NTSInformatica.NTSBarButtonItem()
        Me.tlbCancella = New NTSInformatica.NTSBarButtonItem()
        Me.tlbStampa = New NTSInformatica.NTSBarButtonItem()
        Me.tlbStampaVideo = New NTSInformatica.NTSBarButtonItem()
        Me.tlbRecordCancella = New NTSInformatica.NTSBarButtonItem()
        Me.tlbApri = New NTSInformatica.NTSBarButtonItem()
        Me.tlbCancMultipla = New NTSInformatica.NTSBarButtonItem()
        Me.tlbDuplica = New NTSInformatica.NTSBarButtonItem()
        Me.pnTop = New NTSInformatica.NTSPanel()
        Me.ck_maeriali = New NTSInformatica.NTSCheckBox()
        Me.ck_mezzi = New NTSInformatica.NTSCheckBox()
        Me.ck_risorse = New NTSInformatica.NTSCheckBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NtsBarManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnTop, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnTop.SuspendLayout()
        CType(Me.ck_maeriali.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ck_mezzi.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ck_risorse.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NtsBarManager1
        '
        Me.NtsBarManager1.AllowCustomization = False
        Me.NtsBarManager1.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.tlbMain})
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlTop)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlBottom)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlLeft)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlRight)
        Me.NtsBarManager1.Form = Me
        Me.NtsBarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.tlbNuovo, Me.tlbCancella, Me.tlbStampa, Me.tlbStampaVideo, Me.tlbGuida, Me.tlbEsci, Me.tlbZoom, Me.tlbStrumenti, Me.tlbImpostaStampante, Me.tlbRecordCancella, Me.tlbApri, Me.tlbCancMultipla, Me.tlbDuplica, Me.tlbElabora})
        Me.NtsBarManager1.MaxItemId = 23
        '
        'tlbMain
        '
        Me.tlbMain.BarName = "tlbMain"
        Me.tlbMain.DockCol = 0
        Me.tlbMain.DockRow = 0
        Me.tlbMain.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.tlbMain.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, False, Me.tlbZoom, False), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbElabora), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, False, Me.tlbStrumenti, True), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbGuida, True), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbEsci)})
        Me.tlbMain.OptionsBar.AllowQuickCustomization = False
        Me.tlbMain.OptionsBar.DisableClose = True
        Me.tlbMain.OptionsBar.DrawDragBorder = False
        Me.tlbMain.OptionsBar.UseWholeRow = True
        Me.tlbMain.Text = "tlbMain"
        '
        'tlbZoom
        '
        Me.tlbZoom.Caption = "Zoom"
        Me.tlbZoom.Glyph = CType(resources.GetObject("tlbZoom.Glyph"), System.Drawing.Image)
        Me.tlbZoom.Id = 13
        Me.tlbZoom.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F5)
        Me.tlbZoom.Name = "tlbZoom"
        Me.tlbZoom.Visible = False
        '
        'tlbElabora
        '
        Me.tlbElabora.Caption = "Elabora"
        Me.tlbElabora.Glyph = CType(resources.GetObject("tlbElabora.Glyph"), System.Drawing.Image)
        Me.tlbElabora.Id = 22
        Me.tlbElabora.Name = "tlbElabora"
        Me.tlbElabora.Visible = True
        '
        'tlbStrumenti
        '
        Me.tlbStrumenti.Caption = "Strumenti"
        Me.tlbStrumenti.Glyph = CType(resources.GetObject("tlbStrumenti.Glyph"), System.Drawing.Image)
        Me.tlbStrumenti.Id = 15
        Me.tlbStrumenti.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.tlbImpostaStampante)})
        Me.tlbStrumenti.Name = "tlbStrumenti"
        Me.tlbStrumenti.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionInMenu
        Me.tlbStrumenti.Visible = False
        '
        'tlbImpostaStampante
        '
        Me.tlbImpostaStampante.Caption = "Imposta Stampante"
        Me.tlbImpostaStampante.Id = 16
        Me.tlbImpostaStampante.Name = "tlbImpostaStampante"
        Me.tlbImpostaStampante.Visible = True
        '
        'tlbGuida
        '
        Me.tlbGuida.Caption = "Guida"
        Me.tlbGuida.Glyph = CType(resources.GetObject("tlbGuida.Glyph"), System.Drawing.Image)
        Me.tlbGuida.Id = 11
        Me.tlbGuida.Name = "tlbGuida"
        Me.tlbGuida.Visible = True
        '
        'tlbEsci
        '
        Me.tlbEsci.Caption = "Esci"
        Me.tlbEsci.Glyph = CType(resources.GetObject("tlbEsci.Glyph"), System.Drawing.Image)
        Me.tlbEsci.Id = 12
        Me.tlbEsci.Name = "tlbEsci"
        Me.tlbEsci.Visible = True
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Margin = New System.Windows.Forms.Padding(368, 111, 368, 111)
        Me.barDockControlTop.Size = New System.Drawing.Size(251, 35)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 128)
        Me.barDockControlBottom.Margin = New System.Windows.Forms.Padding(368, 111, 368, 111)
        Me.barDockControlBottom.Size = New System.Drawing.Size(251, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 35)
        Me.barDockControlLeft.Margin = New System.Windows.Forms.Padding(368, 111, 368, 111)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 93)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(251, 35)
        Me.barDockControlRight.Margin = New System.Windows.Forms.Padding(368, 111, 368, 111)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 93)
        '
        'tlbNuovo
        '
        Me.tlbNuovo.Caption = "Nuovo"
        Me.tlbNuovo.Glyph = CType(resources.GetObject("tlbNuovo.Glyph"), System.Drawing.Image)
        Me.tlbNuovo.Id = 0
        Me.tlbNuovo.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F2)
        Me.tlbNuovo.Name = "tlbNuovo"
        Me.tlbNuovo.Visible = True
        '
        'tlbCancella
        '
        Me.tlbCancella.Caption = "Cancella"
        Me.tlbCancella.Glyph = CType(resources.GetObject("tlbCancella.Glyph"), System.Drawing.Image)
        Me.tlbCancella.Id = 2
        Me.tlbCancella.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F4)
        Me.tlbCancella.Name = "tlbCancella"
        Me.tlbCancella.Visible = True
        '
        'tlbStampa
        '
        Me.tlbStampa.Caption = "Stampa"
        Me.tlbStampa.Glyph = CType(resources.GetObject("tlbStampa.Glyph"), System.Drawing.Image)
        Me.tlbStampa.Id = 4
        Me.tlbStampa.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F6)
        Me.tlbStampa.Name = "tlbStampa"
        Me.tlbStampa.Visible = True
        '
        'tlbStampaVideo
        '
        Me.tlbStampaVideo.Caption = "Stampa video"
        Me.tlbStampaVideo.Glyph = CType(resources.GetObject("tlbStampaVideo.Glyph"), System.Drawing.Image)
        Me.tlbStampaVideo.Id = 5
        Me.tlbStampaVideo.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F12)
        Me.tlbStampaVideo.Name = "tlbStampaVideo"
        Me.tlbStampaVideo.Visible = True
        '
        'tlbRecordCancella
        '
        Me.tlbRecordCancella.Caption = "Cancella riga"
        Me.tlbRecordCancella.Glyph = CType(resources.GetObject("tlbRecordCancella.Glyph"), System.Drawing.Image)
        Me.tlbRecordCancella.Id = 17
        Me.tlbRecordCancella.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F4))
        Me.tlbRecordCancella.Name = "tlbRecordCancella"
        Me.tlbRecordCancella.Visible = True
        '
        'tlbApri
        '
        Me.tlbApri.Caption = "Apri"
        Me.tlbApri.Glyph = CType(resources.GetObject("tlbApri.Glyph"), System.Drawing.Image)
        Me.tlbApri.Id = 19
        Me.tlbApri.Name = "tlbApri"
        Me.tlbApri.Visible = True
        '
        'tlbCancMultipla
        '
        Me.tlbCancMultipla.Caption = "CancMultipla"
        Me.tlbCancMultipla.Glyph = CType(resources.GetObject("tlbCancMultipla.Glyph"), System.Drawing.Image)
        Me.tlbCancMultipla.Id = 20
        Me.tlbCancMultipla.Name = "tlbCancMultipla"
        Me.tlbCancMultipla.Visible = True
        '
        'tlbDuplica
        '
        Me.tlbDuplica.Caption = "Duplica"
        Me.tlbDuplica.Glyph = CType(resources.GetObject("tlbDuplica.Glyph"), System.Drawing.Image)
        Me.tlbDuplica.Id = 21
        Me.tlbDuplica.Name = "tlbDuplica"
        Me.tlbDuplica.Visible = True
        '
        'pnTop
        '
        Me.pnTop.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnTop.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.pnTop.Appearance.Options.UseBackColor = True
        Me.pnTop.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple
        Me.pnTop.Controls.Add(Me.ck_maeriali)
        Me.pnTop.Controls.Add(Me.ck_mezzi)
        Me.pnTop.Controls.Add(Me.ck_risorse)
        Me.pnTop.Location = New System.Drawing.Point(4, 38)
        Me.pnTop.Margin = New System.Windows.Forms.Padding(49, 26, 49, 26)
        Me.pnTop.Name = "pnTop"
        Me.pnTop.Size = New System.Drawing.Size(192, 200)
        '
        'ck_maeriali
        '
        Me.ck_maeriali.Location = New System.Drawing.Point(4, 52)
        Me.ck_maeriali.Name = "ck_maeriali"
        Me.ck_maeriali.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.ck_maeriali.Properties.Appearance.Options.UseBackColor = True
        Me.ck_maeriali.Properties.Appearance.Options.UseTextOptions = True
        Me.ck_maeriali.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.ck_maeriali.Properties.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.ck_maeriali.Properties.AutoHeight = False
        Me.ck_maeriali.Properties.Caption = "Materiali"
        Me.ck_maeriali.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.UserDefined
        Me.ck_maeriali.Properties.PictureChecked = CType(resources.GetObject("ck_maeriali.Properties.PictureChecked"), System.Drawing.Image)
        Me.ck_maeriali.Properties.PictureUnchecked = CType(resources.GetObject("ck_maeriali.Properties.PictureUnchecked"), System.Drawing.Image)
        Me.ck_maeriali.Size = New System.Drawing.Size(116, 20)
        '
        'ck_mezzi
        '
        Me.ck_mezzi.Location = New System.Drawing.Point(4, 28)
        Me.ck_mezzi.Name = "ck_mezzi"
        Me.ck_mezzi.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.ck_mezzi.Properties.Appearance.Options.UseBackColor = True
        Me.ck_mezzi.Properties.Appearance.Options.UseTextOptions = True
        Me.ck_mezzi.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.ck_mezzi.Properties.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.ck_mezzi.Properties.AutoHeight = False
        Me.ck_mezzi.Properties.Caption = "Mezzi"
        Me.ck_mezzi.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.UserDefined
        Me.ck_mezzi.Properties.PictureChecked = CType(resources.GetObject("ck_mezzi.Properties.PictureChecked"), System.Drawing.Image)
        Me.ck_mezzi.Properties.PictureUnchecked = CType(resources.GetObject("ck_mezzi.Properties.PictureUnchecked"), System.Drawing.Image)
        Me.ck_mezzi.Size = New System.Drawing.Size(116, 20)
        '
        'ck_risorse
        '
        Me.ck_risorse.Location = New System.Drawing.Point(4, 4)
        Me.ck_risorse.Name = "ck_risorse"
        Me.ck_risorse.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.ck_risorse.Properties.Appearance.Options.UseBackColor = True
        Me.ck_risorse.Properties.Appearance.Options.UseTextOptions = True
        Me.ck_risorse.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        Me.ck_risorse.Properties.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.ck_risorse.Properties.AutoHeight = False
        Me.ck_risorse.Properties.Caption = "Risorse"
        Me.ck_risorse.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.UserDefined
        Me.ck_risorse.Properties.PictureChecked = CType(resources.GetObject("ck_risorse.Properties.PictureChecked"), System.Drawing.Image)
        Me.ck_risorse.Properties.PictureUnchecked = CType(resources.GetObject("ck_risorse.Properties.PictureUnchecked"), System.Drawing.Image)
        Me.ck_risorse.Size = New System.Drawing.Size(116, 20)
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'FRMZCIMAN
        '
        Me.ClientSize = New System.Drawing.Size(251, 128)
        Me.Controls.Add(Me.pnTop)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.Margin = New System.Windows.Forms.Padding(368, 111, 368, 111)
        Me.Name = "FRMZCIMAN"
        Me.Text = "IMPORT DA CRM"
        CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NtsBarManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnTop, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnTop.ResumeLayout(False)
        CType(Me.ck_maeriali.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ck_mezzi.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ck_risorse.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Public Overloads Function Init(ByRef Menu As CLE__MENU, ByRef Param As CLE__CLDP, Optional ByVal Ditta As String = "", Optional ByRef SharedControls As CLE__EVNT = Nothing) As Boolean
        oMenu = Menu
        oApp = oMenu.App
        oCallParams = Param
        If Ditta <> "" Then
            DittaCorrente = Ditta
        Else
            DittaCorrente = oApp.Ditta
        End If
        Me.GctlTipoDoc = ""

        InitializeComponent()
        Me.MinimumSize = Me.Size

        '------------------------------------------------
        'creo e attivo l'entity e inizializzo la funzione che dovrà rilevare gli eventi dall'ENTITY
        Dim strErr As String = ""
        Dim oTmp As Object = Nothing
        If CLN__STD.NTSIstanziaDll(oApp.ServerDir, oApp.NetDir, "BNZCIMAN", "BEZCIMAN", oTmp, strErr, False, "", "") = False Then
            oApp.MsgBoxErr(oApp.Tr(Me, 127930271093451112, "ERRORE in fase di creazione Entity:" & vbCrLf & "|" & strErr & "|"))
            Return False
        End If
        oCleIMAN = CType(oTmp, CLEZCIMAN)

        AddHandler oCleIMAN.RemoteEvent, AddressOf GestisciEventiEntity
        If oCleIMAN.Init(oApp, oScript, oMenu.oCleComm, "", False, "", "") = False Then Return False

        ''------------------------------------------------
        strErr = ""
        oTmp = Nothing
        If CLN__STD.NTSIstanziaDll(oApp.ServerDir, oApp.NetDir, "BNZCIMAN", "BEVEBOLL", oTmp, strErr, False, "", "") = False Then
            oApp.MsgBoxErr(oApp.Tr(Me, 127791222114531250, "ERRORE in fase di creazione Entity:" & vbCrLf & "|" & strErr & "|"))
            Return False
        End If
        oCleBoll = CType(oTmp, CLEVEBOLL)
        AddHandler oCleBoll.RemoteEvent, AddressOf GestisciEventiEntity
        If oCleBoll.Init(oApp, oScript, oMenu.oCleComm, "", False, "", "") = False Then Return False
        If Not oCleBoll.InitExt() Then
            Return False
        End If

        ''------------------------------------------------

        Return True
    End Function

    Public Overridable Sub FRMZCIMAN_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        'salvo l'impostazione della griglia. devo farlo anche al cambio del tipo documento (ad esempio in bsveboll
        Try
        Catch
        End Try
    End Sub

    Public Overridable Sub FRMZCIMAN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim i As Integer
        Try
            '-------------------------------------------------
            'predispongo i controlli
            InitControls()


            '-------------------------------------------------
            'sempre alla fine di questa funzione: applico le regole della gctl
            GctlSetRoules()

        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub

    Public Overridable Sub InitControls()
        Dim dttOrtipo As New DataTable()
        Try
            '-------------------------------------------------
            'carico le immagini della toolbar
            Try
                tlbElabora.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\elabora.png")
                tlbZoom.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\zoom.png")
                tlbStrumenti.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\options.png")
                tlbGuida.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\help.png")
                tlbEsci.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\exit.png")

                ck_maeriali.NTSSetParam(oMenu, oApp.Tr(Me, 132545739825409632, "Materiali"), "S", "N")
                ck_risorse.NTSSetParam(oMenu, oApp.Tr(Me, 132361887722038626, "Risorse"), "S", "N")
                ck_mezzi.NTSSetParam(oMenu, oApp.Tr(Me, 132357356439873147, "Mezzi"), "S", "N")
            Catch ex As Exception
                'non gestisco l'errore: se non c'è una immagine prendo quella standard
            End Try

            tlbMain.NTSSetToolTip()
            '-------------------------------------------------
            'chiamo lo script per inizializzare i controlli caricati con source ext
            NTSScriptExec("InitControls", Me, Nothing)
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub


#Region "Eventi Toolbar"

    Public Overridable Sub tlbZoom_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbZoom.ItemClick
        Try
            '------------------------------------
            'zoom standard di textbox e griglia
            NTSCallStandardZoom()
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub

    Public Overridable Sub tlbGuida_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbGuida.ItemClick
        SendKeys.Send("{F1}")
    End Sub

    Public Overridable Sub tlbEsci_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbEsci.ItemClick
        Me.Close()
    End Sub

#End Region


    Public Overridable Sub tlbGiornaliera_ItemClick(sender As Object, e As ItemClickEventArgs) Handles tlbElabora.ItemClick
        Dim oParam As New CLE__CLDP
        Dim RigheImp As Integer = 0

        Try

            oCleIMAN.LogStart("BSZCIMAN_" & Today.Year & "-" & Today.Month & "-" & Today.Day & "_" & Now.Hour & "h" & Now.Minute & "m" & Now.Second & "s", Me.Text)


            Me.Cursor = Cursors.WaitCursor

            RigheImp = 0

            If oCleIMAN.Elabora(ck_risorse.Checked, ck_mezzi.Checked, ck_maeriali.Checked, RigheImp, oCleBoll) Then
                Me.Cursor = Cursors.Default
                ' oApp.MsgBoxInfo(oApp.Tr(Me, 130178543158859884, "Operazione Eseguita. Importate " & RigheImp & " Righe."))
                oApp.MsgBoxInfo(oApp.Tr(Me, 130178543158859884, "Operazione Eseguita."))

            Else
                Me.Cursor = Cursors.Default
                NTSProcessStart("notepad", oCleIMAN.LogFileName)
            End If



        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
            oCleIMAN.LogWrite(ex.Message & vbCrLf & ex.StackTrace, True)
        Finally
            oCleIMAN.LogStop()
        End Try
    End Sub




End Class
