Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class Form1
    Inherits Form

    Private ReadOnly txtFolder As New TextBox()
    Private ReadOnly txtIconFolder As New TextBox()
    Private ReadOnly cboIconNumber As New ComboBox()

    Private ReadOnly btnBrowseFolder As New Button()
    Private ReadOnly btnBrowseIconFolder As New Button()
    Private ReadOnly btnApply As New Button()

    Private ReadOnly lblStatus As New Label()
    Private ReadOnly lstIcons As New ListBox()
    Private ReadOnly picPreview As New PictureBox()

    Private ReadOnly lblTitle As New Label()
    Private ReadOnly lblSubTitle As New Label()

    Private ReadOnly panelTop As New Panel()
    Private ReadOnly panelMain As New Panel()
    Private ReadOnly panelRight As New Panel()
    Private ReadOnly panelPreview As New Panel()
    Private ReadOnly lblPreviewText As New Label()

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Sub SHChangeNotify(
        wEventId As Integer,
        uFlags As UInteger,
        dwItem1 As String,
        dwItem2 As IntPtr
    )
    End Sub

    Public Sub New()
        Me.Text = "FolderColors"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(980, 620)
        Me.MinimumSize = New Size(980, 620)
        Me.Font = New Font("Segoe UI", 10.0F)
        Me.BackColor = Color.FromArgb(236, 244, 234)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        ' Icône de l'application :
        ' C:\Windows\File.ico
        Dim appIconPath As String =
            Path.Combine("C:\Windows", "File.ico")

        If File.Exists(appIconPath) Then
            Try
                ' CORRECTION : appIconPath est une chaîne de caractères.
                ' Il ne faut pas utiliser New Icon(iconFile).
                Me.Icon = New Icon(appIconPath)
            Catch ex As Exception
                MessageBox.Show(
                    "Impossible de charger l'icône de l'application :" &
                    Environment.NewLine &
                    ex.Message,
                    "Erreur",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        Else
            MessageBox.Show(
                "L'icône est introuvable :" &
                Environment.NewLine &
                appIconPath,
                "Icône introuvable",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
        End If

        BuildInterface()
        RefreshIconList()
    End Sub

    Private Function GetDefaultIconFolder() As String
        Return "C:\Windows"
    End Function

    Private Sub BuildInterface()
        ' =========================
        ' Barre supérieure
        ' =========================
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 95
        panelTop.BackColor = Color.FromArgb(28, 88, 58)
        Me.Controls.Add(panelTop)

        lblTitle.Text = "FolderColors"
        lblTitle.ForeColor = Color.White
        lblTitle.Font = New Font("Segoe UI", 22.0F, FontStyle.Bold)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(24, 16)
        panelTop.Controls.Add(lblTitle)

        lblSubTitle.Text =
            "Sélectionnez un dossier, choisissez une icône et appliquez-la en un clic."

        lblSubTitle.ForeColor = Color.FromArgb(210, 230, 214)
        lblSubTitle.Font = New Font("Segoe UI", 10.0F)
        lblSubTitle.AutoSize = True
        lblSubTitle.Location = New Point(26, 55)
        panelTop.Controls.Add(lblSubTitle)

        ' =========================
        ' Panneau principal
        ' =========================
        panelMain.Location = New Point(22, 115)
        panelMain.Size = New Size(640, 455)
        panelMain.BackColor = Color.FromArgb(248, 251, 246)
        panelMain.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(panelMain)

        ' =========================
        ' Panneau aperçu
        ' =========================
        panelRight.Location = New Point(680, 115)
        panelRight.Size = New Size(275, 455)
        panelRight.BackColor = Color.FromArgb(248, 251, 246)
        panelRight.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(panelRight)

        ' =========================
        ' Dossier cible
        ' =========================
        Dim lblFolder As New Label()
        lblFolder.Text = "Dossier cible"
        lblFolder.AutoSize = True
        lblFolder.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lblFolder.ForeColor = Color.FromArgb(28, 88, 58)
        lblFolder.Location = New Point(22, 22)
        panelMain.Controls.Add(lblFolder)

        txtFolder.Location = New Point(22, 50)
        txtFolder.Size = New Size(470, 29)
        txtFolder.Text =
            Environment.GetFolderPath(
                Environment.SpecialFolder.DesktopDirectory
            )
        txtFolder.BackColor = Color.White
        panelMain.Controls.Add(txtFolder)

        btnBrowseFolder.Text = "Parcourir"
        btnBrowseFolder.Location = New Point(505, 48)
        btnBrowseFolder.Size = New Size(105, 33)
        btnBrowseFolder.BackColor = Color.FromArgb(61, 138, 87)
        btnBrowseFolder.ForeColor = Color.White
        btnBrowseFolder.FlatStyle = FlatStyle.Flat
        AddHandler btnBrowseFolder.Click, AddressOf BrowseFolder_Click
        panelMain.Controls.Add(btnBrowseFolder)

        ' =========================
        ' Dossier des fichiers ICO
        ' =========================
        Dim lblIconFolder As New Label()
        lblIconFolder.Text = "Dossier contenant les fichiers .ico"
        lblIconFolder.AutoSize = True
        lblIconFolder.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lblIconFolder.ForeColor = Color.FromArgb(28, 88, 58)
        lblIconFolder.Location = New Point(22, 98)
        panelMain.Controls.Add(lblIconFolder)

        txtIconFolder.Location = New Point(22, 126)
        txtIconFolder.Size = New Size(470, 29)
        txtIconFolder.Text = GetDefaultIconFolder()
        txtIconFolder.BackColor = Color.White
        panelMain.Controls.Add(txtIconFolder)

        btnBrowseIconFolder.Text = "Parcourir"
        btnBrowseIconFolder.Location = New Point(505, 124)
        btnBrowseIconFolder.Size = New Size(105, 33)
        btnBrowseIconFolder.BackColor = Color.FromArgb(61, 138, 87)
        btnBrowseIconFolder.ForeColor = Color.White
        btnBrowseIconFolder.FlatStyle = FlatStyle.Flat
        AddHandler btnBrowseIconFolder.Click,
            AddressOf BrowseIconFolder_Click
        panelMain.Controls.Add(btnBrowseIconFolder)

        ' =========================
        ' Sélection de l'icône
        ' =========================
        Dim lblNumber As New Label()
        lblNumber.Text = "Choisir l'icône"
        lblNumber.AutoSize = True
        lblNumber.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lblNumber.ForeColor = Color.FromArgb(28, 88, 58)
        lblNumber.Location = New Point(22, 178)
        panelMain.Controls.Add(lblNumber)

        cboIconNumber.DropDownStyle = ComboBoxStyle.DropDownList
        cboIconNumber.Location = New Point(22, 206)
        cboIconNumber.Size = New Size(120, 29)

        For i As Integer = 1 To 15
            cboIconNumber.Items.Add(i.ToString())
        Next

        cboIconNumber.SelectedIndex = 0
        AddHandler cboIconNumber.SelectedIndexChanged,
            AddressOf IconSelectionChanged
        panelMain.Controls.Add(cboIconNumber)

        btnApply.Text = "Appliquer l'icône"
        btnApply.Location = New Point(160, 202)
        btnApply.Size = New Size(180, 37)
        btnApply.BackColor = Color.FromArgb(33, 122, 73)
        btnApply.ForeColor = Color.White
        btnApply.FlatStyle = FlatStyle.Flat
        AddHandler btnApply.Click, AddressOf ApplyIcon_Click
        panelMain.Controls.Add(btnApply)

        ' =========================
        ' Liste des icônes
        ' =========================
        Dim lblList As New Label()
        lblList.Text = "Icônes détectées"
        lblList.AutoSize = True
        lblList.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lblList.ForeColor = Color.FromArgb(28, 88, 58)
        lblList.Location = New Point(22, 268)
        panelMain.Controls.Add(lblList)

        lstIcons.Location = New Point(22, 298)
        lstIcons.Size = New Size(600, 130)
        lstIcons.BackColor = Color.White
        AddHandler lstIcons.SelectedIndexChanged,
            AddressOf IconListSelectionChanged
        panelMain.Controls.Add(lstIcons)

        ' =========================
        ' Statut
        ' =========================
        lblStatus.Text = "En attente..."
        lblStatus.AutoSize = True
        lblStatus.Font = New Font("Segoe UI", 9.0F)
        lblStatus.ForeColor = Color.FromArgb(100, 100, 100)
        lblStatus.Location = New Point(22, 432)
        panelMain.Controls.Add(lblStatus)

        ' =========================
        ' Aperçu
        ' =========================
        panelPreview.Location = New Point(5, 5)
        panelPreview.Size = New Size(265, 445)
        panelPreview.BackColor = Color.White
        panelPreview.BorderStyle = BorderStyle.FixedSingle
        panelRight.Controls.Add(panelPreview)

        lblPreviewText.Text = "Aperçu"
        lblPreviewText.AutoSize = True
        lblPreviewText.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lblPreviewText.ForeColor = Color.FromArgb(28, 88, 58)
        lblPreviewText.Location = New Point(8, 8)
        panelPreview.Controls.Add(lblPreviewText)

        picPreview.Location = New Point(15, 35)
        picPreview.Size = New Size(240, 240)
        picPreview.SizeMode = PictureBoxSizeMode.Zoom
        picPreview.BackColor = Color.White
        panelPreview.Controls.Add(picPreview)
    End Sub

    Private Sub RefreshIconList()
        lstIcons.Items.Clear()

        Dim folder As String = txtIconFolder.Text.Trim()

        If String.IsNullOrWhiteSpace(folder) OrElse
           Not Directory.Exists(folder) Then

            lblStatus.Text = "Dossier d'icônes invalide."
            Return
        End If

        Try
            Dim files As String() =
                Directory.GetFiles(folder, "*.ico")

            If files.Length = 0 Then
                lblStatus.Text = "Aucun fichier .ico trouvé."
                Return
            End If

            Array.Sort(files)

            For Each file As String In files
                lstIcons.Items.Add(Path.GetFileName(file))
            Next

            lblStatus.Text =
                files.Length.ToString() & " icône(s) trouvée(s)."

            Dim premiereIcone As String = Path.Combine(
                folder,
                "My-Folder-1.ico"
            )

            If File.Exists(premiereIcone) Then
                LoadPreview(premiereIcone)
            End If

        Catch ex As Exception
            lblStatus.Text =
                "Erreur lors de la lecture du dossier d'icônes."
        End Try
    End Sub

    Private Sub BrowseFolder_Click(
        sender As Object,
        e As EventArgs
    )
        Using d As New FolderBrowserDialog()
            d.Description =
                "Sélectionnez le dossier auquel appliquer l'icône."

            d.SelectedPath = txtFolder.Text
            d.ShowNewFolderButton = False

            If d.ShowDialog() = DialogResult.OK Then
                txtFolder.Text = d.SelectedPath
                lblStatus.Text = "Dossier cible sélectionné."
            End If
        End Using
    End Sub

    Private Sub BrowseIconFolder_Click(
        sender As Object,
        e As EventArgs
    )
        Using d As New FolderBrowserDialog()
            d.Description =
                "Sélectionnez le dossier contenant les fichiers .ico."

            d.SelectedPath = txtIconFolder.Text
            d.ShowNewFolderButton = False

            If d.ShowDialog() = DialogResult.OK Then
                txtIconFolder.Text = d.SelectedPath
                RefreshIconList()
            End If
        End Using
    End Sub

    Private Sub IconSelectionChanged(
        sender As Object,
        e As EventArgs
    )
        If cboIconNumber.SelectedIndex < 0 Then
            Return
        End If

        Dim numero As Integer =
            cboIconNumber.SelectedIndex + 1

        Dim cheminIcone As String = Path.Combine(
            txtIconFolder.Text.Trim(),
            "My-Folder-" & numero.ToString() & ".ico"
        )

        If File.Exists(cheminIcone) Then
            LoadPreview(cheminIcone)
        Else
            ClearPreview()

            lblStatus.Text =
                "Icône introuvable : " &
                Path.GetFileName(cheminIcone)
        End If
    End Sub

    Private Sub IconListSelectionChanged(
        sender As Object,
        e As EventArgs
    )
        If lstIcons.SelectedIndex < 0 Then
            Return
        End If

        Dim nomFichier As String =
            lstIcons.SelectedItem.ToString()

        Dim cheminIcone As String = Path.Combine(
            txtIconFolder.Text.Trim(),
            nomFichier
        )

        If File.Exists(cheminIcone) Then
            LoadPreview(cheminIcone)
        End If
    End Sub

    Private Sub LoadPreview(cheminIcone As String)
        Try
            Using iconFile As New Icon(cheminIcone)
                ' Une icône doit être convertie en Bitmap
                ' pour être affichée dans un PictureBox.
                Dim nouvelleImage As Bitmap =
                    iconFile.ToBitmap()

                ClearPreview()
                picPreview.Image = nouvelleImage
            End Using

            lblStatus.Text =
                "Aperçu chargé : " &
                Path.GetFileName(cheminIcone)

        Catch ex As Exception
            ClearPreview()

            lblStatus.Text =
                "Impossible de charger cette icône."
        End Try
    End Sub

    Private Sub ClearPreview()
        If picPreview.Image IsNot Nothing Then
            Dim ancienneImage As Image =
                picPreview.Image

            picPreview.Image = Nothing
            ancienneImage.Dispose()
        End If
    End Sub

    Private Sub ApplyIcon_Click(
        sender As Object,
        e As EventArgs
    )
        Dim dossierCible As String =
            txtFolder.Text.Trim()

        If String.IsNullOrWhiteSpace(dossierCible) OrElse
           Not Directory.Exists(dossierCible) Then

            MessageBox.Show(
                "Veuillez sélectionner un dossier cible valide.",
                "Dossier invalide",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            Return
        End If

        If cboIconNumber.SelectedIndex < 0 Then
            MessageBox.Show(
                "Veuillez sélectionner une icône.",
                "Icône non sélectionnée",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            Return
        End If

        Dim numero As Integer =
            cboIconNumber.SelectedIndex + 1

        Dim fichierIcone As String = Path.Combine(
            txtIconFolder.Text.Trim(),
            "My-Folder-" & numero.ToString() & ".ico"
        )

        If Not File.Exists(fichierIcone) Then
            MessageBox.Show(
                "Le fichier suivant est introuvable :" &
                Environment.NewLine &
                fichierIcone,
                "Icône introuvable",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            Return
        End If

        Try
            Dim nomIcone As String =
                "FolderColors-" & numero.ToString() & ".ico"

            Dim iconeDansDossier As String =
                Path.Combine(dossierCible, nomIcone)

            Dim desktopIni As String =
                Path.Combine(dossierCible, "desktop.ini")

            ' Copie l'icône dans le dossier cible.
            File.Copy(
                fichierIcone,
                iconeDansDossier,
                True
            )

            Dim contenuDesktopIni As String =
                "[.ShellClassInfo]" &
                Environment.NewLine &
                "IconResource=" &
                nomIcone &
                ",0" &
                Environment.NewLine &
                "IconFile=" &
                nomIcone &
                Environment.NewLine &
                "IconIndex=0" &
                Environment.NewLine

            File.WriteAllText(
                desktopIni,
                contenuDesktopIni,
                Encoding.Unicode
            )

            File.SetAttributes(
                desktopIni,
                FileAttributes.Hidden Or FileAttributes.System
            )

            Dim attributsDossier As FileAttributes =
                File.GetAttributes(dossierCible)

            File.SetAttributes(
                dossierCible,
                attributsDossier Or FileAttributes.System
            )

            ' Actualisation de l'Explorateur Windows.
            SHChangeNotify(
                &H8000000,
                &H1000UI,
                dossierCible,
                IntPtr.Zero
            )

            lblStatus.Text =
                "Icône appliquée avec succès."

            MessageBox.Show(
                "L'icône a été appliquée au dossier." &
                Environment.NewLine &
                "Actualisez l'Explorateur Windows si nécessaire.",
                "Opération terminée",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        Catch ex As Exception
            lblStatus.Text =
                "Erreur lors de l'application de l'icône."

            MessageBox.Show(
                "Impossible d'appliquer l'icône :" &
                Environment.NewLine &
                ex.Message,
                "Erreur",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
        End Try
    End Sub

    Protected Overrides Sub OnFormClosed(
        e As FormClosedEventArgs
    )
        ClearPreview()
        MyBase.OnFormClosed(e)
    End Sub

End Class
