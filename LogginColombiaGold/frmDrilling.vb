Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports clsDHPlatform
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Imports System.Math
Imports vb = Microsoft.VisualBasic

Public Class frmDrilling
    Dim sFotoTopo As String
    Dim sFile As String


    Dim correo As New System.Net.Mail.MailMessage
    Dim oCollars As New clsDHCollars
    Dim oPlatform As New clsDHPlatform
    Dim sValidar As Boolean
    Dim sValidaDr As Boolean
    Dim sValidaPr As Boolean
    Dim sValidarTopo As Boolean
    Dim folder As New DirectoryInfo("C:\")
    Dim sEdit As String = "0"
    Dim sEditDr As String = "0"
    Dim sEditDP As String = "0"
    Dim sEditTo As String = "0"
    Dim sDelHP As Integer
    Dim Raiz1 As Double
    Dim Raiz2 As Double
    Dim Raiz3 As Double
    Dim X1 As Double
    Dim Y1 As Double
    Dim Z1 As Double
    Dim X2 As Double
    Dim Y2 As Double
    Dim Z2 As Double
    Dim X3 As Double
    Dim Y3 As Double
    Dim Z3 As Double
    Dim sMotivo As String

    Dim L1H As String
    Dim N1H As String
    Dim L2H As String

    Dim ValidaTopo As Boolean
    Dim ValidaST As Boolean
    Dim ValidaCs As Boolean
    Dim ValorCorreo As Double
    Dim Coordenada As String

    Dim sMetros As String
    Dim sVeta As String
    Dim sDepth As String
    Dim sMensaje As String

    Dim sFrom As String
    Dim sTo As String
    Dim sSubject As String
    Dim sUserSend As String
    Dim sPassSend As String
    Dim sServer As String

    Dim oRf As New clsRf

    'Company Drill
    '****************************************************
    Dim sEditCd As String = "0"
    Dim sID As Integer
    Dim sEditMeter As String = "0"
    Dim IdDc As Integer
    Dim sValidarMD As Boolean
    Dim sValidarDT As Boolean
    Dim sValidarCC As Boolean
    Dim sValidarTS As Boolean
    Dim sValidarLT As Boolean
    Dim sValidarBA As Boolean
    Dim sValidarDC As Boolean
    Dim sValidarDrTm As Boolean
    Dim sEditDrTm As Boolean

    Dim sEditDt As String = "0"
    Dim sEditCc As String = "0"
    Dim sEditTs As String = "0"
    Dim sEditBCom As String = "0"
    Dim sEditBCon As String = "0"
    Dim sEditLt As String = "0"
    Dim sEditBA As String = "0"

    Dim sRowEditDrilling As Integer
    Dim sRowEdit As Integer
    Dim sIDn As Integer
    Dim sEditPoll As Integer
    Dim sEditPollC As Integer
    Dim sValidarPoll As Boolean
    Dim svalidarPollC As Boolean
    Dim sValorPollH As Integer
    Dim sIdEdit As Integer
    Dim svalidarPollImpact As Boolean
    Dim sEditPollImpact As Boolean
    Dim sPollImpAdd As Boolean
    Dim sPollImpEdt As Boolean


    Private Function AutoCompleteCmb(ByVal _dtAutoComplete As DataTable) As AutoCompleteStringCollection
        Try

            Dim stringCol As New AutoCompleteStringCollection()
            For Each row As DataRow In _dtAutoComplete.Rows
                stringCol.Add(Convert.ToString(row("Comb")))

            Next

            Return stringCol

        Catch ex As Exception
            Throw New Exception(ex.Message)

        End Try

    End Function
    Function SoloNumeros(ByVal Keyascii As Short) As Short
        If InStr("1234567890.-,", Chr(Keyascii)) = 0 Then
            SoloNumeros = 0
        Else
            SoloNumeros = Keyascii
        End If
        Select Case Keyascii
            Case 8
                SoloNumeros = Keyascii
            Case 13
                SoloNumeros = Keyascii
        End Select
    End Function

    Private Sub FillPlatformIDForm()
        Try
            oPlatform.sPlatformID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtPlatform As DataTable = oPlatform.getDHPlatform()
            Dim drC As DataRow = dtPlatform.NewRow()
            drC(0) = "Select an option.."
            drC(1) = "Select an option.."
            dtPlatform.Rows.Add(drC)
            cmbHoleId.DataSource = dtPlatform
            cmbHoleId.DisplayMember = "PlatformID"
            cmbHoleId.ValueMember = "PlatformID"
            cmbHoleId.SelectedValue = "Select an option.."

            'cmbHoleId.AutoCompleteCustomSource = AutoCompleteCmb(dtPlatform)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillDrillProgress()
        Try
            oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
            Dim dtDrillProgress As DataTable = oPlatform.getDHDrillProgress()
            dgDrillProgress1.DataSource = dtDrillProgress
            dgDrillProgress1.Columns(3).Visible = False
            dgDrillProgress1.Columns(7).Visible = False
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub FillMaxTo()
        Try
            oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
            Dim dtPlatformProgress As DataTable = oPlatform.getDHDrillProgressMaxTo
            txtFrom1.Text = dtPlatformProgress.Rows(0)("To").ToString

            'MsgBox(dtPlatformProgress.Rows(0)("To").ToString)
            If dtPlatformProgress.Rows.Count = 0 Then
                'If dtPlatformProgress.Rows(0)("To").ToString = "" Then
                txtFrom1.Text = ""
            End If
        Catch ex As Exception
            'Throw New Exception(ex.Message)
            txtFrom1.Text = ""
        End Try
    End Sub

    Private Sub FillHoleInProgress()
        Try
            oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
            Dim dtHoleInProgress As DataTable = oPlatform.getDHHoleInProgress()
            dgHoleInProgress1.DataSource = dtHoleInProgress
        Catch ex As Exception
            'MsgBox(ex.Message)
            'Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub FillEnvironmentPoll()
        Try
            oPlatform.sPlatform = txtPlatform.Text.ToString
            Dim dtEnvironmentPoll As DataTable = oPlatform.getDH_Environment_Poll_Platform()
            dgEnvironmentPollQuery.DataSource = dtEnvironmentPoll

            dgEnvironmentPollQuery.Columns(0).Visible = False
            dgEnvironmentPollQuery.Columns(1).Visible = False
            dgEnvironmentPollQuery.Columns(2).Visible = False
            'dgEnvironmentPollQuery.

        Catch ex As Exception
            'MsgBox(ex.Message)
            'Throw New Exception(ex.Message)
        End Try
    End Sub
    'Private Sub FillPlatformIDHold()
    '    Try
    '        oPlatform.sPlatformIDHold = cmbHoleId.SelectedValue.ToString
    '        'cmbHoleId.ValueMember = ""
    '        Dim dtPlatformHold As DataTable = oPlatform.getDHPlatformHold()
    '        Dim drC As DataRow = dtPlatformHold.NewRow()
    '        dgPozos.DataSource = dtPlatformHold

    '        'cmbHoleId.AutoCompleteCustomSource = AutoCompleteCmb(dtPlatform)

    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try

    'End Sub

    'Private Sub FillHoleIDPlatform()
    '    Try
    '        oPlatform.sHolePlatform = txtPlatform.Text.ToString
    '        'cmbHoleId.ValueMember = ""
    '        Dim dtHoldPlatform As DataTable = oPlatform.getHolePlatform
    '        Dim drC As DataRow = dtHoldPlatform.NewRow()
    '        dgHolePlatform.DataSource = dtHoldPlatform

    '        'cmbHoleId.AutoCompleteCustomSource = AutoCompleteCmb(dtPlatform)

    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try

    'End Sub
    Private Sub llenarPorcentaje()
        Try
            'oPlatform.sRfTorelocate = ""
            'cmbToRelocateD.ValueMember = ""
            'Dim dtToRelocate As DataTable = oPlatform.getRfTorelocate()
            'Dim drC As DataRow = dtToRelocate.NewRow()
            cmbPorcentajeTool.Items.Clear()
            Dim i As Integer
            For i = 0 To 100 Step 10
                cmbPorcentajeTool.Items.Add(i)
            Next
            cmbPorcentajeTool.SelectedItem = 0

            cmbPorcentajeAdmon.Items.Clear()
            For i = 0 To 100 Step 5
                cmbPorcentajeAdmon.Items.Add(i)
            Next
            cmbPorcentajeAdmon.SelectedItem = 0

            'drC(0) = "-1"
            'drC(1) = "Select an option.."
            'dtToRelocate.Rows.Add(drC)


            'cmbToRelocateD.DataSource = dtToRelocate

            'cmbToRelocateD.DisplayMember = "Description"
            'cmbToRelocateD.ValueMember = "Code"
            'cmbToRelocateD.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillTorelocate()
        Try
            oPlatform.sRfTorelocate = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtToRelocate As DataTable = oPlatform.getRfTorelocate()
            Dim drC As DataRow = dtToRelocate.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtToRelocate.Rows.Add(drC)
            cmbToRelocateD.DataSource = dtToRelocate
            cmbToRelocateD.DisplayMember = "Description"
            cmbToRelocateD.ValueMember = "Code"
            cmbToRelocateD.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillStatusPlatform()
        Try
            oPlatform.sRfStatusPlatform = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtStatusPlatform As DataTable = oPlatform.getRfStatusPlatform()
            Dim drC As DataRow = dtStatusPlatform.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtStatusPlatform.Rows.Add(drC)
            cmbStatusPlatform.DataSource = dtStatusPlatform
            cmbStatusPlatform.DisplayMember = "Description"
            cmbStatusPlatform.ValueMember = "Code"
            cmbStatusPlatform.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillZone1()
        Try
            oPlatform.sRfZone = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtZone As DataTable = oPlatform.getRfZone()
            Dim drC As DataRow = dtZone.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtZone.Rows.Add(drC)
            cmbZone1.DataSource = dtZone
            cmbZone1.DisplayMember = "Description"
            cmbZone1.ValueMember = "Code"
            cmbZone1.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillZone2()
        Try
            oPlatform.sRfZone = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtZone As DataTable = oPlatform.getRfZone()
            Dim drC As DataRow = dtZone.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtZone.Rows.Add(drC)
            cmbZone2.DataSource = dtZone
            cmbZone2.DisplayMember = "Description"
            cmbZone2.ValueMember = "Code"
            cmbZone2.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillZone3()
        Try
            oPlatform.sRfZone = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtZone3 As DataTable = oPlatform.getRfZone()
            Dim drC As DataRow = dtZone3.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtZone3.Rows.Add(drC)
            cmbZone3.DataSource = dtZone3
            cmbZone3.DisplayMember = "Description"
            cmbZone3.ValueMember = "Code"
            cmbZone3.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillBeta1()
        Try
            oPlatform.sRfTarguet = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtBeta As DataTable = oPlatform.getRfTarguet()
            Dim drC As DataRow = dtBeta.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtBeta.Rows.Add(drC)
            cmbBeta1.DataSource = dtBeta
            cmbBeta1.DisplayMember = "Description"
            cmbBeta1.ValueMember = "Code"
            cmbBeta1.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillTarguet()
        Try
            oPlatform.sRfTarguet = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtTarguet As DataTable = oPlatform.getRfTarguet()
            Dim drC As DataRow = dtTarguet.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtTarguet.Rows.Add(drC)
            cmbLocation.DataSource = dtTarguet
            cmbLocation.DisplayMember = "Description"
            cmbLocation.ValueMember = "Code"
            cmbLocation.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillBeta2()
        Try
            oPlatform.sRfTarguet = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtBeta2 As DataTable = oPlatform.getRfTarguet()
            Dim drC As DataRow = dtBeta2.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtBeta2.Rows.Add(drC)
            cmbBeta2.DataSource = dtBeta2
            cmbBeta2.DisplayMember = "Description"
            cmbBeta2.ValueMember = "Code"
            cmbBeta2.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillBeta3()
        Try
            oPlatform.sRfTarguet = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtBeta3 As DataTable = oPlatform.getRfTarguet()
            Dim drC As DataRow = dtBeta3.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtBeta3.Rows.Add(drC)
            cmbBeta3.DataSource = dtBeta3
            cmbBeta3.DisplayMember = "Description"
            cmbBeta3.ValueMember = "Code"
            cmbBeta3.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillPriority()
        Try
            oPlatform.sRfPriority = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtPriority As DataTable = oPlatform.getRfPriority()
            Dim drC As DataRow = dtPriority.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtPriority.Rows.Add(drC)
            cmbPriority.DataSource = dtPriority
            cmbPriority.DisplayMember = "Description"
            cmbPriority.ValueMember = "Code"
            cmbPriority.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillGroup()
        Try
            oPlatform.sGroup = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtEnvironmentGroup As DataTable = oPlatform.getRFEnvironmentGroup()
            Dim drC As DataRow = dtEnvironmentGroup.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtEnvironmentGroup.Rows.Add(drC)
            cmbGroup.DataSource = dtEnvironmentGroup
            cmbGroup.DisplayMember = "QuestionGroup"
            cmbGroup.ValueMember = "ID"
            cmbGroup.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillSubGroup()
        Try
            oPlatform.sGroup = cmbGroup.SelectedValue.ToString
            'cmbToRelocateD.ValueMember = ""
            Dim dtEnvironmentSubGroup As DataTable = oPlatform.getRFEnvironmentSubGroup()
            Dim drCG As DataRow = dtEnvironmentSubGroup.NewRow()
            drCG(0) = "-1"
            'drCG(1) = "Select an option.."
            ''cmbSubGroup.SelectedValue = dtEnvironmentSubGroup.Rows("ID")
            dtEnvironmentSubGroup.Rows.Add(drCG)
            cmbSubGroup.DataSource = dtEnvironmentSubGroup
            cmbSubGroup.DisplayMember = "QuestionSubGroup"
            cmbSubGroup.ValueMember = "ID"
            cmbSubGroup.SelectedValue = "-1"
            cmbSubGroup.SelectedText = "Select an option.."

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillImpact()
        Try
            'oPlatform.sSubGroup = cmbSubGroup.SelectedValue.ToString

            oPlatform.sID = ""
            Dim dtEnvironmentImpact As DataTable = oPlatform.getDH_Environment_Impact()

            Dim i As Integer

            dgImpact.Rows.Clear()
            For i = 0 To dtEnvironmentImpact.Rows.Count - 1
                dgImpact.Rows.Add(dtEnvironmentImpact.Rows(i)("ID"), dtEnvironmentImpact.Rows(i)("Impact"))
            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillQuestion()
        Try
            oPlatform.sSubGroup = cmbSubGroup.SelectedValue.ToString

            Dim dtEnvironmentQuestion As DataTable = oPlatform.getRFEnvironmentQuestion()

            Dim i As Integer

            dgQuestion.Rows.Clear()
            For i = 0 To dtEnvironmentQuestion.Rows.Count - 1
                dgQuestion.Rows.Add(dtEnvironmentQuestion.Rows(i)("ID"), dtEnvironmentQuestion.Rows(i)("ID_SG"), dtEnvironmentQuestion.Rows(i)("Question"))

            Next

            With Me.dgQuestion.RowTemplate
                '.DefaultCellStyle.BackColor = Color.Bisque
                .Height = 44
                .MinimumHeight = 20
            End With
            dgQuestion.RowTemplate.Height = 44

            ' dgQuestion.Refresh()


        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillSurface()
        Try
            oPlatform.sRfZone = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtSurface As DataTable = oPlatform.getRfSurface()
            Dim drC As DataRow = dtSurface.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtSurface.Rows.Add(drC)
            cmbSurface.DataSource = dtSurface
            cmbSurface.DisplayMember = "Description"
            cmbSurface.ValueMember = "Code"
            cmbSurface.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillSurveryor()
        Try
            oPlatform.sRfSurveryor = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtSurveryor As DataTable = oPlatform.getRfSurveryor()
            Dim drC As DataRow = dtSurveryor.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtSurveryor.Rows.Add(drC)
            cmbSurveryor.DataSource = dtSurveryor
            cmbSurveryor.DisplayMember = "Description"
            cmbSurveryor.ValueMember = "Code"
            cmbSurveryor.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillHoleStructure()

        Try
            oPlatform.sIDProject = ConfigurationSettings.AppSettings("IDProject").ToString()
            'cmbToRelocateD.ValueMember = ""
            Dim dtHoleStructure As DataTable = oPlatform.getDHHoleStructure()
            L1H = dtHoleStructure.Rows(0)("L1H").ToString()
            N1H = dtHoleStructure.Rows(0)("N1H").ToString()
            L2H = dtHoleStructure.Rows(0)("L2H").ToString()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillSurveryorSt()
        Try
            oPlatform.sRfSurveryor = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtSurveryor As DataTable = oPlatform.getRfSurveryor()
            Dim drC As DataRow = dtSurveryor.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtSurveryor.Rows.Add(drC)
            cmbSurveryorST.DataSource = dtSurveryor
            cmbSurveryorST.DisplayMember = "Description"
            cmbSurveryorST.ValueMember = "Code"
            cmbSurveryorST.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillCompanyServices()
        Try
            oPlatform.sRfCompanyServices = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtCompanyServices As DataTable = oPlatform.getRfCompanyServices()
            Dim drC As DataRow = dtCompanyServices.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtCompanyServices.Rows.Add(drC)
            cmdCompanyService.DataSource = dtCompanyServices
            cmdCompanyService.DisplayMember = "CompanyServices"
            cmdCompanyService.ValueMember = "Code"
            cmdCompanyService.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillLandPermit()
        Try
            oPlatform.sRfLandPermit = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtLandPermit As DataTable = oPlatform.getRfLandPermit()
            Dim drC As DataRow = dtLandPermit.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtLandPermit.Rows.Add(drC)
            cmbLandPermit.DataSource = dtLandPermit
            cmbLandPermit.DisplayMember = "Description"
            cmbLandPermit.ValueMember = "Code"
            cmbLandPermit.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillLandPermitStatus()
        Try
            oPlatform.sRfLandPermitStatus = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtLandPermit As DataTable = oPlatform.getRfLandPermitStatus()
            Dim drC As DataRow = dtLandPermit.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtLandPermit.Rows.Add(drC)
            cmbLandPermitStatus.DataSource = dtLandPermit
            cmbLandPermitStatus.DisplayMember = "Description"
            cmbLandPermitStatus.ValueMember = "Code"
            cmbLandPermitStatus.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    'Private Sub FillPercentProgress()
    '    Try
    '        oPlatform.sRfPercentProgress = ""
    '        'cmbToRelocateD.ValueMember = ""
    '        Dim dtPercentProgress As DataTable = oPlatform.getRfPercentProgress()
    '        Dim drC As DataRow = dtPercentProgress.NewRow()
    '        drC(0) = "-1"
    '        drC(1) = "Select an option.."
    '        dtPercentProgress.Rows.Add(drC)
    '        cmbProgress.DataSource = dtPercentProgress
    '        cmbProgress.DisplayMember = "Description"
    '        cmbProgress.ValueMember = "Code"
    '        cmbProgress.SelectedValue = "-1"

    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try

    'End Sub
    Private Sub FillContractor()
        Try
            oPlatform.sRfContractor = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtContractor As DataTable = oPlatform.getRfContractor()
            Dim drC As DataRow = dtContractor.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtContractor.Rows.Add(drC)
            cmbContractorDrill1.DataSource = dtContractor
            cmbContractorDrill1.DisplayMember = "Name"
            cmbContractorDrill1.ValueMember = "ID"
            cmbContractorDrill1.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillHoleIDForm()
        Try
            oCollars.sHoleID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtCollars As DataTable = oCollars.getDHCollars()
            Dim drC As DataRow = dtCollars.NewRow()
            drC(0) = "Select an option.."
            dtCollars.Rows.Add(drC)
            cmbHoleId.DataSource = dtCollars
            cmbHoleId.DisplayMember = "HoleID"
            cmbHoleId.ValueMember = "HoleID"
            cmbHoleId.SelectedValue = "Select an option.."

            'cmbHoleId.AutoCompleteCustomSource = AutoCompleteCmb(dtCollars)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillLocation()
        Try
            oPlatform.sCode = ""
            'cmbHoleId.ValueMember = ""
            Dim dtLocation As DataTable = oPlatform.getRfLocation()
            Dim drC As DataRow = dtLocation.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtLocation.Rows.Add(drC)
            cmbLocation.DataSource = dtLocation
            cmbLocation.DisplayMember = "Description"
            cmbLocation.ValueMember = "Code"
            cmbLocation.SelectedValue = "-1"

            'cmbHoleId.AutoCompleteCustomSource = AutoCompleteCmb(dtCollars)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillHoleIDCD()
        Try
            oCollars.sHoleID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtCollars As DataTable = oCollars.getDHCollars()
            Dim drC As DataRow = dtCollars.NewRow()
            drC(0) = "Select an option.."
            dtCollars.Rows.Add(drC)
            cmbHoleIdCD.DataSource = dtCollars
            cmbHoleIdCD.DisplayMember = "HoleID"
            cmbHoleIdCD.ValueMember = "HoleID"
            cmbHoleIdCD.SelectedValue = "Select an option.."

            'cmbHoleId.AutoCompleteCustomSource = AutoCompleteCmb(dtCollars)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillHoleDownTCD()
        Try
            oPlatform.sID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtDownTime As DataTable = oPlatform.getRfDrillDownTCD()
            Dim drC As DataRow = dtDownTime.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtDownTime.Rows.Add(drC)
            cmbDownTimeCD.DataSource = dtDownTime
            cmbDownTimeCD.DisplayMember = "Description"
            cmbDownTimeCD.ValueMember = "ID"
            cmbDownTimeCD.SelectedValue = "-1"

            'cmbHoleId.AutoCompleteCustomSource = AutoCompleteCmb(dtCollars)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillChangeCrownCD()
        Try
            oPlatform.sID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtDownTime As DataTable = oPlatform.getRfChangeCrownCD()
            Dim drC As DataRow = dtDownTime.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtDownTime.Rows.Add(drC)
            cmbChaCrown.DataSource = dtDownTime
            cmbChaCrown.DisplayMember = "Description"
            cmbChaCrown.ValueMember = "ID"
            cmbChaCrown.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillTurnSuppliesCD()
        Try
            oPlatform.sID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtTurnSupplies As DataTable = oPlatform.getRfTurnSuppliesCD()
            Dim drC As DataRow = dtTurnSupplies.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtTurnSupplies.Rows.Add(drC)
            cmbTurnSupplies.DataSource = dtTurnSupplies
            cmbTurnSupplies.DisplayMember = "Description"
            cmbTurnSupplies.ValueMember = "ID"
            cmbTurnSupplies.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillBiabilityCompCD()
        Try
            oPlatform.sID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtBiability As DataTable = oPlatform.getRfBiabilityCD()
            Dim drC As DataRow = dtBiability.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtBiability.Rows.Add(drC)
            cmbBuabilityCompany.DataSource = dtBiability
            cmbBuabilityCompany.DisplayMember = "Description"
            cmbBuabilityCompany.ValueMember = "ID"
            cmbBuabilityCompany.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillBiabilityContractorCD()
        Try
            oPlatform.sID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtBiability As DataTable = oPlatform.getRfBiabilityCD()
            Dim drC As DataRow = dtBiability.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtBiability.Rows.Add(drC)
            cmbBiabilityContractor.DataSource = dtBiability
            cmbBiabilityContractor.DisplayMember = "Description"
            cmbBiabilityContractor.ValueMember = "ID"
            cmbBiabilityContractor.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub FillLostToolsCD()
        Try
            oPlatform.sID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtLostTools As DataTable = oPlatform.getRfLostToolsCD()
            Dim drC As DataRow = dtLostTools.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtLostTools.Rows.Add(drC)
            cmbLostTools.DataSource = dtLostTools
            cmbLostTools.DisplayMember = "Description"
            cmbLostTools.ValueMember = "ID"
            cmbLostTools.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillBillableAdditivesCD()
        Try
            oPlatform.sID = ""
            'cmbHoleId.ValueMember = ""
            Dim dtBillableAdditives As DataTable = oPlatform.getRfBillableAdditivesCD()
            Dim drC As DataRow = dtBillableAdditives.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtBillableAdditives.Rows.Add(drC)
            cmbBillableAddit.DataSource = dtBillableAdditives
            cmbBillableAddit.DisplayMember = "Description"
            cmbBillableAddit.ValueMember = "ID"
            cmbBillableAddit.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillPlatformID()
        Try
            Dim dtPlatform As DataTable = oPlatform.getDHPlatform
            'Dim drC As DataRow = dtCollars.NewRow()
            'MsgBox(dtCollars.Rows(0)("PlatFormID"))
            txtPlatform.Text = dtPlatform.Rows(0)("PlatFormID").ToString
            'MsgBox(dtPlatform.Rows(0)("HoleID").ToString)
            txtHoleID.Text = dtPlatform.Rows(0)("HoleID").ToString
            txtSection.Text = dtPlatform.Rows(0)("Section").ToString
            txtEastPlanned.Text = dtPlatform.Rows(0)("EastPlanned").ToString
            txtNorthPlanned.Text = dtPlatform.Rows(0)("NorthPlanned").ToString
            txtElevationPlanned.Text = dtPlatform.Rows(0)("ElevationPlanned").ToString
            txtAzPlanned.Text = dtPlatform.Rows(0)("AzimuthPlanned").ToString
            txtInclinationPlanned.Text = dtPlatform.Rows(0)("InclinationPlanned").ToString
            txtLengthPlanned.Text = dtPlatform.Rows(0)("LengthPlanned").ToString
            If dtPlatform.Rows(0)("ToRelocate").ToString = "" Then
                cmbToRelocateD.SelectedValue = "-1"
            Else
                cmbToRelocateD.SelectedValue = dtPlatform.Rows(0)("ToRelocate").ToString
            End If

            If dtPlatform.Rows(0)("Status").ToString = "" Then
                cmbStatusPlatform.SelectedValue = "-1"
            Else
                cmbStatusPlatform.SelectedValue = dtPlatform.Rows(0)("Status").ToString
            End If

            If dtPlatform.Rows(0)("Surface").ToString = "" Then
                cmbSurface.SelectedValue = "-1"
            Else
                cmbSurface.SelectedValue = dtPlatform.Rows(0)("Surface").ToString
            End If

            If dtPlatform.Rows(0)("PriorityPlan").ToString = "" Then
                cmbPriority.SelectedValue = "-1"
            Else
                cmbPriority.SelectedValue = dtPlatform.Rows(0)("PriorityPlan").ToString
            End If


            If dtPlatform.Rows(0)("Location").ToString = "" Then
                cmbLocation.SelectedValue = "-1"
            Else
                cmbLocation.SelectedValue = dtPlatform.Rows(0)("Location").ToString
            End If

            txtCommentsPlanned.Text = dtPlatform.Rows(0)("CommentsPlanned").ToString

            'Opciones adicionadas 29/03/2012

            txtDepth1.Text = dtPlatform.Rows(0)("Depth1").ToString
            txtDepth2.Text = dtPlatform.Rows(0)("Depth2").ToString
            txtDepth3.Text = dtPlatform.Rows(0)("Depth3").ToString
            If dtPlatform.Rows(0)("target1").ToString = "" Then
                cmbBeta1.SelectedValue = "-1"
            Else
                cmbBeta1.SelectedValue = dtPlatform.Rows(0)("target1").ToString
            End If

            If dtPlatform.Rows(0)("target2").ToString = "" Then
                cmbBeta2.SelectedValue = "-1"
            Else
                cmbBeta2.SelectedValue = dtPlatform.Rows(0)("target2").ToString
            End If

            If dtPlatform.Rows(0)("target3").ToString = "" Then
                cmbBeta3.SelectedValue = "-1"
            Else
                cmbBeta3.SelectedValue = dtPlatform.Rows(0)("target3").ToString
            End If

            If dtPlatform.Rows(0)("Orientation1").ToString = "" Then
                cmbZone1.SelectedValue = "-1"
            Else
                cmbZone1.SelectedValue = dtPlatform.Rows(0)("Orientation1").ToString
            End If
            If dtPlatform.Rows(0)("Orientation2").ToString = "" Then
                cmbZone2.SelectedValue = "-1"
            Else
                cmbZone2.SelectedValue = dtPlatform.Rows(0)("Orientation2").ToString
            End If
            If dtPlatform.Rows(0)("Orientation3").ToString = "" Then
                cmbZone3.SelectedValue = "-1"
            Else
                cmbZone3.SelectedValue = dtPlatform.Rows(0)("Orientation3").ToString
            End If



            'Drilling
            txtHoleIDDrill1.Text = dtPlatform.Rows(0)("HoleID").ToString
            txtEOHDrill1.Text = dtPlatform.Rows(0)("EOH").ToString
            If dtPlatform.Rows(0)("StartDate").ToString = "" Then
                dtpStartDateDrill1.Value = "01/01/1900"
            Else
                dtpStartDateDrill1.Value = dtPlatform.Rows(0)("StartDate").ToString
            End If
            If dtPlatform.Rows(0)("FinalDate").ToString = "" Then
                dtpEndDateDrill1.Value = "01/01/1900"
            Else
                dtpEndDateDrill1.Value = dtPlatform.Rows(0)("FinalDate").ToString
            End If
            cmbRigUsedDrill1.SelectedValue = dtPlatform.Rows(0)("RigUsed").ToString
            txtRodLostDrill1.Text = dtPlatform.Rows(0)("RodLost").ToString
            txtCasingDrill1.Text = dtPlatform.Rows(0)("Casing").ToString
            txtLocationPlatformCompanyService.Text = dtPlatform.Rows(0)("Location").ToString
            txtEastCS.Text = dtPlatform.Rows(0)("EastCS").ToString
            txtNorthCS.Text = dtPlatform.Rows(0)("NorthCS").ToString
            txtElevationCS.Text = dtPlatform.Rows(0)("ElevationCS").ToString
            If dtPlatform.Rows(0)("DateCS").ToString = "" Then
                dtpDateCS.Value = "01/01/1900"
            Else
                dtpDateCS.Value = dtPlatform.Rows(0)("DateCS").ToString
            End If
            If dtPlatform.Rows(0)("Contractor").ToString = "" Then
                cmbContractorDrill1.SelectedValue = "-1"
            Else
                cmbContractorDrill1.SelectedValue = dtPlatform.Rows(0)("Contractor").ToString
            End If

            'Topografía
            txtEastGpsTopo.Text = dtPlatform.Rows(0)("EastGPS").ToString
            txtNorthGpsTopo.Text = dtPlatform.Rows(0)("NorthGPS").ToString
            txtElevationGpsTopo.Text = dtPlatform.Rows(0)("ElevationGPS").ToString

            If dtPlatform.Rows(0)("Surveryor").ToString = "" Then
                cmbSurveryor.SelectedValue = "-1"
            Else
                cmbSurveryor.SelectedValue = dtPlatform.Rows(0)("Surveryor").ToString
            End If
            If dtPlatform.Rows(0)("SurveryorSt").ToString = "" Then
                cmbSurveryorST.SelectedValue = "-1"
            Else
                cmbSurveryorST.SelectedValue = dtPlatform.Rows(0)("SurveryorST").ToString
            End If

            If dtPlatform.Rows(0)("DateGPS").ToString = "" Then
                dtpDateGpsTopo.Value = "01/01/1900"
            Else
                dtpDateGpsTopo.Value = dtPlatform.Rows(0)("DateGPS").ToString
            End If
            txtEastStTopo.Text = dtPlatform.Rows(0)("EastSt").ToString
            txtNorthStTopo.Text = dtPlatform.Rows(0)("NorthSt").ToString
            txtElevationStTopo.Text = dtPlatform.Rows(0)("ElevationSt").ToString
            If dtPlatform.Rows(0)("DateSt").ToString = "" Then
                dtpDateStTopo.Value = "01/01/1900"
            Else
                dtpDateStTopo.Value = dtPlatform.Rows(0)("DateSt").ToString
            End If


            If dtPlatform.Rows(0)("CompanyService").ToString = "" Then
                cmdCompanyService.SelectedValue = "-1"
            Else
                cmdCompanyService.SelectedValue = dtPlatform.Rows(0)("CompanyService").ToString
            End If
            txtLocationPlatformCompanyService.Text = dtPlatform.Rows(0)("location").ToString
            txtEastCS.Text = dtPlatform.Rows(0)("EastCS").ToString
            txtNorthCS.Text = dtPlatform.Rows(0)("NorthCS").ToString
            txtElevationCS.Text = dtPlatform.Rows(0)("ElevationCS").ToString
            If dtPlatform.Rows(0)("DateCS").ToString = "" Then
                dtpDateCS.Value = "01/01/1900"
            Else
                dtpDateCS.Value = dtPlatform.Rows(0)("DateCS").ToString
            End If

            txtCommentsTopo.Text = dtPlatform.Rows(0)("CommentsTopo").ToString

            'LandPermit
            If dtPlatform.Rows(0)("LandPermitStatus").ToString = "" Then
                cmbLandPermitStatus.SelectedValue = "-1"
            Else
                cmbLandPermitStatus.SelectedValue = dtPlatform.Rows(0)("LandPermitStatus").ToString
            End If
            txtIdCatastralLandPermit.Text = dtPlatform.Rows(0)("CatastralFolioID").ToString
            txtLandOwnerLandPermit.Text = dtPlatform.Rows(0)("LandOwner").ToString
            If dtPlatform.Rows(0)("LandPermit").ToString = "" Then
                cmbLandPermit.SelectedValue = "Select an option.."
            Else
                cmbLandPermit.SelectedValue = dtPlatform.Rows(0)("LandPermit").ToString
            End If
            txtCommentLandPermit.Text = dtPlatform.Rows(0)("CommentsLand").ToString

            'Environment
            If dtPlatform.Rows(0)("progress").ToString = "" Then
                'cmbProgress.SelectedValue = "-1"
            Else
                ' cmbProgress.SelectedValue = dtPlatform.Rows(0)("progress").ToString
            End If
            ' txtRecordEnvironment.Text = dtPlatform.Rows(0)("Environment").ToString


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub FillHoleID()
        Try
            Dim dtCollars As DataTable = oCollars.getDHCollars
            'Dim drC As DataRow = dtCollars.NewRow()
            'MsgBox(dtCollars.Rows(0)("PlatFormID"))
            txtPlatform.Text = dtCollars.Rows(0)("PlatFormID").ToString
            txtHoleID.Text = dtCollars.Rows(0)("HoleID").ToString
            txtSection.Text = dtCollars.Rows(0)("Section").ToString
            txtEastPlanned.Text = dtCollars.Rows(0)("EastPlanned").ToString
            txtNorthPlanned.Text = dtCollars.Rows(0)("NorthPlanned").ToString
            txtElevationPlanned.Text = dtCollars.Rows(0)("ElevationPlanned").ToString
            txtAzPlanned.Text = dtCollars.Rows(0)("AzimuthPlanned").ToString
            txtInclinationPlanned.Text = dtCollars.Rows(0)("InclinationPlanned").ToString
            txtLengthPlanned.Text = dtCollars.Rows(0)("LengthPlanned").ToString
            If dtCollars.Rows(0)("ToRelocate").ToString = "" Then
                cmbToRelocateD.SelectedValue = "Select an option.."
            Else
                cmbToRelocateD.SelectedValue = dtCollars.Rows(0)("ToRelocate").ToString
            End If
            If dtCollars.Rows(0)("Status").ToString = "" Then
                cmbStatusPlatform.SelectedValue = "-1"
            Else
                cmbStatusPlatform.SelectedValue = dtCollars.Rows(0)("Status").ToString
            End If

            If dtCollars.Rows(0)("Surface").ToString = "" Then
                cmbSurface.SelectedValue = "-1"
            Else
                cmbSurface.SelectedValue = dtCollars.Rows(0)("Surface").ToString
            End If
            If dtCollars.Rows(0)("PriorityPlan").ToString = "" Then
                cmbPriority.SelectedValue = "-1"
            Else
                cmbPriority.SelectedValue = dtCollars.Rows(0)("PriorityPlan").ToString
            End If

            If dtCollars.Rows(0)("Location").ToString = "" Then
                cmbLocation.SelectedValue = "-1"
            Else
                cmbLocation.SelectedValue = dtCollars.Rows(0)("Location").ToString
            End If

            txtCommentsPlanned.Text = dtCollars.Rows(0)("CommentsPlanned").ToString


            'Opciones adicionadas 29/03/2012

            txtDepth1.Text = dtCollars.Rows(0)("Depth1").ToString
            txtDepth2.Text = dtCollars.Rows(0)("Depth2").ToString
            txtDepth3.Text = dtCollars.Rows(0)("Depth3").ToString
            If dtCollars.Rows(0)("target1").ToString = "" Then
                cmbBeta1.SelectedValue = "-1"
            Else
                cmbBeta1.SelectedValue = dtCollars.Rows(0)("target1").ToString
            End If

            If dtCollars.Rows(0)("target2").ToString = "" Then
                cmbBeta2.SelectedValue = "-1"
            Else
                cmbBeta2.SelectedValue = dtCollars.Rows(0)("target2").ToString
            End If

            If dtCollars.Rows(0)("target3").ToString = "" Then
                cmbBeta3.SelectedValue = "-1"
            Else
                cmbBeta3.SelectedValue = dtCollars.Rows(0)("target3").ToString
            End If

            If dtCollars.Rows(0)("Orientation1").ToString = "" Then
                cmbZone1.SelectedValue = "-1"
            Else
                cmbZone1.SelectedValue = dtCollars.Rows(0)("Orientation1").ToString
            End If

            If dtCollars.Rows(0)("Orientation2").ToString = "" Then
                cmbZone2.SelectedValue = "-1"
            Else
                cmbZone2.SelectedValue = dtCollars.Rows(0)("Orientation2").ToString
            End If

            If dtCollars.Rows(0)("Orientation3").ToString = "" Then
                cmbZone3.SelectedValue = "-1"
            Else
                cmbZone3.SelectedValue = dtCollars.Rows(0)("Orientation3").ToString
            End If

            'If dtCollars.Rows(0)("Target").ToString = "" Then
            '    cmbLocation.SelectedValue = "-1"
            'Else
            '    cmbLocation.SelectedValue = dtCollars.Rows(0)("Target").ToString
            'End If

            'Drilling

            txtHoleIDDrill1.Text = dtCollars.Rows(0)("HoleID").ToString
            txtEOHDrill1.Text = dtCollars.Rows(0)("EOH").ToString
            If dtCollars.Rows(0)("StartDate").ToString = "" Then
                dtpStartDateDrill1.Value = "01/01/1900"
            Else
                dtpStartDateDrill1.Value = dtCollars.Rows(0)("StartDate").ToString
            End If
            If dtCollars.Rows(0)("FinalDate").ToString = "" Then
                dtpEndDateDrill1.Value = "01/01/1900"
            Else
                dtpEndDateDrill1.Value = dtCollars.Rows(0)("FinalDate").ToString
            End If
            If dtCollars.Rows(0)("RigUsed").ToString = "" Then
                cmbRigUsedDrill1.SelectedValue = "-1"
            Else
                cmbRigUsedDrill1.SelectedValue = dtCollars.Rows(0)("RigUsed").ToString
            End If

            'txtRigUsedDrill1.Text = dtCollars.Rows(0)("RigUsed").ToString
            If dtCollars.Rows(0)("Contractor").ToString = "" Then
                cmbContractorDrill1.SelectedValue = "Select an opction.."
            Else
                cmbContractorDrill1.SelectedValue = dtCollars.Rows(0)("Contractor").ToString
            End If

            txtRodLostDrill1.Text = dtCollars.Rows(0)("RodLost").ToString
            txtCasingDrill1.Text = dtCollars.Rows(0)("Casing").ToString

            'Topografía
            txtEastGpsTopo.Text = dtCollars.Rows(0)("EastGPS").ToString
            txtNorthGpsTopo.Text = dtCollars.Rows(0)("NorthGPS").ToString
            txtElevationGpsTopo.Text = dtCollars.Rows(0)("ElevationGPS").ToString

            If dtCollars.Rows(0)("Surveryor").ToString = "" Then
                cmbSurveryor.SelectedValue = "-1"
            Else
                cmbSurveryor.SelectedValue = dtCollars.Rows(0)("Surveryor").ToString
            End If
            If dtCollars.Rows(0)("SurveryorSt").ToString = "" Then
                cmbSurveryorST.SelectedValue = "-1"
            Else
                cmbSurveryorST.SelectedValue = dtCollars.Rows(0)("SurveryorST").ToString
            End If

            If dtCollars.Rows(0)("DateGPS").ToString = "" Then
                dtpDateGpsTopo.Value = "01/01/1900"
            Else
                dtpDateGpsTopo.Value = dtCollars.Rows(0)("DateGPS").ToString
            End If
            txtEastStTopo.Text = dtCollars.Rows(0)("EastSt").ToString
            txtNorthStTopo.Text = dtCollars.Rows(0)("NorthSt").ToString
            txtElevationStTopo.Text = dtCollars.Rows(0)("ElevationSt").ToString
            If dtCollars.Rows(0)("DateSt").ToString = "" Then
                dtpDateStTopo.Value = "01/01/1900"
            Else
                dtpDateStTopo.Value = dtCollars.Rows(0)("DateSt").ToString
            End If
            If dtCollars.Rows(0)("CompanyService").ToString = "" Then
                cmdCompanyService.SelectedValue = "-1"
            Else
                cmdCompanyService.SelectedValue = dtCollars.Rows(0)("CompanyService").ToString
            End If
            txtLocationPlatformCompanyService.Text = dtCollars.Rows(0)("location").ToString
            txtEastCS.Text = dtCollars.Rows(0)("EastCS").ToString
            txtNorthCS.Text = dtCollars.Rows(0)("NorthCS").ToString
            txtElevationCS.Text = dtCollars.Rows(0)("ElevationCS").ToString
            If dtCollars.Rows(0)("DateCS").ToString = "" Then
                dtpDateCS.Value = "01/01/1900"
            Else
                dtpDateCS.Value = dtCollars.Rows(0)("DateCS").ToString
            End If

            txtCommentsTopo.Text = dtCollars.Rows(0)("CommentsTopo").ToString

            'LandPermit
            If dtCollars.Rows(0)("LandPermitStatus").ToString = "" Then
                cmbLandPermitStatus.SelectedValue = "-1"
            Else
                cmbLandPermitStatus.SelectedValue = dtCollars.Rows(0)("LandPermitStatus").ToString
            End If
            txtIdCatastralLandPermit.Text = dtCollars.Rows(0)("CatastralFolioID").ToString
            txtLandOwnerLandPermit.Text = dtCollars.Rows(0)("LandOwner").ToString
            If dtCollars.Rows(0)("LandPermit").ToString = "" Then
                cmbLandPermit.SelectedValue = "Select an option.."
            Else
                cmbLandPermit.SelectedValue = dtCollars.Rows(0)("LandPermit").ToString
            End If
            txtCommentLandPermit.Text = dtCollars.Rows(0)("CommentsLand").ToString

            'Environment
            If dtCollars.Rows(0)("progress").ToString = "" Then
                'cmbProgress.SelectedValue = "Select an option.."
            Else
                'cmbProgress.SelectedValue = dtCollars.Rows(0)("progress").ToString
            End If
            'txtRecordEnvironment.Text = dtCollars.Rows(0)("Environment").ToString

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub limpiar()
        txtPlatform.Text = ""
        txtHoleID.Text = ""
        txtSection.Text = ""
        txtEastPlanned.Text = ""
        txtNorthPlanned.Text = ""
        txtElevationPlanned.Text = ""
        txtAzPlanned.Text = ""
        txtInclinationPlanned.Text = ""
        txtLengthPlanned.Text = ""
        txtCommentsPlanned.Text = ""
        txtDepth1.Text = ""
        txtDepth2.Text = ""
        txtDepth3.Text = ""

        'Drilling
        txtHoleIDDrill1.Text = ""
        txtEOHDrill1.Text = ""
        dtpStartDateDrill1.Value = "01/01/1900"
        dtpEndDateDrill1.Value = "01/01/1900"
        'txtRigUsedDrill1.Text = ""
        cmbRigUsedDrill1.SelectedValue = "-1"
        txtRodLostDrill1.Text = ""
        txtCasingDrill1.Text = ""
        dgDrillProgress1.DataSource = ""
        dgHoleInProgress1.DataSource = ""


        'Topo
        txtEastGpsTopo.Text = ""
        txtNorthGpsTopo.Text = ""
        txtElevationGpsTopo.Text = ""
        dtpDateGpsTopo.Value = "01/01/1900"
        txtEastStTopo.Text = ""
        txtNorthStTopo.Text = ""
        txtElevationStTopo.Text = ""
        dtpDateStTopo.Value = "01/01/1900"
        txtLocationPlatformCompanyService.Text = ""
        txtEastCS.Text = ""
        txtNorthCS.Text = ""
        txtElevationCS.Text = ""
        dtpDateCS.Value = "01/01/1900"
        txtCommentsTopo.Text = ""

        'Land Permit
        txtIdCatastralLandPermit.Text = ""
        txtLandOwnerLandPermit.Text = ""
        txtCommentLandPermit.Text = ""

        'Environment
        'txtRecordEnvironment.Text = ""



        CallCmb()
        'dgPozos.DataSource = ""
        'dgHolePlatform.DataSource = ""
    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frmDrilling_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Show()

        CallCmb()

        'For Each file As FileInfo In folder.GetFiles()
        '    lst_FilesEnvironment.Items.Add(file.Name)
        'Next

        'For Each file As FileInfo In folder.GetFiles()
        '    lstPdfFile.Items.Add(file.Name)
        'Next

        If optPlatform.Checked Then
            FillPlatformIDForm()
        End If
        If optHole.Checked Then
            FillHoleIDForm()
        End If

        Dim drSeg As DataRow()
        drSeg = clsRf.dsPermisos.Tables(1).Select("nombre_Real_Form = 'frmDrillingPl' and Accion = 'Insertar'")
        If drSeg.Length > 0 Then
            btnAddPlanned.Enabled = True
            btnNew.Enabled = True
        Else
            btnAddPlanned.Enabled = False
            btnNew.Enabled = False
        End If
        drSeg = clsRf.dsPermisos.Tables(1).Select("nombre_Real_Form = 'frmDrillingDr' and Accion = 'Insertar'")
        If drSeg.Length > 0 Then
            btnAddDrill1.Enabled = True
            btnAddProgress1.Enabled = True
            btnNewDC.Enabled = True
            dgCompanyDrill.Enabled = True
        Else
            btnAddDrill1.Enabled = False
            btnAddProgress1.Enabled = False
            btnNewDC.Enabled = False
            dgCompanyDrill.Enabled = False
        End If

        drSeg = clsRf.dsPermisos.Tables(1).Select("nombre_Real_Form = 'frmDrillingTo' and Accion = 'Insertar'")
        If drSeg.Length > 0 Then
            btnAddTopo.Enabled = True
            btnSelectPTopo.Enabled = True
            btnAddImg.Enabled = True
        Else
            btnAddTopo.Enabled = False
            btnSelectPTopo.Enabled = False
            btnAddImg.Enabled = False
        End If
        drSeg = clsRf.dsPermisos.Tables(1).Select("nombre_Real_Form = 'frmDrillingLa' and Accion = 'Insertar'")
        If drSeg.Length > 0 Then
            btnAddLandPermit.Enabled = True
            btnSelectFileLp.Enabled = True
            btnAddFileLp.Enabled = True
        Else
            btnAddLandPermit.Enabled = False
            btnSelectFileLp.Enabled = False
            btnAddFileLp.Enabled = False
        End If

        drSeg = clsRf.dsPermisos.Tables(1).Select("nombre_Real_Form = 'frmDrillingEn' and Accion = 'Insertar'")
        If drSeg.Length > 0 Then
            'btnAddEnvironment.Enabled = True
            'btnAddFileEnvironment.Enabled = True
            'btnSelectFileEnvironment.Enabled = True
        Else
            'btnAddEnvironment.Enabled = False
            'btnAddFileEnvironment.Enabled = False
            'btnSelectFileEnvironment.Enabled = False
        End If

        'FillHoleStructure()
        tabOpciones.TabPages.Remove(Me.TabLand)
        'tabOpciones.TabPages.Remove(Me.TabEnv)
        'TabControl1.TabPages.Remove(Me.TabPage2)
        tbcOpciones.TabPages.Remove(Me.Tab1)
        tbcOpciones.TabPages.Remove(Me.TabPage2)
        tbcOpciones.TabPages.Remove(Me.TabPage4)
        tbcOpciones.TabPages.Remove(Me.TabPage5)
        tbcOpciones.TabPages.Remove(Me.TabPage7)
        tbcOpciones.TabPages.Remove(Me.TabPage10)
        tbcOpciones.TabPages.Remove(Me.TabPage6)

        FillCompanyDrill()

    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim ruta As String
        'ruta = "C:\" & lst_FilesEnvironment.SelectedItem
        'System.Diagnostics.Process.Start(ruta)
        'MsgBox(ListBox1.SelectedItem)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CallCmb()
        Try
            llenarPorcentaje()

            FillTorelocate()
            FillStatusPlatform()

            'FillZone1()
            'FillZone2()
            'FillZone3()

            FillSurface()
            FillPriority()
            FillContractor()
            FillSurveryor()
            FillSurveryorSt()
            FillCompanyServices()
            FillLandPermitStatus()
            FillLandPermit()
            'FillPercentProgress()
            FillCoreDiameter()
            FillLocation()

            FillBeta1()
            FillBeta2()
            FillBeta3()


            'FillTarguet()

            FillHoleIDCD()
            FillRigDill()
            'Company Drill
            FillRig()
            FillTur()
            FillHoleDownTCD()
            'FillChangeCrownCD()
            'FillTurnSuppliesCD()
            FillLostToolsCD()
            FillBillableAdditivesCD()
            'FillBiabilityContractorCD()
            'FillBiabilityCompCD()

            FillGroup()
            FillImpact()
            'FillSubGroup()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub
    Private Sub cmbHoleId_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbHoleId.SelectionChangeCommitted
        lblErrorFt.Visible = False
        limpiar()
        txtPlatform.Enabled = True
        If optPlatform.Checked Then
            'oPlatform.sPlatformID = cmbHoleId.SelectedValue.ToString
            oPlatform.sPlatformID = cmbHoleId.SelectedValue.ToString
            FillPlatformID()
            'FillPlatformIDHold()
            txtPlatform.Text = cmbHoleId.SelectedValue.ToString
            sEdit = "1"
            txtPlatform.Enabled = False

        End If
        If optHole.Checked Then
            oCollars.sHoleID = cmbHoleId.SelectedValue.ToString
            FillHoleID()
            sEdit = "1"
            txtPlatform.Enabled = False
        End If

        FillDrillProgress()
        FillHoleInProgress()
        FillCompanyDrill()

        FillMaxTo()
        ValidaFtAvanceDiario()
        ListarImgTopo()
        ListarPDF()
        FillEnvironmentPoll()
        ListarPictureEnv()

    End Sub
    Private Function ValidaFtAvanceDiario() As Boolean
        Try
            Dim i As Integer
            'Dim a As Decimal
            'Dim C As Decimal
            'Validar Litho
            Dim r, g, b As Integer
            r = "255"
            g = "0"
            b = "0"

            'MsgBox(r & g & b)

            Color.FromArgb(r, g, b)

            'MsgBox(r & g & b)

            'MsgBox(Color.FromArgb(r, g, b))

            For i = 0 To dgDrillProgress1.Rows.Count - 1
                'a = dg_Datos.Rows(i).Cells(1).Value
                'C = dg_Datos.Rows(i).Cells(2).Value
                If i > 0 Then
                    If dgDrillProgress1.Rows(i).Cells(1).Value <> dgDrillProgress1.Rows(i - 1).Cells(2).Value Then
                        If dgDrillProgress1.Rows(i).Cells(0).Value <> "" Then
                            Dim cell1 As DataGridViewCell = dgDrillProgress1.Rows(i - 1).Cells(2)
                            Dim cell As DataGridViewCell = dgDrillProgress1.Rows(i).Cells(1)
                            cell1.Style.BackColor = Color.FromArgb(r, g, b)
                            cell.Style.BackColor = Color.FromArgb(r, g, b)
                            lblErrorFt.Visible = True
                        End If
                    End If
                End If
            Next
            dgDrillProgress1.Refresh()
        Catch ex As Exception
        End Try
    End Function
    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPlanned.Click

    End Sub

    Private Sub tabOpciones_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tabOpciones.KeyDown
        If e.KeyValue = Keys.Enter Then
            SendKeys.Send("{tab}")
        End If
    End Sub


    Private Sub cmbHoleId_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHoleId.SelectedIndexChanged

    End Sub

    Private Sub dgPozos_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub dgPozos_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

        'oCollars.sHoleID = dgPozos.Rows(e.RowIndex).Cells(0).Value.ToString()
        'FillHoleID()

    End Sub

    Private Sub txtCommentLandPermit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCommentLandPermit.TextChanged

    End Sub

    Private Sub cmbToRelocateD_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbToRelocateD.SelectedIndexChanged
    End Sub

    Private Sub cmbToRelocateD_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbToRelocateD.SelectionChangeCommitted
        'MsgBox(cmbToRelocateD.SelectedValue.ToString)
    End Sub

    Private Sub txtEastPlanned_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEastPlanned.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtEastPlanned_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEastPlanned.TextChanged

    End Sub

    Private Sub txtNorthPlanned_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNorthPlanned.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNorthPlanned_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNorthPlanned.TextChanged

    End Sub

    Private Sub txtElevationPlanned_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtElevationPlanned.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtElevationPlanned_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtElevationPlanned.TextChanged

    End Sub

    Private Sub txtAzPlanned_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAzPlanned.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtAzPlanned_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAzPlanned.TextChanged

    End Sub

    Private Sub txtInclinationPlanned_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInclinationPlanned.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtInclinationPlanned_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInclinationPlanned.TextChanged

    End Sub

    Private Sub txtLengthPlanned_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLengthPlanned.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtLengthPlanned_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLengthPlanned.TextChanged

    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            limpiar()
            txtPlatform.Focus()
            txtPlatform.Enabled = True
            sEdit = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
        

    End Sub

    Private Sub ValidarVacios()
        Try

            sValidar = True
            If txtPlatform.Text = "" Then
                sValidar = False
                MsgBox("Error in Platform", MsgBoxStyle.Critical)
                txtPlatform.Focus()
                Exit Sub
            End If
            If txtSection.Text = "" Then
                sValidar = False
                MsgBox("Error in Section", MsgBoxStyle.Critical)
                txtSection.Focus()
                Exit Sub
            End If
            If txtEastPlanned.Text = "" Then
                sValidar = False
                MsgBox("Error in East Planned", MsgBoxStyle.Critical)
                txtEastPlanned.Focus()
                Exit Sub
            End If
            If txtNorthPlanned.Text = "" Then
                sValidar = False
                MsgBox("Error in North Planned", MsgBoxStyle.Critical)
                txtNorthPlanned.Focus()
                Exit Sub
            End If
            If txtElevationPlanned.Text = "" Then
                sValidar = False
                MsgBox("Error in Elevation Planned", MsgBoxStyle.Critical)
                txtElevationPlanned.Focus()
                Exit Sub
            End If
            If txtAzPlanned.Text = "" Then
                sValidar = False
                MsgBox("Error in Azimuth Planned", MsgBoxStyle.Critical)
                txtAzPlanned.Focus()
                Exit Sub
            End If
            If txtInclinationPlanned.Text = "" Then
                sValidar = False
                MsgBox("Error in Inclination Planned", MsgBoxStyle.Critical)
                txtInclinationPlanned.Focus()
                Exit Sub
            End If
            If txtLengthPlanned.Text = "" Then
                sValidar = False
                MsgBox("Error in Length Planned", MsgBoxStyle.Critical)
                txtLengthPlanned.Focus()
                Exit Sub
            End If
            'If txtLengthPlanned.Text <> "" Then
            '    If txtLengthPlanned.Text < 1 And txtLengthPlanned.Text > 1200 Then
            '        sValidar = False
            '        MsgBox("Error in Length, >1 and < 1200", MsgBoxStyle.Critical)
            '        txtLengthPlanned.Focus()
            '        Exit Sub
            '    End If
            'End If
            If txtAzPlanned.Text < 0 Or txtAzPlanned.Text > 360 Then
                sValidar = False
                MsgBox("Error in Azimuth Planned, range (0 to 360", MsgBoxStyle.Critical)
                txtAzPlanned.Focus()
                Exit Sub
            End If
            If txtInclinationPlanned.Text < -90 Or txtInclinationPlanned.Text > 90 Then
                sValidar = False
                MsgBox("Error in Inclination Planned, range (-90 and 90)", MsgBoxStyle.Critical)
                txtInclinationPlanned.Focus()
                Exit Sub
            End If

            If txtElevationPlanned.Text < -200 Or txtElevationPlanned.Text > 1000 Then
                sValidar = False
                MsgBox("Error in Elevation Planned, range (-200 and 1000)", MsgBoxStyle.Critical)
                txtElevationPlanned.Focus()
                Exit Sub
            End If

            If cmbLocation.SelectedValue = "-1" Then
                sValidar = False
                MsgBox("Error in Location", MsgBoxStyle.Critical)
                cmbLocation.Focus()
                Exit Sub
            End If

            'If cmbZone1.SelectedValue.ToString = "-1" Or cmbZone1.SelectedText.ToString = "" Then
            '    sValidar = False
            '    MsgBox("Select a Orientation 1", MsgBoxStyle.Critical)
            '    cmbZone1.Focus()
            '    Exit Sub
            'End If
            'If cmbZone2.SelectedValue.ToString = "-1" Then
            '    sValidar = False
            '    MsgBox("Select a Orientation 2", MsgBoxStyle.Critical)
            '    cmbZone1.Focus()
            '    Exit Sub
            'End If
            'If cmbZone3.SelectedValue.ToString = "-1" Then
            '    sValidar = False
            '    MsgBox("Select a Orientation 3", MsgBoxStyle.Critical)
            '    cmbZone1.Focus()
            '    Exit Sub
            'End If

            If cmbSurface.SelectedValue.ToString = "-1" Then
                sValidar = False
                MsgBox("Select a Surface", MsgBoxStyle.Critical)
                cmbSurface.Focus()
                Exit Sub
            End If

            If txtDepth1.Text = "" And cmbBeta1.SelectedValue <> "-1" Then
                sValidar = False
                MsgBox("Error in Expected Cut 1", MsgBoxStyle.Critical)
                txtDepth1.Focus()
                Exit Sub
            End If

            'If txtDepth1.Text <> "" And (cmbBeta1.SelectedValue = "-1" Or cmbZone1.SelectedValue = "-1") Then
            '    sValidar = False
            '    MsgBox("Error in Vein 1", MsgBoxStyle.Critical)
            '    cmbBeta1.Focus()
            '    Exit Sub
            'End If

            If txtDepth1.Text <> "" And cmbBeta1.SelectedValue = "-1" Then
                sValidar = False
                MsgBox("Error in Vein 1", MsgBoxStyle.Critical)
                cmbBeta1.Focus()
                Exit Sub
            End If

            If txtDepth2.Text = "" And cmbBeta2.SelectedValue <> "-1" Then
                sValidar = False
                MsgBox("Error in Expected Cut 2", MsgBoxStyle.Critical)
                txtDepth2.Focus()
                Exit Sub
            End If

            If txtDepth2.Text <> "" And cmbBeta2.SelectedValue = "-1" Then
                sValidar = False
                MsgBox("Error in Vein 2", MsgBoxStyle.Critical)
                cmbBeta2.Focus()
                Exit Sub
            End If

            If txtDepth3.Text = "" And cmbBeta3.SelectedValue <> "-1" Then
                sValidar = False
                MsgBox("Error in Expected Cut 3", MsgBoxStyle.Critical)
                txtDepth3.Focus()
                Exit Sub
            End If

            If txtDepth3.Text <> "" And cmbBeta3.SelectedValue = "-1" Then
                sValidar = False
                MsgBox("Error in Vein 3", MsgBoxStyle.Critical)
                cmbBeta3.Focus()
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
        


    End Sub

    Private Sub btnAddPlanned_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPlanned.Click
        Try
            ValidarVacios()
            If sValidar = True Then
                'Hago Insercion en la tabla Platform
                If sEdit = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sPlatform = txtPlatform.Text.ToString()
                    oPlatform.sSection = txtSection.Text.ToString()
                    oPlatform.sEastPlanned = txtEastPlanned.Text.ToString()
                    oPlatform.sNothPlanned = txtNorthPlanned.Text.ToString()
                    oPlatform.sElevationPlanned = txtElevationPlanned.Text.ToString()
                    oPlatform.sAzimuthPlanned = txtAzPlanned.Text.ToString()
                    oPlatform.sInclinationPlanned = txtInclinationPlanned.Text.ToString()
                    oPlatform.sLengthPlanned = txtLengthPlanned.Text.ToString()
                    oPlatform.sTorelocate = cmbToRelocateD.SelectedValue.ToString()
                    oPlatform.sStatus = cmbStatusPlatform.SelectedValue.ToString()
                    oPlatform.sZone = cmbLocation.SelectedValue.ToString()
                    oPlatform.sSurface = cmbSurface.SelectedValue.ToString()
                    oPlatform.sPriorityPlan = cmbPriority.SelectedValue.ToString()
                    oPlatform.sCommentsPlanned = txtCommentsPlanned.Text.ToString()
                    oPlatform.sEdit = 0

                    'Opciones adicionadas el 29/03/2012
                    oPlatform.sDepth1 = txtDepth1.Text.ToString
                    oPlatform.sDepth2 = txtDepth2.Text.ToString
                    oPlatform.sDepth3 = txtDepth3.Text.ToString

                    oPlatform.sBeta1 = cmbBeta1.SelectedValue.ToString
                    oPlatform.sBeta2 = cmbBeta2.SelectedValue.ToString
                    oPlatform.sBeta3 = cmbBeta3.SelectedValue.ToString

                    'oPlatform.sOrientation1 = cmbZone1.SelectedValue.ToString
                    'oPlatform.sOrientation2 = cmbZone2.SelectedValue.ToString
                    'oPlatform.sOrientation3 = cmbZone3.SelectedValue.ToString

                    oPlatform.sOrientation1 = ""
                    oPlatform.sOrientation2 = ""
                    oPlatform.sOrientation3 = ""


                    'oPlatform.sTarguet = cmbLocation.SelectedValue.ToString

                    Dim sResp As String = oPlatform.DH_Platform_Planned_Add()

                    If sResp = "OK" Then
                        oRf.InsertTrans("DH_Platform", "Insert", clsRf.sUser.ToString(), _
                        "Platform : " + txtPlatform.Text.ToString() + ". " + _
                        "East Planned: " + txtEastPlanned.Text.ToString() + ". " + _
                        "North Planned: " + txtNorthPlanned.Text.ToString() + ". " + _
                        "Elevation Planned: " + txtElevationPlanned.Text.ToString() + ". " + _
                        "Azimuth Planned: " + txtAzPlanned.Text.ToString() + ". " + _
                        "Inclination Planne: " + txtInclinationPlanned.Text.ToString() + ". " + _
                        "Event Date " + Date.Now())
                        MsgBox("Platform Planned Save.", MsgBoxStyle.Information)
                    End If
                    limpiar()
                Else
                    'input box
                    sMotivo = InputBox("Motivo de la Actualización (Solo se guardara el registro en historico si ingresa un motivo)", "Actualización", "")
                    If sMotivo <> "" Then
                        oPlatform.sPlatform = txtPlatform.Text.ToString
                        oPlatform.sCommentsHistory = sMotivo
                        oPlatform.sUser = clsRf.sUser
                        'Dim sResp As String = 
                        oPlatform.DH_Platform_History_Add()
                    Else
                        Exit Sub
                    End If

                    oPlatform.sOpcion = "2"
                    oPlatform.sPlatform = txtPlatform.Text.ToString
                    oPlatform.sSection = txtSection.Text.ToString
                    oPlatform.sEastPlanned = txtEastPlanned.Text.ToString()
                    oPlatform.sNothPlanned = txtNorthPlanned.Text.ToString()
                    oPlatform.sElevationPlanned = txtElevationPlanned.Text.ToString()
                    oPlatform.sAzimuthPlanned = txtAzPlanned.Text.ToString()
                    oPlatform.sInclinationPlanned = txtInclinationPlanned.Text.ToString
                    oPlatform.sLengthPlanned = txtLengthPlanned.Text.ToString()
                    oPlatform.sTorelocate = cmbToRelocateD.SelectedValue.ToString
                    'validacion temporal
                    If cmbStatusPlatform.SelectedValue.ToString = "" Then
                        'MsgBox(cmbStatusPlatform.SelectedValue.ToString)
                        oPlatform.sStatus = "-1"
                    Else
                        oPlatform.sStatus = cmbStatusPlatform.SelectedValue.ToString
                    End If
                    oPlatform.sZone = cmbLocation.SelectedValue.ToString
                    oPlatform.sSurface = cmbSurface.SelectedValue.ToString
                    oPlatform.sPriorityPlan = cmbPriority.SelectedValue.ToString
                    oPlatform.sCommentsPlanned = txtCommentsPlanned.Text.ToString


                    'Opciones adicionadas el 29/03/2012
                    oPlatform.sDepth1 = txtDepth1.Text.ToString
                    oPlatform.sDepth2 = txtDepth2.Text.ToString
                    oPlatform.sDepth3 = txtDepth3.Text.ToString

                    oPlatform.sBeta1 = cmbBeta1.SelectedValue.ToString
                    oPlatform.sBeta2 = cmbBeta2.SelectedValue.ToString
                    oPlatform.sBeta3 = cmbBeta3.SelectedValue.ToString

                    oPlatform.sOrientation1 = "-1"
                    oPlatform.sOrientation2 = "-1"
                    oPlatform.sOrientation3 = "-1"

                    oPlatform.sEdit = 1
                    'oPlatform.sTarguet = cmbLocation.SelectedValue.ToString

                    Dim sResp As String = oPlatform.DH_Platform_Planned_Add()

                    If sResp = "OK" Then
                        oRf.InsertTrans("DH_Platform", "Update", clsRf.sUser.ToString(), _
                        "Platform : " + txtPlatform.Text.ToString() + ". " + _
                        "East Planned: " + txtEastPlanned.Text.ToString() + ". " + _
                        "North Planned: " + txtNorthPlanned.Text.ToString() + ". " + _
                        "Elevation Planned: " + txtElevationPlanned.Text.ToString() + ". " + _
                        "Azimuth Planned: " + txtAzPlanned.Text.ToString() + ". " + _
                        "Inclination Planne: " + txtInclinationPlanned.Text.ToString() + ". " + _
                        "Event Date " + Date.Now())
                        MsgBox("Platform Planned Updated.", MsgBoxStyle.Information)
                    End If
                    sEdit = "1"
                End If
                'limpiar()
            End If
            'txtPlatform.Enabled = True
            txtPlatform.Focus()
            FillPlatformIDForm()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub btnAddHoleID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim sValida As Boolean = True
            If txtHoleID.Text = "" Then
                MsgBox("Error in Hole ID", MsgBoxStyle.Critical)
                txtHoleID.Focus()
                sValida = False
            End If
            If txtPlatform.Text = "" Then
                MsgBox("Error in Platform", MsgBoxStyle.Critical)
                txtPlatform.Focus()
                sValida = False
            End If
            If sValida = True Then
                oPlatform.sHoleID = txtHoleID.Text.ToString
                oPlatform.sPlatform = txtPlatform.Text.ToString

                Dim sResp As String = oPlatform.DH_Collar_Platform_Add()
                If sResp = "OK" Then
                    MsgBox("Hole Save.")
                    'FillHoleIDPlatform()
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub optPlatform_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPlatform.CheckedChanged
        Try
            cmbHoleId.DisplayMember = ""
            FillPlatformIDForm()
            cmbHoleId.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
       
    End Sub

    Private Sub optHole_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optHole.CheckedChanged
        Try

            cmbHoleId.DisplayMember = ""
            FillHoleIDForm()
            cmbHoleId.Focus()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub validardrilling()
        Try
            sValidaDr = True
            Dim Cadena As String

            Cadena = Mid(txtHoleIDDrill1.Text, 1, 3)
            'MsgBox(Cadena)


            If txtHoleIDDrill1.Text = "" Then
                sValidaDr = False
                MsgBox("Error in HoleID", MsgBoxStyle.Critical)
                txtHoleIDDrill1.Focus()
                Exit Sub
            End If
            'MsgBox(dtpStartDateDrill.Value)
            If txtHoleIDDrill1.Text <> "" And dtpStartDateDrill1.Value = "01/01/1900" Then
                sValidaDr = False
                MsgBox("Error in StartDate, if HoleID = Null", MsgBoxStyle.Critical)
                dtpStartDateDrill1.Focus()
                Exit Sub
            End If

            If txtEOHDrill1.Text <> "" And dtpEndDateDrill1.Value = "01/01/1900" Then
                sValidaDr = False
                MsgBox("Error in EndDate, if OEH <> Null", MsgBoxStyle.Critical)
                txtEOHDrill1.Focus()
                Exit Sub
            End If

            If dtpEndDateDrill1.Value <> "01/01/1900" And txtEOHDrill1.Text = "" Then
                sValidaDr = False
                MsgBox("EndDate and EOH, must be entered at the same time", MsgBoxStyle.Critical)
                txtEOHDrill1.Focus()
                Exit Sub
            End If

            If cmbContractorDrill1.SelectedValue = "-1" Then
                sValidaDr = False
                MsgBox("Error in Contractor", MsgBoxStyle.Critical)
                cmbContractorDrill1.Focus()
                Exit Sub
            End If
            If cmbRigUsedDrill1.SelectedValue = "-1" Then
                sValidaDr = False
                MsgBox("Error in Rig Used", MsgBoxStyle.Critical)
                cmbRigUsedDrill1.Focus()
                Exit Sub
            End If

            If txtEOHDrill1.Text <> "" Then
                If txtEOHDrill1.Text > 2 Then
                    If txtEOHDrill1.Text <> "" And dtpEndDateDrill1.Value = "01/01/1900" Then
                        sValidaDr = False
                        MsgBox("Error in End Date", MsgBoxStyle.Critical)
                        dtpEndDateDrill1.Focus()
                        Exit Sub
                    End If
                Else
                    sValidaDr = False
                    MsgBox("Error in EOH", MsgBoxStyle.Critical)
                    txtEOHDrill1.Focus()
                    Exit Sub
                End If
            End If
            If txtPlatform.Text = "" Then
                sValidaDr = False
                MsgBox("Select a Platform", MsgBoxStyle.Critical)
                txtPlatform.Focus()
                Exit Sub
            End If

            If txtEOHDrill1.Text <> "" Then
                If dtpEndDateDrill1.Value < dtpStartDateDrill1.Value Then
                    sValidaDr = False
                    MsgBox("End Date < Start Date", MsgBoxStyle.Critical)
                    dtpEndDateDrill1.Focus()
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub ValidarDrillProgress()
        Try
            sValidaPr = True
            If txtHoleIDDrill1.Text = "" Then
                sValidaPr = False
                MsgBox("Error in Hole ID", MsgBoxStyle.Critical)
                txtHoleIDDrill1.Focus()
                Exit Sub
            End If
            If txtFrom1.Text = "" Then
                sValidaPr = False
                MsgBox("Error in From", MsgBoxStyle.Critical)
                txtFrom1.Focus()
                Exit Sub
            End If
            If txtTo1.Text = "" Then
                sValidaPr = False
                MsgBox("Error in To", MsgBoxStyle.Critical)
                txtTo1.Focus()
                Exit Sub
            End If
            'MsgBox
            If Val(txtFrom1.Text.ToString) > Val(txtTo1.Text.ToString) Then
                sValidaPr = False
                MsgBox("Error, From > To", MsgBoxStyle.Critical)
                txtTo1.Focus()
                Exit Sub
            End If
            If dtpDateProgress1.Text = "01/01/1900" Then
                sValidaPr = False
                MsgBox("Date Invalid", MsgBoxStyle.Critical)
                dtpDateProgress1.Focus()
                Exit Sub
            End If

            If cmdCoreDiameter.SelectedValue = "-1" Then
                sValidaPr = False
                MsgBox("Error in Core Diameter", MsgBoxStyle.Critical)
                cmdCoreDiameter.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub ValidarTopo()
        Try
            sValidarTopo = False

            ValidaTopo = False
            ValidaST = False
            ValidaCs = False

            X1 = 0
            X2 = 0
            X3 = 0

            Y1 = 0
            Y2 = 0
            Y3 = 0

            Z1 = 0
            Z2 = 0
            Z3 = 0

            If txtEastPlanned.Text <> "" Then


                If (txtEastGpsTopo.Text <> "" Or txtNorthGpsTopo.Text <> "" Or txtElevationGpsTopo.Text <> "") And dtpDateGpsTopo.Value <> "01/01/1900" Then
                    If txtEastGpsTopo.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtEastGpsTopo.Focus()
                        Exit Sub
                    End If
                    If txtNorthGpsTopo.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtNorthGpsTopo.Focus()
                    End If
                    If txtElevationGpsTopo.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtElevationGpsTopo.Focus()
                    End If

                    X1 = (txtEastGpsTopo.Text - txtEastPlanned.Text) ^ 2
                    Y1 = (txtNorthGpsTopo.Text - txtNorthPlanned.Text) ^ 2
                    Z1 = (txtElevationGpsTopo.Text - txtElevationPlanned.Text) ^ 2
                    Raiz1 = Sqrt(X1 + Y1 + Z2)
                    If Raiz1 < 20 Then
                        sValidarTopo = True
                        If MsgBox("Existe una diferencia entre Planned y GPS = : " & Raiz1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            sValidarTopo = True
                            ValorCorreo = Raiz1
                            Coordenada = "Planned - GPS"
                            EnviarCorreo()
                            'Enviar Correo
                        Else
                            sValidarTopo = False
                            Exit Sub
                        End If
                    End If
                    If Raiz1 > 20 Then
                        sValidarTopo = False
                        MsgBox("Existe una diferencia entre Planned - GPS, >= : " & Raiz1 & " El registro no se puede guardar, Por favor Cominicarse con DB", MsgBoxStyle.Critical, "ERROR")
                        Exit Sub
                    End If
                Else
                    sValidarTopo = False
                    MsgBox("Verifique los datos GPS", MsgBoxStyle.Critical, "Error")
                    Exit Sub
                End If
                ValidaTopo = True
            End If
            'Poner misma validacion de planned
            'Si los campos estan vacios no permite guardar
            If txtEastGpsTopo.Text <> "" Then

                If (txtEastStTopo.Text <> "" Or txtNorthStTopo.Text <> "" Or txtElevationStTopo.Text <> "") And dtpDateStTopo.Value <> "01/01/1900" Then

                    If txtEastStTopo.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtEastStTopo.Focus()
                        Exit Sub
                    End If
                    If txtNorthStTopo.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtNorthStTopo.Focus()
                        Exit Sub
                    End If
                    If txtElevationStTopo.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtElevationStTopo.Focus()
                        Exit Sub
                    End If


                    X2 = (txtEastStTopo.Text - txtEastGpsTopo.Text) ^ 2
                    Y2 = (txtNorthStTopo.Text - txtNorthGpsTopo.Text) ^ 2
                    Z2 = (txtElevationStTopo.Text - txtElevationGpsTopo.Text) ^ 2
                    Raiz2 = Sqrt(X2 + Y2 + Z2)

                    If Raiz2 > 2 And Raiz2 < 10 Then
                        sValidarTopo = True
                        If MsgBox("Existe una diferencia entre GPS y ST = : " & Raiz2, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            ValorCorreo = Raiz2
                            Coordenada = "GPS y ST"
                            'Enviar Correo
                        Else
                            sValidarTopo = False
                            Exit Sub
                        End If
                    End If
                    If Raiz2 > 10 Then
                        sValidarTopo = False
                        MsgBox("Existe una diferencia entre GPS y ST, >= : " & Raiz2 & " El registro no se puede guardar, Por favor Cominicarse con DB", MsgBoxStyle.Critical, "ERROR")

                        Exit Sub
                    End If
                    ValidaST = True
                End If

                If ValidaST = False Then
                    If ValidaTopo = True And (txtEastStTopo.Text <> "" Or txtNorthStTopo.Text <> "" Or txtElevationStTopo.Text <> "" And dtpDateStTopo.Value <> "01/01/1900") Then
                        MsgBox("Verifique los datos ST", MsgBoxStyle.Critical, "Error")
                        ValidaST = False
                        sValidarTopo = False
                        Exit Sub
                    End If
                End If
            Else
                If ValidaTopo = True And (txtEastStTopo.Text <> "" Or txtNorthStTopo.Text <> "" Or txtElevationStTopo.Text <> "" And dtpDateStTopo.Value <> "01/01/1900") Then
                    sValidarTopo = False
                    MsgBox("Verifique los datos ST", MsgBoxStyle.Critical, "Error")
                    Exit Sub
                End If
            End If

            ' Si los campos estan vacios no validar
            If txtEastStTopo.Text <> "" Then
                If (txtEastCS.Text <> "" And txtNorthCS.Text <> "" And txtElevationCS.Text <> "") And dtpDateCS.Value <> "01/01/1900" Then

                    If txtEastCS.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtEastCS.Focus()
                        Exit Sub
                    End If
                    If txtNorthCS.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtNorthCS.Focus()
                    End If
                    If txtEastStTopo.Text = "" Then
                        sValidarTopo = False
                        MsgBox("Debe tener un valor real", MsgBoxStyle.Critical)
                        txtEastStTopo.Focus()
                    End If

                    X3 = (txtEastCS.Text - txtEastStTopo.Text) ^ 2
                    Y3 = (txtNorthCS.Text - txtNorthStTopo.Text) ^ 2
                    Z3 = (txtElevationCS.Text - txtElevationStTopo.Text) ^ 2
                    Raiz3 = Sqrt(X3 + Y3 + Z3)

                    If Raiz3 > 2 And Raiz3 < 10 Then
                        sValidarTopo = True
                        If MsgBox("Existe una diferencia entre ST y CS = : " & Raiz3, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            'Enviar Correo
                            ValorCorreo = Raiz3
                            Coordenada = "ST y CS"
                        Else
                            sValidarTopo = False
                            Exit Sub
                        End If
                    End If
                    If Raiz3 > 10 Then
                        sValidarTopo = False
                        MsgBox("Existe una diferencia entre ST y CS, >= : " & Raiz3 & " El registro no se puede guardar, Por favor Cominicarse con DB", MsgBoxStyle.Critical, "ERROR")
                        'Enviar Correo
                    End If
                    ValidaCs = True
                Else
                    If ValidaCs = False Then
                        If ValidaCs = True Or (txtEastCS.Text <> "" Or txtNorthCS.Text <> "" Or txtElevationCS.Text <> "" And dtpDateCS.Value <> "01/01/1900") Then
                            MsgBox("Verifique los datos CS", MsgBoxStyle.Critical, "Error")
                            ValidaST = False
                            sValidarTopo = False
                            Exit Sub
                        End If
                    End If
                End If
                If ValidaST = True And (txtEastCS.Text = "" Or txtNorthCS.Text = "" Or txtElevationCS.Text = "") Then
                    oPlatform.sDateSt = "01/01/1900"
                    ValidaST = False
                    'MsgBox("Verifique los datos CS", MsgBoxStyle.Critical, "Error")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub btnAddDrill1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub LimpiarDrillProgress()
        Try

            txtFrom1.Text = ""
            txtTo1.Text = ""
            txtCommentsDrill1.Text = ""
            txtFrom1.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub
    Private Sub btnAddProgress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtRodLostDrill1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try

            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub



    Private Sub txtCasingDrill1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnGExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGExcel.Click
        Try
            Dim oXL As Excel.Application
            Dim oWB As Excel._Workbook
            Dim oSheet As Excel._Worksheet
            Dim oRng As Excel.Range

            oXL = New Excel.Application()
            oXL.Visible = True

            oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings("Ruta_DrillProgram").ToString(), 0, False, 5, Type.Missing, Type.Missing, _
             False, Type.Missing, Type.Missing, True, False, Type.Missing, _
             False, False, False)

            oSheet = DirectCast(oWB.ActiveSheet, Excel._Worksheet)

            oSheet.Cells(2, 7) = clsRf.sUser
            oSheet.Cells(3, 7) = Date.Now


            oPlatform.sPlatformID = txtPlatform.Text
            Dim dtPlatformPlanned As DataTable = oPlatform.getDHPlatformPlanned
            'MsgBox(dtPlatformPlanned.Rows.Count)
            Dim iInicial As Integer = 6
            For i As Integer = 0 To dtPlatformPlanned.Rows.Count - 1
                'dtCollars.Rows(0)("CatastralFolioID").ToString()
                'MsgBox(dtPlatformPlanned.Rows(i)("PLATFORM").ToString())
                oSheet.Cells(iInicial, 1) = dtPlatformPlanned.Rows(i)("HoleID").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformPlanned.Rows(i)("Platform").ToString()
                oSheet.Cells(iInicial, 3) = dtPlatformPlanned.Rows(i)("Section").ToString()
                oSheet.Cells(iInicial, 4) = dtPlatformPlanned.Rows(i)("East").ToString()
                oSheet.Cells(iInicial, 5) = dtPlatformPlanned.Rows(i)("North").ToString()
                oSheet.Cells(iInicial, 6) = dtPlatformPlanned.Rows(i)("Elevation").ToString()
                oSheet.Cells(iInicial, 7) = dtPlatformPlanned.Rows(i)("Instrument").ToString()
                oSheet.Cells(iInicial, 8) = dtPlatformPlanned.Rows(i)("AzimuthPlanned").ToString()
                oSheet.Cells(iInicial, 9) = dtPlatformPlanned.Rows(i)("InclinationPlanned").ToString()
                oSheet.Cells(iInicial, 10) = dtPlatformPlanned.Rows(i)("LengthPlanned").ToString()
                oSheet.Cells(iInicial, 11) = dtPlatformPlanned.Rows(i)("EOH").ToString()
                oSheet.Cells(iInicial, 12) = dtPlatformPlanned.Rows(i)("Surf").ToString()
                oSheet.Cells(iInicial, 13) = dtPlatformPlanned.Rows(i)("Zon").ToString()
                oSheet.Cells(iInicial, 14) = dtPlatformPlanned.Rows(i)("Location").ToString()
                oSheet.Cells(iInicial, 15) = dtPlatformPlanned.Rows(i)("Start_Drilling").ToString()
                oSheet.Cells(iInicial, 16) = dtPlatformPlanned.Rows(i)("Final_Drilling").ToString()
                oSheet.Cells(iInicial, 17) = dtPlatformPlanned.Rows(i)("Status").ToString()
                oSheet.Cells(iInicial, 18) = dtPlatformPlanned.Rows(i)("Contractor").ToString()
                oSheet.Cells(iInicial, 19) = dtPlatformPlanned.Rows(i)("torelo").ToString()
                oSheet.Cells(iInicial, 20) = dtPlatformPlanned.Rows(i)("landper").ToString()
                oSheet.Cells(iInicial, 21) = dtPlatformPlanned.Rows(i)("landpermitsta").ToString()
                oSheet.Cells(iInicial, 22) = dtPlatformPlanned.Rows(i)("CatastralFolioID").ToString()
                oSheet.Cells(iInicial, 23) = dtPlatformPlanned.Rows(i)("CommentsPlanned").ToString()
                oSheet.Cells(iInicial, 24) = dtPlatformPlanned.Rows(i)("CommentsLand").ToString()
                oSheet.Cells(iInicial, 25) = dtPlatformPlanned.Rows(i)("LandOwner").ToString()
                iInicial += 1
            Next

            oXL.Visible = True

            oXL.UserControl = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtFrom1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtTo1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub dgDrillProgress1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try

            If MsgBox("Remove the item?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oPlatform.sID = dgDrillProgress1.CurrentRow.Cells.Item("id").Value
                Dim sResp As String = oPlatform.DelDHHoleInProgress()
                If sResp = "OK" Then
                    MsgBox("Drill Progress delete.", MsgBoxStyle.Information)
                    FillDrillProgress()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnAddTopo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTopo.Click
        Try
            ValidarTopo()
            If sValidarTopo = True Then
                If sEditTo = "0" Then
                    oPlatform.sPlatform = txtPlatform.Text.ToString

                    oPlatform.sEastGPS = txtEastGpsTopo.Text.ToString
                    oPlatform.sNothGPS = txtNorthGpsTopo.Text.ToString
                    oPlatform.sElevationGPS = txtElevationGpsTopo.Text.ToString
                    oPlatform.sSurveryor = cmbSurveryor.SelectedValue.ToString
                    oPlatform.sDateGps = dtpDateGpsTopo.Value.ToString

                    oPlatform.sEastST = txtEastStTopo.Text.ToString
                    oPlatform.sNothST = txtNorthStTopo.Text.ToString
                    oPlatform.sElevationST = txtElevationStTopo.Text.ToString
                    If cmbSurveryorST.SelectedText = "" Then
                        oPlatform.sSurveryorST = "-1"
                    Else
                        oPlatform.sSurveryorST = cmbSurveryorST.SelectedValue.ToString
                    End If

                    oPlatform.sDateSt = dtpDateStTopo.Value.ToString

                    If cmdCompanyService.SelectedText = "" Then
                        oPlatform.sCompanyService = "-1"
                    Else
                        oPlatform.sCompanyService = cmdCompanyService.SelectedValue.ToString
                    End If
                    'oPlatform.sLocation = txtLocationPlatformCompanyService.Text.ToString
                    oPlatform.sEastCS = txtEastCS.Text.ToString
                    oPlatform.sNorthCS = txtNorthCS.Text.ToString
                    oPlatform.sElevationCS = txtElevationCS.Text.ToString
                    oPlatform.sDateCs = dtpDateCS.Value.ToString
                    oPlatform.sCommentsTopo = txtCommentsTopo.Text.ToString

                    Dim sResp As String = oPlatform.DH_Platform_Topo_Upd()
                    If sResp = "OK" Then
                        MsgBox("Topography Save.", MsgBoxStyle.Information)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        
    End Sub

    Private Sub txtEastGpsTopo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEastGpsTopo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtNorthGpsTopo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNorthGpsTopo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtElevationGpsTopo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtElevationGpsTopo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtEastStTopo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEastStTopo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtNorthStTopo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNorthStTopo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtElevationStTopo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtElevationStTopo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtEastCS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEastCS.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtNorthCS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNorthCS.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtElevationCS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtElevationCS.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub FillMailSend()
        Try
            oPlatform.sIDProject = (ConfigurationSettings.AppSettings("IDProject").ToString)
            oPlatform.sModule = "Drilling"

            Dim dtMail As DataTable = oPlatform.getDHMailSend()
            sFrom = dtMail.Rows(0)("MailFrom").ToString
            sTo = dtMail.Rows(0)("MailTo").ToString
            sSubject = dtMail.Rows(0)("Subject").ToString
            sUserSend = dtMail.Rows(0)("UserSend").ToString
            sPassSend = dtMail.Rows(0)("PassSend").ToString
            sServer = dtMail.Rows(0)("Server").ToString
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    '    Public Sub EnviarCorreo(ByVal radicador, ByVal entregado, ByVal radicado, ByVal notas)
    Public Sub EnviarCorreo()
        Try
            FillMailSend()
            Dim UsuariosEnvio As String
            UsuariosEnvio = ConfigurationSettings.AppSettings("DrillingGroup").ToString()


            correo.From = New System.Net.Mail.MailAddress(sFrom)
            correo.To.Clear()

            correo.To.Add(UsuariosEnvio)

            correo.Subject = sSubject & " - Plataforma : " & txtPlatform.Text.ToString & ", Coordenada: " & Coordenada
            'correo.Subject = sSubject & " - Plataforma : " & txtPlatform.Text.ToString
            correo.Body = "Notas :" & vbNewLine
            correo.Body = correo.Body & "--------" & vbNewLine
            correo.Body = correo.Body & "Coordenada :" & Coordenada & " tiene una diferencia de : " & ValorCorreo
            correo.Body = correo.Body & vbNewLine
            correo.Body = correo.Body & "----------------------------------------------------------------------------------------------------" & vbNewLine
            correo.Body = correo.Body & "Correo Enviado Automaticamente por el Sistema de Perforaciones (Drilling)" & vbNewLine
            correo.Body = correo.Body & "SOPORTE TECNICO: Edwin O. Londoño G. - edwin.londono@grancolombiagold.com.co" & vbNewLine
            correo.IsBodyHtml = False
            correo.Priority = System.Net.Mail.MailPriority.High

            Dim smtp As New System.Net.Mail.SmtpClient
            smtp.Host = sServer
            smtp.Credentials = New System.Net.NetworkCredential(sUserSend, sPassSend)


            Try
                smtp.Send(correo)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim oXL As Excel.Application
            Dim oWB As Excel._Workbook
            Dim oSheet As Excel._Worksheet
            Dim oRng As Excel.Range

            oXL = New Excel.Application()
            oXL.Visible = True
            'Get a new workbook.
            'oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            'oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            'oWB = oXL.Workbooks.Open(@"D:/Template_Shipment_Sgs.xls", 0, true, 5,


            oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings("Ruta_DrillProgress").ToString(), 0, False, 5, Type.Missing, Type.Missing, _
             False, Type.Missing, Type.Missing, True, False, Type.Missing, _
             False, False, False)

            oSheet = DirectCast(oWB.ActiveSheet, Excel._Worksheet)
            Dim hoy As Date
            Dim manana As Date

            hoy = Date.Today
            manana = hoy.AddDays(-1)

            oSheet.Cells(4, 2) = hoy
            oSheet.Cells(5, 2) = manana


            oPlatform.sHoleID = ""
            Dim dtDrillProgress As DataTable = oPlatform.getDHDrillProgressReport
            'MsgBox(dtPlatformPlanned.Rows.Count)
            Dim iInicial As Integer = 11
            For i As Integer = 0 To dtDrillProgress.Rows.Count - 1
                oSheet.Cells(iInicial, 1) = dtDrillProgress.Rows(i)("Platform").ToString()
                oSheet.Cells(iInicial, 2) = dtDrillProgress.Rows(i)("HoleID").ToString()
                oSheet.Cells(iInicial, 3) = dtDrillProgress.Rows(i)("From").ToString()
                oSheet.Cells(iInicial, 4) = dtDrillProgress.Rows(i)("To").ToString()
                oSheet.Cells(iInicial, 5) = dtDrillProgress.Rows(i)("Meters").ToString()
                oSheet.Cells(iInicial, 6) = dtDrillProgress.Rows(i)("EOH").ToString()
                oSheet.Cells(iInicial, 7) = dtDrillProgress.Rows(i)("AzimuthPlanned").ToString()
                oSheet.Cells(iInicial, 8) = dtDrillProgress.Rows(i)("InclinationPlanned").ToString()
                oSheet.Cells(iInicial, 9) = dtDrillProgress.Rows(i)("StartDate").ToString()
                oSheet.Cells(iInicial, 10) = dtDrillProgress.Rows(i)("FinalDate").ToString()
                oSheet.Cells(iInicial, 11) = dtDrillProgress.Rows(i)("Contractor").ToString()
                oSheet.Cells(iInicial, 12) = dtDrillProgress.Rows(i)("Comments").ToString()
                oSheet.Cells(iInicial, 13) = dtDrillProgress.Rows(i)("RigUsed").ToString()
                iInicial += 1
            Next

            oXL.Visible = True

            oXL.UserControl = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnSelectPTopo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectPTopo.Click
        'Dim sVfDialogTopo As New SaveFileDialog
        'sVfDialogTopo.Filter = "Imagen JPeg|*.jpg"
        'sVfDialogTopo.Title = "Guardar Imágenes Topodráficas"
        'sVfDialogTopo.ShowDialog()
        Try
            fDialog.Filter = "Imagen JPeg|*.jpg"
            fDialog.Title = "Seleccione Imágenes Topodráficas"
            fDialog.ShowDialog()
            txtPicRTopo.Text = fDialog.FileName


            Dim oFi As New FileInfo(fDialog.FileName)
            Dim sExt As String = oFi.Extension.ToString()
            sFotoTopo = oFi.Name.Substring(0, oFi.Name.ToString().Length)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try



    End Sub

    Private Sub ListarImgTopo()
        Try
            lst_ImagenesTopo.Items.Clear()
            Dim Source As String
            Dim Destino As String
            Dim ArchivosT As New DirectoryInfo(ConfigurationSettings.AppSettings("Ruta_ImgTopo").ToString & txtHoleIDDrill1.Text.ToString)



            For Each file As FileInfo In ArchivosT.GetFiles()
                lst_ImagenesTopo.Items.Add(file.Name)
            Next
        Catch ex As Exception
            lst_ImagenesTopo.Items.Clear()
        End Try

    End Sub

    Private Sub ListarPDF()
        Try
            lstPdfFile.Items.Clear()
            Dim Source As String
            Dim Destino As String
            Dim ArchivosT As New DirectoryInfo(ConfigurationSettings.AppSettings("Ruta_PDF").ToString & cmbContractorDrill1.Text & "\Invoices\Scan\" & txtHoleIDDrill1.Text.ToString & "\InterventoryApproved\" & sFile)

            For Each file As FileInfo In ArchivosT.GetFiles()
                lstPdfFile.Items.Add(file.Name)
            Next
        Catch ex As Exception
            lstPdfFile.Items.Clear()
        End Try

    End Sub

    Private Sub btnAddImg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddImg.Click
        Try
            lst_ImagenesTopo.Items.Clear()
            Dim Source As String
            Dim Destino As String
            Dim ArchivosT As New DirectoryInfo(ConfigurationSettings.AppSettings("Ruta_ImgTopo").ToString & txtHoleIDDrill1.Text.ToString)

            Source = txtPicRTopo.Text
            Destino = ConfigurationSettings.AppSettings("Ruta_ImgTopo").ToString & txtHoleIDDrill1.Text.ToString & "\" & sFotoTopo
            'MsgBox(Destino)

            System.IO.File.Copy(Source, Destino, True)

            For Each file As FileInfo In ArchivosT.GetFiles()
                lst_ImagenesTopo.Items.Add(file.Name)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
            lst_ImagenesTopo.Items.Clear()
        End Try

        txtPicRTopo.Text = ""


    End Sub

    Private Sub lst_FilesEnvironment_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub lst_FilesEnvironment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub lst_ImagenesTopo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lst_ImagenesTopo.DoubleClick
        Try
            Dim ruta As String
            ruta = ConfigurationSettings.AppSettings("Ruta_ImgTopo").ToString & txtHoleIDDrill1.Text.ToString & "\" & lst_ImagenesTopo.SelectedItem
            System.Diagnostics.Process.Start(ruta)
        Catch ex As Exception


        End Try

    End Sub



    Private Sub txtDepth1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDepth1.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtDepth1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDepth1.TextChanged

    End Sub

    Private Sub txtDepth2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDepth2.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtDepth2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDepth2.TextChanged

    End Sub

    Private Sub txtDepth3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDepth3.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtHoleIDDrill1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHoleIDDrill1.TextChanged

    End Sub

    Private Sub txtHoleIDDrill1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHoleIDDrill1.KeyPress

    End Sub

    Private Sub btnAddDrill1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddDrill1.Click
        Try
            validardrilling()
            If sValidaDr = True Then
                If sEditDr = "0" Then
                    oPlatform.sPlatform = txtPlatform.Text.ToString
                    oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
                    oPlatform.sEOH = txtEOHDrill1.Text.ToString
                    oPlatform.sStartDate = dtpStartDateDrill1.Value.ToString
                    oPlatform.sFinalDate = dtpEndDateDrill1.Value.ToString
                    oPlatform.sRigUsed = cmbRigUsedDrill1.SelectedValue.ToString
                    oPlatform.sContractor = cmbContractorDrill1.SelectedValue.ToString
                    oPlatform.sRodLost = txtRodLostDrill1.Text.ToString
                    oPlatform.sCasing = txtCasingDrill1.Text.ToString
                    oPlatform.sEdit = 1

                    Dim sResp As String = oPlatform.DH_Platform_Drilling_Upd()
                    If sResp = "OK" Then

                        oRf.InsertTrans("Drilling", "Update", clsRf.sUser.ToString(), _
                        "Hole ID: " + txtHoleIDDrill1.Text.ToString + ". " + _
                        "Start Date: " + dtpStartDateDrill1.Value + ". " + _
                        "Contractor: " + cmbContractorDrill1.SelectedValue + ". " + _
                        "Rig: " + cmbRigUsedDrill1.SelectedValue + ". " + _
                        "Event Date " + Date.Now())

                        MsgBox("Drilling Updated.", MsgBoxStyle.Information)
                    End If
                    sEditDr = "0"
                End If
            End If

            FillHoleInProgress()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub btnAddProgress1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddProgress1.Click
        Try
            ValidarDrillProgress()

            If sValidaPr = True Then
                If txtHoleIDDrill1.Text.ToString <> "" And txtFrom1.Text.ToString <> "" And txtTo1.Text.ToString <> "" Then

                    oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
                    oPlatform.sFrom = txtFrom1.Text.ToString
                    oPlatform.sTo = txtTo1.Text.ToString
                    Dim dtValidacion As DataTable = oPlatform.getDHDrillProgressValidacionFT()
                    If sRowEditDrilling > 0 Then
                    Else
                        If dtValidacion.Rows.Count <> 0 Then
                            MsgBox("Range 'From To' Overlaps", MsgBoxStyle.Critical, "Error")
                            Exit Sub
                        End If
                    End If
                End If

                If sEditDP = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
                    oPlatform.sFrom = txtFrom1.Text.ToString
                    oPlatform.sTo = txtTo1.Text.ToString
                    oPlatform.sComments = txtCommentsDrill1.Text.ToString
                    oPlatform.sDate = dtpDateProgress1.Value.ToString
                    oPlatform.sRfCoreDiameter = cmdCoreDiameter.SelectedValue.ToString
                    oPlatform.sID = ""
                    Dim sResp As String = oPlatform.DH_Drill_Progress_Add()
                    If sResp = "OK" Then
                        oRf.InsertTrans("DH_DrillProgress", "Insert", clsRf.sUser.ToString(), _
                        "Hole ID: " + txtHoleIDDrill1.Text.ToString + ". " + _
                        "From: " + txtFrom1.Text.ToString + ". " + _
                        "To: " + txtTo1.Text.ToString + ". " + _
                        "Date: " + dtpDateProgress1.Value.ToString + ". " + _
                        "Event Date " + Date.Now())


                        MsgBox("Drill Progress Save.", MsgBoxStyle.Information)

                        ' validar si está llegando a los límites.
                        oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString

                        Dim dtDepth As DataTable = oPlatform.getDH_DrillingTime_Depth()
                        Dim Depth1, Depth2, Depth3, Mail1, Mail2, Mail3 As String
                        Dim VeinConst As Decimal
                        Dim ValorTo As String = txtTo1.Text.ToString
                        Dim ValorAproxMayor As Decimal
                        Dim ValorAproxMenor As Decimal

                        VeinConst = ConfigurationSettings.AppSettings("ValSampleMail").ToString()

                        Depth1 = dtDepth.Rows(0)("Depth1").ToString
                        Mail1 = dtDepth.Rows(0)("Mail1").ToString

                        ValorTo = txtTo1.Text.ToString

                        If Mail1 = "" Then
                            Mail1 = False
                        End If
                        If Depth1 <> "" Then
                            ValorAproxMayor = Depth1 + VeinConst
                            ValorAproxMenor = Depth1 - VeinConst
                            If (ValorTo >= ValorAproxMenor And ValorTo <= ValorAproxMayor) Then

                                sMetros = Depth1 - ValorTo
                                If sMetros < 0 Then
                                    sMensaje = "Superamos por "
                                Else
                                    sMensaje = "Estamos a "
                                End If
                                sDepth = Depth1
                                sVeta = dtDepth.Rows(0)("Vein1").ToString

                                EnviarCorreoAlertaVein()
                            End If
                        End If

                        Depth2 = dtDepth.Rows(0)("Depth2").ToString
                        Mail2 = dtDepth.Rows(0)("Mail2").ToString

                        If Depth2 <> "" Then
                            ValorAproxMayor = Depth2 + VeinConst
                            ValorAproxMenor = Depth2 - VeinConst
                            If (ValorTo >= ValorAproxMenor And ValorTo <= ValorAproxMayor) Then
                                sMetros = Depth2 - ValorTo
                                If sMetros < 0 Then
                                    sMensaje = "Superamos por "
                                Else
                                    sMensaje = "Estamos a "
                                End If
                                sDepth = Depth2
                                sVeta = dtDepth.Rows(0)("Vein2").ToString

                                EnviarCorreoAlertaVein()
                            End If
                        End If

                        Depth3 = dtDepth.Rows(0)("Depth3").ToString
                        Mail3 = dtDepth.Rows(0)("Mail3").ToString

                        If Depth3 <> "" Then
                            ValorAproxMayor = Depth3 + VeinConst
                            ValorAproxMenor = Depth3 - VeinConst
                            If (ValorTo >= ValorAproxMenor And ValorTo <= ValorAproxMayor) Then
                                sMetros = Depth3 - ValorTo
                                If sMetros < 0 Then
                                    sMensaje = "Superamos por "
                                Else
                                    sMensaje = "Estamos a "
                                End If
                                sDepth = Depth3
                                sVeta = dtDepth.Rows(0)("Vein3").ToString

                                EnviarCorreoAlertaVein()
                            End If
                        End If

                        'FillMaxTo()
                        FillDrillProgress()
                        sEditDP = "0"



                    End If

                Else
                    oPlatform.sOpcion = "2"
                    oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
                    oPlatform.sFrom = txtFrom1.Text.ToString
                    oPlatform.sTo = txtTo1.Text.ToString
                    oPlatform.sComments = txtCommentsDrill1.Text.ToString
                    oPlatform.sDate = dtpDateProgress1.Value.ToString
                    oPlatform.sRfCoreDiameter = cmdCoreDiameter.SelectedValue.ToString
                    oPlatform.sID = sRowEditDrilling
                    Dim sResp As String = oPlatform.DH_Drill_Progress_Add()
                    If sResp = "OK" Then

                        oRf.InsertTrans("DH_DrillProgress", "Update", clsRf.sUser.ToString(), _
                        "Hole ID: " + txtHoleIDDrill1.Text.ToString + ". " + _
                        "From: " + txtFrom1.Text.ToString + ". " + _
                        "To: " + txtTo1.Text.ToString + ". " + _
                        "Date: " + dtpDateProgress1.Value.ToString + ". " + _
                        "Event Date " + Date.Now())


                        MsgBox("Drill Progress Update.", MsgBoxStyle.Information)
                        FillDrillProgress()
                        sEditDP = "0"
                    End If
                End If
                FillMaxTo()
                'txtFrom1.Text = txtTo1.Text
                txtTo1.Text = ""
                txtCommentsDrill1.Text = ""
                txtTo1.Focus()

                sRowEditDrilling = 0
                'sEditDP = 0

            End If
            ValidaFtAvanceDiario()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub dgDrillProgress1_CellContentDoubleClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgDrillProgress1.CellContentDoubleClick
        Try
            If MsgBox("Remove the item?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oPlatform.sID = dgDrillProgress1.CurrentRow.Cells.Item("id").Value
                Dim sResp As String = oPlatform.DelDHHoleInProgress()
                If sResp = "OK" Then

                    oRf.InsertTrans("DH_Platform", "Delete", clsRf.sUser.ToString(), _
                    "Hole ID: " + txtHoleIDDrill1.Text.ToString + ". " + _
                    "From: " + txtFrom1.Text.ToString + ". " + _
                    "To: " + txtTo1.Text.ToString + ". " + _
                    "Date: " + dtpDateProgress1.Value.ToString + ". " + _
                    "Event Date " + Date.Now())

                    MsgBox("Drill Progress delete.", MsgBoxStyle.Information)
                    FillDrillProgress()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub cmbPorcentajeTool_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPorcentajeTool.SelectedIndexChanged
        ' MsgBox(cmbPorcentajeTool.SelectedItem)
    End Sub

    Private Sub btnExcelDrillProgress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcelDrillProgress.Click
        Try
            Dim oXL As Excel.Application
            Dim oWB As Excel._Workbook
            Dim oSheet As Excel._Worksheet
            Dim oRng As Excel.Range

            oXL = New Excel.Application()
            oXL.Visible = True


            oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings("Ruta_DrillProgress").ToString(), 0, False, 5, Type.Missing, Type.Missing, _
             False, Type.Missing, Type.Missing, True, False, Type.Missing, _
             False, False, False)

            oSheet = DirectCast(oWB.ActiveSheet, Excel._Worksheet)

            'MsgBox(DateAdd(DateInterval.Day, -1, Date.Now))
            oSheet.Cells(4, 2) = Date.Now
            oSheet.Cells(5, 2) = DateAdd(DateInterval.Day, -1, Date.Now)

            oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString
            Dim dtPlatformProgress As DataTable = oPlatform.getDHDrillProgressReport
            'MsgBox(dtPlatformPlanned.Rows.Count)
            Dim iInicial As Integer = 11
            For i As Integer = 0 To dtPlatformProgress.Rows.Count - 1
                'dtCollars.Rows(0)("CatastralFolioID").ToString()
                'MsgBox(dtPlatformPlanned.Rows(i)("PLATFORM").ToString())
                oSheet.Cells(iInicial, 1) = dtPlatformProgress.Rows(i)("Platform").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformProgress.Rows(i)("HoleID").ToString()
                oSheet.Cells(iInicial, 3) = dtPlatformProgress.Rows(i)("From").ToString()
                oSheet.Cells(iInicial, 4) = dtPlatformProgress.Rows(i)("To").ToString()
                oSheet.Cells(iInicial, 5) = dtPlatformProgress.Rows(i)("Meters").ToString()
                oSheet.Cells(iInicial, 6) = dtPlatformProgress.Rows(i)("EOH").ToString()
                oSheet.Cells(iInicial, 7) = dtPlatformProgress.Rows(i)("AzimuthPlanned").ToString()
                oSheet.Cells(iInicial, 8) = dtPlatformProgress.Rows(i)("InclinationPlanned").ToString()
                oSheet.Cells(iInicial, 9) = dtPlatformProgress.Rows(i)("LengthPlanned").ToString()
                oSheet.Cells(iInicial, 10) = dtPlatformProgress.Rows(i)("StartDate").ToString()
                oSheet.Cells(iInicial, 11) = dtPlatformProgress.Rows(i)("FinalDate").ToString()
                oSheet.Cells(iInicial, 12) = dtPlatformProgress.Rows(i)("Status").ToString()
                oSheet.Cells(iInicial, 13) = dtPlatformProgress.Rows(i)("Contractor").ToString()
                oSheet.Cells(iInicial, 14) = dtPlatformProgress.Rows(i)("RigUsed").ToString()
                oSheet.Cells(iInicial, 15) = dtPlatformProgress.Rows(i)("Comments").ToString()
                iInicial += 1
            Next

            oXL.Visible = True
            oXL.UserControl = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnNewDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewDC.Click
        Try
            nuevodrillcompany()

            btnAddDC.Enabled = True
            btnNewDC.Enabled = False
            dtDateCD.Enabled = True
            cmbTurnDC.Enabled = True
            dtRgNoDC.Enabled = True
            sEditCd = "0"

            tbcOpciones.Enabled = False
            dtDateCD.Focus()
            txtCommentsDC.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub nuevodrillcompany()
        Try
            dtDateCD.Value = "01/01/1900"
            cmbTurnDC.SelectedValue = "-1"
            dtRgNoDC.SelectedValue = "-1"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub ValidarDrillCompany()
        Try
            sValidarDC = True
            If dtDateCD.Value = "01/01/1900" Then
                sValidarDC = False
                MsgBox("Error in Date", MsgBoxStyle.Critical, "Error")
                dtDateCD.Focus()
                Exit Sub
            End If

            If cmbTurnDC.SelectedValue = "-1" Then
                sValidarDC = False
                MsgBox("Error in Turn", MsgBoxStyle.Critical, "Error")
                cmbTurnDC.Focus()
                Exit Sub
            End If

            If dtDateCD.Value > Date.Now() Then
                sValidarDC = False
                MsgBox("Error in Date selected > present date", MsgBoxStyle.Critical, "Error")
                dtDateCD.Focus()
                Exit Sub
            End If

            If dtRgNoDC.SelectedValue = "-1" Then
                sValidarDC = False
                MsgBox("Error in Rig", MsgBoxStyle.Critical, "Error")
                dtRgNoDC.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        


    End Sub

    Private Sub btnAddDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddDC.Click
        Try
            ValidarDrillCompany()
            If sValidarDC = True Then

                If sEditCd = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sIdDc = 0
                    oPlatform.sDate = dtDateCD.Value.ToString
                    oPlatform.sTurn = cmbTurnDC.SelectedValue.ToString
                    oPlatform.sRig = dtRgNoDC.SelectedValue.ToString
                    oPlatform.sComments = txtCommentsDC.Text.ToString
                    oPlatform.sProject = ConfigurationSettings.AppSettings("IDProject").ToString

                    oRf.InsertTrans("Drilling Time", "Insert", clsRf.sUser.ToString(), _
                    "Rig: " + dtRgNoDC.SelectedText + ". " + _
                    "Date: " + dtDateCD.Value + ". " + _
                    "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                    "Event Date " + Date.Now())


                    Dim sResp As String = oPlatform.DH_Company_Drill_Add()
                    If sResp <> "" Then
                        MsgBox("Company Drill Inserted.", MsgBoxStyle.Information)
                        sEditCd = "0"
                        'FillCompanyDrill()
                    End If
                Else

                    oPlatform.sOpcion = "2"
                    oPlatform.sID = sID
                    oPlatform.sDate = dtDateCD.Value.ToString
                    oPlatform.sTurn = cmbTurnDC.SelectedValue.ToString
                    oPlatform.sRig = dtRgNoDC.SelectedValue.ToString
                    oPlatform.sComments = txtCommentsDC.Text.ToString
                    oPlatform.sProject = ConfigurationSettings.AppSettings("IDProject").ToString

                    oRf.InsertTrans("Drilling Time", "Update", clsRf.sUser.ToString(), _
                    "Rig: " + dtRgNoDC.SelectedText + ". " + _
                    "Date: " + dtDateCD.Value + ". " + _
                    "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                    "Event Date " + Date.Now())

                    Dim sResp As String = oPlatform.DH_Company_Drill_Add()
                    If sResp = "OK" Then
                        MsgBox("Company Drill Update.", MsgBoxStyle.Information)
                        sEditCd = "1"
                    End If

                    'sEditCd = "1"
                End If
                FillCompanyDrill()

                btnAddDC.Enabled = False
                btnNewDC.Enabled = True
                dtDateCD.Enabled = False
                cmbTurnDC.Enabled = False
                dtRgNoDC.Enabled = False
                tbcOpciones.Enabled = False
                txtCommentsDC.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub FillCoreDiameter()
        Try
            oPlatform.sRfCoreDiameter = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtCore As DataTable = oPlatform.getRfCoreDiameter()
            Dim drC As DataRow = dtCore.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtCore.Rows.Add(drC)
            cmdCoreDiameter.DataSource = dtCore
            cmdCoreDiameter.DisplayMember = "Description"
            cmdCoreDiameter.ValueMember = "ID"
            cmdCoreDiameter.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    'Company Drill
    Private Sub FillRig()
        Try
            oPlatform.sRfRig = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtRig As DataTable = oPlatform.getRfRig()
            Dim drC As DataRow = dtRig.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtRig.Rows.Add(drC)
            dtRgNoDC.DataSource = dtRig
            dtRgNoDC.DisplayMember = "Description"
            dtRgNoDC.ValueMember = "RigID"
            dtRgNoDC.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillTur()
        Try
            oPlatform.sRfTurn = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtRig As DataTable = oPlatform.getRfTurn()
            Dim drC As DataRow = dtRig.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtRig.Rows.Add(drC)
            cmbTurnDC.DataSource = dtRig
            cmbTurnDC.DisplayMember = "Description"
            cmbTurnDC.ValueMember = "ID"
            cmbTurnDC.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub FillCompanyDrill()
        Try
            oPlatform.sRegistro = 1
            Dim dtCompanyDrill As DataTable = oPlatform.getRegistroListt()
            dgCompanyDrill.DataSource = dtCompanyDrill
            'dgCompanyDrill.Columns(5).Visible = False
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub dgCompanyDrill_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgCompanyDrill.CellClick
        Try
            oPlatform.sRegistro = dgCompanyDrill.CurrentRow.Cells.Item("id").Value
            sID = dgCompanyDrill.CurrentRow.Cells.Item("id").Value
            Dim dtRegistro As DataTable = oPlatform.getRegistroListtID
            dtDateCD.Value = dtRegistro.Rows(0)("date").ToString
            cmbTurnDC.SelectedValue = dtRegistro.Rows(0)("Turn").ToString
            dtRgNoDC.SelectedValue = dtRegistro.Rows(0)("Rig").ToString
            txtCommentsDC.Text = dtRegistro.Rows(0)("Comments").ToString
            sEditCd = "1"
            btnAddDC.Enabled = True
            tbcOpciones.Enabled = True

            'FillMeterTurn()
            'FillDownTime()
            'FillChangeCrown()
            'FillTurnSuppplies()
            'FillBiabilityCon()
            'FillBiabilityCom()
            FillDrillingTime()
            FillLostTools()
            FillBillableAdditives()



            dtDateCD.Enabled = True
            cmbTurnDC.Enabled = True
            dtRgNoDC.Enabled = True

            cmbDownTimeCD.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub



    'Private Sub dgCompanyDrill_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgCompanyDrill.CellContentDoubleClick
    '    ''
    '    oPlatform.sRegistro = dgCompanyDrill.CurrentRow.Cells.Item("id").Value
    '    sID = dgCompanyDrill.CurrentRow.Cells.Item("id").Value
    '    Dim dtRegistro As DataTable = oPlatform.getRegistroListtID
    '    dtDateCD.Value = dtRegistro.Rows(0)("date").ToString
    '    cmbTurnDC.SelectedValue = dtRegistro.Rows(0)("Turn").ToString
    '    dtRgNoDC.SelectedValue = dtRegistro.Rows(0)("Rig").ToString
    '    txtCommentsDC.Text = dtRegistro.Rows(0)("Comments").ToString
    '    sEditCd = "1"
    '    btnAddDC.Enabled = True
    '    tbcOpciones.Enabled = True

    '    FillMeterTurn()
    '    FillDownTime()
    '    FillChangeCrown()
    '    FillTurnSuppplies()
    '    FillBiabilityCon()
    '    FillBiabilityCom()
    '    FillLostTools()
    '    FillBillableAdditives()
    '    FillDrillingTime()

    '    dtDateCD.Enabled = True
    '    cmbTurnDC.Enabled = True
    '    dtRgNoDC.Enabled = True

    '    cmbDownTimeCD.Focus()
    'End Sub

    Private Sub btnAdd_MetersTurn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd_MetersTurn.Click
        Try
            Call ValidarMeterTurn()
            If sValidarMD = True Then

                If sEditMeter = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sHoleID = cmbHoleIdCD.SelectedValue.ToString
                    oPlatform.sIdDc = sID
                    oPlatform.sAzimuth = txtMeterTurnAz.Text.ToString
                    oPlatform.sSize = txtMeterTurnSize.Text.ToString
                    oPlatform.sFrom = txtMeterTurnFrom.Text.ToString
                    oPlatform.sTo = txtMeterTurnTo.Text.ToString
                    oPlatform.sComments = txtMeterTurnComments.Text.ToString
                    oPlatform.sID = ""

                    Dim sResp As String = oPlatform.DH_DrillMeterTurn_Add()
                    If sResp <> "" Then
                        'MsgBox("Company Drill Inserted.", MsgBoxStyle.Information)

                        sEditCd = "0"
                        FillCompanyDrill()
                    End If
                Else

                    oPlatform.sOpcion = "2"
                    oPlatform.sHoleID = cmbHoleIdCD.SelectedValue.ToString
                    oPlatform.sID = sID
                    oPlatform.sAzimuth = txtMeterTurnAz.Text.ToString
                    oPlatform.sSize = txtMeterTurnSize.Text.ToString
                    oPlatform.sFrom = txtMeterTurnFrom.Text.ToString
                    oPlatform.sTo = txtMeterTurnTo.Text.ToString
                    oPlatform.sComments = txtMeterTurnComments.Text.ToString

                    Dim sResp As String = oPlatform.DH_DrillMeterTurn_Add()
                    If sResp = "OK" Then
                        'MsgBox("Company Drill Update.", MsgBoxStyle.Information)
                        sEditCd = "1"
                    End If
                    'sEditCd = "1"
                End If

            End If
            FillMeterTurn()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub FillMeterTurn()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getDrillMeterTurn_IdDc()
            dgMeterTurn.DataSource = dtMeterTurn
            'dgMeterTurn.Columns(0).Visible = False
            'dgMeterTurn.Columns(1).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillDownTime()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getDrillDownTime_IdDc()
            dgDownTime.DataSource = dtMeterTurn
            'dgDownTime.Columns(0).Visible = False
            'dgDownTime.Columns(1).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillChangeCrown()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getDrillChangeCrown_IdDc()
            dgChangeCrown.DataSource = dtMeterTurn
            'dgChangeCrown.Columns(0).Visible = False
            'dgChangeCrown.Columns(1).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillTurnSuppplies()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getTurnSupplies_IdDc()
            dgTurnSuppplies.DataSource = dtMeterTurn
            'dgTurnSuppplies.Columns(0).Visible = False
            'dgTurnSuppplies.Columns(1).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillBiabilityCom()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getDH_BiabilityOfTimeCom_IdDc()
            dgBiabilityCompany.DataSource = dtMeterTurn
            'dgBiabilityCompany.Columns(0).Visible = False
            'dgBiabilityCompany.Columns(1).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillBiabilityCon()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getDH_BiabilityOfTimeCon_IdDc()
            dgBiabilityContractor.DataSource = dtMeterTurn
            'dgBiabilityContractor.Columns(0).Visible = False
            'dgBiabilityContractor.Columns(1).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillLostTools()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getDH_LostTools_IdDc()
            dgLostTools.DataSource = dtMeterTurn
            dgLostTools.Columns(0).Visible = False
            dgLostTools.Columns(1).Visible = False
            dgLostTools.Columns(2).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillDrillingTime()
        Try
            oPlatform.sIdDc = sID
            Dim dtDrilling As DataTable = oPlatform.getDH_DrillingTime_IdDc()
            dgDrillingTime.DataSource = dtDrilling
            dgDrillingTime.Columns(0).Visible = False
            dgDrillingTime.Columns(1).Visible = False
            dgDrillingTime.Columns(2).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub FillBillableAdditives()
        Try
            oPlatform.sIdDc = sID
            Dim dtMeterTurn As DataTable = oPlatform.getDH_BillableAdditives_IdDc()
            dbBillableAdditives.DataSource = dtMeterTurn
            dbBillableAdditives.Columns(0).Visible = False
            dbBillableAdditives.Columns(1).Visible = False
            dbBillableAdditives.Columns(2).Visible = False

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub ValidarMeterTurn()
        Try
            sValidarMD = True
            If cmbHoleIdCD.SelectedValue.ToString = "-1" Then
                sValidarMD = False
                MsgBox("Error in Hole", MsgBoxStyle.Critical)
                cmbHoleIdCD.Focus()
                Exit Sub
            End If
            If txtMeterTurnAz.Text = "" Then
                sValidarMD = False
                MsgBox("Error in Azimuth", MsgBoxStyle.Critical)
                txtMeterTurnAz.Focus()
                Exit Sub
            End If
            If txtMeterTurnSize.Text = "" Then
                sValidarMD = False
                MsgBox("Error in Size", MsgBoxStyle.Critical)
                txtMeterTurnSize.Focus()
                Exit Sub
            End If
            If txtMeterTurnFrom.Text = "" Then
                sValidarMD = False
                MsgBox("Error in From", MsgBoxStyle.Critical)
                txtMeterTurnFrom.Focus()
                Exit Sub
            End If
            If txtMeterTurnTo.Text = "" Then
                sValidarMD = False
                MsgBox("Error in To", MsgBoxStyle.Critical)
                txtMeterTurnTo.Focus()
                Exit Sub
            End If
            If Val(txtMeterTurnFrom.Text.ToString) > Val(txtMeterTurnTo.Text.ToString) Then
                sValidarMD = False
                MsgBox("Error in From", MsgBoxStyle.Critical)
                txtMeterTurnFrom.Focus()
                Exit Sub
            End If
            If txtMeterTurnAz.Text < 0 Or txtMeterTurnAz.Text > 360 Then
                sValidarMD = False
                MsgBox("Error in Azimuth Planned, range (0 to 360", MsgBoxStyle.Critical)
                txtMeterTurnAz.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dgCompanyDrill_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgCompanyDrill.CellContentClick

    End Sub

    Private Sub txtMeterTurnAz_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMeterTurnAz.KeyPress

    End Sub

    Private Sub txtMeterTurnAz_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMeterTurnAz.TextChanged

    End Sub

    Private Sub txtMeterTurnSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMeterTurnSize.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtMeterTurnSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMeterTurnSize.TextChanged

    End Sub

    Private Sub txtMeterTurnFrom_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMeterTurnFrom.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtMeterTurnTo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMeterTurnTo.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMeterTurnFrom_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMeterTurnFrom.TextChanged

    End Sub

    Private Sub txtDownTimeFrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDownTimeFrom.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub


    Private Sub txtDownTimeTo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDownTimeTo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub


    Private Sub txtChaCrownFrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtChaCrownFrom.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtChaCrownTo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtChaCrownTo.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub ValidarDt()
        Try
            sValidarDT = True
            If cmbDownTimeCD1.SelectedValue = "-1" Then
                sValidarDT = False
                MsgBox("Error in Description", MsgBoxStyle.Critical, "Error")
                cmbDownTimeCD1.Focus()
                Exit Sub
            End If

            If txtDownTimeFrom.Text = "" Then
                sValidarDT = False
                MsgBox("Error in From", MsgBoxStyle.Critical, "Error")
                txtDownTimeFrom.Focus()
                Exit Sub
            End If

            If txtDownTimeTo.Text = "" Then
                sValidarDT = False
                MsgBox("Error in To", MsgBoxStyle.Critical, "Error")
                txtDownTimeTo.Focus()
                Exit Sub
            End If


            If Val(txtDownTimeFrom.Text.ToString) > Val(txtDownTimeTo.Text.ToString) Then
                sValidarDT = False
                MsgBox("From > To", MsgBoxStyle.Critical, "Error")
                txtDownTimeFrom.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub ValidarCC()
        Try
            sValidarCC = True
            If cmbChaCrown.SelectedValue = "-1" Then
                sValidarCC = False
                MsgBox("Error in Description", MsgBoxStyle.Critical, "Error")
                cmbChaCrown.Focus()
                Exit Sub
            End If
            If txtChaCrownSerial.Text = "" Then
                sValidarCC = False
                MsgBox("Error in Serial", MsgBoxStyle.Critical, "Error")
                txtChaCrownSerial.Focus()
                Exit Sub
            End If

            If txtChaCrownFrom.Text = "" Then
                sValidarCC = False
                MsgBox("Error in From", MsgBoxStyle.Critical, "Error")
                txtChaCrownFrom.Focus()
                Exit Sub
            End If

            If txtChaCrownTo.Text = "" Then
                sValidarCC = False
                MsgBox("Error in To", MsgBoxStyle.Critical, "Error")
                txtChaCrownTo.Focus()
                Exit Sub
            End If

            If Val(txtChaCrownFrom.Text.ToString) > Val(txtChaCrownTo.Text.ToString) Then
                sValidarCC = False
                MsgBox("Error in To", MsgBoxStyle.Critical, "Error")
                txtChaCrownFrom.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub btnDownTimeAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownTimeAdd.Click
        Try
            ValidarDt()
            If sValidarDT = True Then
                If sEditDt = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sIdDc = sID
                    oPlatform.sIdDt = cmbDownTimeCD1.SelectedValue.ToString
                    oPlatform.sFrom = txtDownTimeFrom.Text.ToString
                    oPlatform.sTo = txtDownTimeTo.Text.ToString
                    oPlatform.sComments = txtMeterTurnComments.Text.ToString
                    oPlatform.sID = ""

                    Dim sResp As String = oPlatform.DH_DrillDownTime_Add()
                    If sResp <> "" Then
                        'MsgBox("Company Drill Inserted.", MsgBoxStyle.Information)

                        sEditDt = "0"
                        FillDownTime()
                    End If
                Else

                    oPlatform.sOpcion = "2"
                    oPlatform.sIdDc = sID
                    oPlatform.sIdDt = cmbDownTimeCD1.SelectedValue.ToString
                    oPlatform.sFrom = txtDownTimeFrom.Text.ToString
                    oPlatform.sTo = txtDownTimeTo.Text.ToString
                    oPlatform.sComments = txtMeterTurnComments.Text.ToString
                    oPlatform.sID = "" 'Cuando selecciones un item

                    Dim sResp As String = oPlatform.DH_DrillDownTime_Add()
                    If sResp = "OK" Then
                        'MsgBox("Company Drill Update.", MsgBoxStyle.Information)
                        sEditDt = "1"
                    End If
                End If

            End If
            FillDownTime()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub ValidarTS()
        Try
            sValidarTS = True
            If cmbTurnSupplies.SelectedValue = "-1" Then
                sValidarTS = False
                MsgBox("Error in Description", MsgBoxStyle.Critical, "Error")
                cmbTurnSupplies.Focus()
            End If
            If txtTurnSupliesAmount.Text = "" Then
                sValidarTS = False
                MsgBox("Error in Amount", MsgBoxStyle.Critical, "Error")
                txtTurnSupliesAmount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub btnChaCrownAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChaCrownAdd.Click
        Try
            ValidarCC()
            If sValidarCC = True Then
                If sEditCc = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sIdDc = sID
                    oPlatform.sIdCc = cmbChaCrown.SelectedValue.ToString
                    oPlatform.sSerial = txtChaCrownSerial.Text.ToString
                    oPlatform.sFrom = txtChaCrownFrom.Text.ToString
                    oPlatform.sTo = txtChaCrownTo.Text.ToString
                    oPlatform.sComments = txtChaCrownComments.Text.ToString
                    oPlatform.sID = ""

                    Dim sResp As String = oPlatform.DH_DrillChageCrown_Add()
                    If sResp <> "" Then
                        'MsgBox("Company Drill Inserted.", MsgBoxStyle.Information)

                        sEditCc = "0"
                        FillChangeCrown()
                    End If
                Else

                    oPlatform.sOpcion = "2"
                    oPlatform.sIdDc = sID
                    oPlatform.sIdCc = cmbChaCrown.SelectedValue.ToString
                    oPlatform.sSerial = txtChaCrownSerial.Text.ToString
                    oPlatform.sFrom = txtChaCrownFrom.Text.ToString
                    oPlatform.sTo = txtChaCrownTo.Text.ToString
                    oPlatform.sComments = txtChaCrownComments.Text.ToString
                    oPlatform.sID = "" 'Cuando selecciones un item

                    Dim sResp As String = oPlatform.DH_DrillChageCrown_Add()
                    If sResp = "OK" Then
                        'MsgBox("Company Drill Update.", MsgBoxStyle.Information)
                        sEditCc = "1"
                    End If
                End If

            End If
            FillChangeCrown()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub



    Private Sub dgDrillProgress1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgDrillProgress1.CellContentClick
        Try
            sEditDP = 1
            sRowEditDrilling = dgDrillProgress1.CurrentRow.Cells.Item("id").Value
            ' MsgBox(sRowEditDrilling)
            txtFrom1.Text = dgDrillProgress1.CurrentRow.Cells.Item("From").Value
            txtTo1.Text = dgDrillProgress1.CurrentRow.Cells.Item("To").Value
            cmdCoreDiameter.SelectedValue = dgDrillProgress1.CurrentRow.Cells.Item("CoreID").Value
            txtCommentsDrill1.Text = dgDrillProgress1.CurrentRow.Cells.Item("Comments").Value
            dtpDateProgress1.Value = dgDrillProgress1.CurrentRow.Cells.Item("Date").Value
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub txtTurnSupliesAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTurnSupliesAmount.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnTurnSupliesAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTurnSupliesAdd.Click
        Try
            ValidarTS()
            If sValidarTS = True Then
                If sEditTs = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sIdDc = sID
                    oPlatform.sIdTs = cmbTurnSupplies.SelectedValue.ToString
                    oPlatform.sAmount = txtTurnSupliesAmount.Text.ToString
                    oPlatform.sComments = txtTurnSupliesAmount.Text.ToString
                    oPlatform.sID = ""

                    Dim sResp As String = oPlatform.DH_DrillTurnSupplies_Add()
                    If sResp <> "" Then
                        sEditTs = "0"
                        FillTurnSuppplies()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub btnBiabilityCompanyAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBiabilityCompanyAdd.Click
        Try
            If cmbBuabilityCompany.SelectedValue = "-1" Then
                MsgBox("Error in Biability Company", MsgBoxStyle.Critical, "Error")
                Exit Sub
            Else
                If sEditBCom = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sIdDc = sID
                    oPlatform.sIdTs = cmbBuabilityCompany.SelectedValue.ToString
                    oPlatform.sID = ""

                    Dim sResp As String = oPlatform.DH_BiabilityOfTimeCom_Add()
                    If sResp <> "" Then
                        sEditBCom = "0"
                        FillBiabilityCom()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub btnBiabilityContractor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBiabilityContractor.Click
        Try
            If cmbBiabilityContractor.SelectedValue = "-1" Then
                MsgBox("Error in Biability Contractor", MsgBoxStyle.Critical, "Error")
                Exit Sub
            Else
                If sEditBCom = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sIdDc = sID
                    oPlatform.sIdTs = cmbBiabilityContractor.SelectedValue.ToString
                    oPlatform.sID = ""

                    Dim sResp As String = oPlatform.DH_BiabilityOfTimeCon_Add()
                    If sResp <> "" Then
                        sEditBCom = "0"
                        FillBiabilityCon()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
       
    End Sub

    Private Sub validarLostTools()
        Try
            sValidarLT = True
            If cmbLostTools.SelectedValue = "-1" Then
                sValidarLT = False
                MsgBox("Error in Description", MsgBoxStyle.Critical, "Error")
                cmbLostTools.Focus()
                Exit Sub
            End If

            If txtLostToolsAmount.Text = "" Then
                sValidarLT = False
                MsgBox("Error in Amount", MsgBoxStyle.Critical, "Error")
                txtLostToolsAmount.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub


    Private Sub btnLostToolsAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLostToolsAdd.Click
        Try
            validarLostTools()
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Drilling") = MsgBoxResult.Yes Then
                If sValidarLT = True Then
                    If sEditLt = "0" Then
                        oPlatform.sOpcion = "1"
                        oPlatform.sIdDc = sID
                        oPlatform.sIdTs = cmbLostTools.SelectedValue.ToString
                        oPlatform.sAmount = txtLostToolsAmount.Text.ToString
                        oPlatform.sPercentPay = cmbPorcentajeTool.SelectedItem.ToString
                        oPlatform.sPercentPayAdmon = cmbPorcentajeAdmon.SelectedItem.ToString
                        oPlatform.sComments = txtLostToolsComments.Text.ToString
                        oPlatform.sID = ""


                        'Registra auditoria'
                        oRf.InsertTrans("Lost Tools", "Insert", clsRf.sUser.ToString(), _
                        "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                        "Date: " + dtDateCD.Value + ". " + _
                        "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                        "Amount: " + txtLostToolsAmount.Text.ToString() + ". " + _
                        "%Pay: " + cmbPorcentajeTool.SelectedItem.ToString() + ". " + _
                        "%Pay Admon: " + cmbPorcentajeAdmon.SelectedItem.ToString() + ". " + _
                        "Event Date " + Date.Now())


                        Dim sResp As String = oPlatform.DH_LostTools_Add()
                        If sResp <> "" Then
                            'MsgBox("Inserted")
                            sEditLt = "0"
                            FillLostTools()
                        End If

                    Else

                        oPlatform.sOpcion = "2"
                        oPlatform.sIdDc = sID
                        oPlatform.sIdTs = cmbLostTools.SelectedValue.ToString
                        oPlatform.sAmount = txtLostToolsAmount.Text.ToString
                        oPlatform.sPercentPay = cmbPorcentajeTool.SelectedItem.ToString
                        oPlatform.sPercentPayAdmon = cmbPorcentajeAdmon.SelectedItem.ToString
                        oPlatform.sComments = txtLostToolsComments.Text.ToString
                        oPlatform.sID = sRowEdit

                        'Registra auditoria'
                        oRf.InsertTrans("Lost Tools", "Update", clsRf.sUser.ToString(), _
                        "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                        "Date: " + dtDateCD.Value + ". " + _
                        "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                        "Amount: " + txtLostToolsAmount.Text.ToString() + ". " + _
                        "%Pay: " + cmbPorcentajeTool.SelectedItem.ToString() + ". " + _
                        "%Pay Admon: " + cmbPorcentajeAdmon.SelectedItem.ToString() + ". " + _
                        "Event Date " + Date.Now())

                        Dim sResp As String = oPlatform.DH_LostTools_Add()
                        If sResp <> "" Then
                            'MsgBox("Updated")
                            sEditLt = "0"
                            FillLostTools()
                        End If
                    End If
                End If
            End If
            cmbLostTools.SelectedValue = "-1"
            txtLostToolsAmount.Text = ""
            cmbPorcentajeTool.SelectedItem = 0
            cmbPorcentajeAdmon.SelectedItem = 0
            txtLostToolsComments.Text = ""

            cmbLostTools.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub ValidarBillableAdditives()
        Try
            sValidarBA = True
            If cmbBillableAddit.SelectedValue = "-1" Then
                sValidarBA = False
                MsgBox("Error in Description", MsgBoxStyle.Critical, "Error")
                cmbBillableAddit.Focus()
                Exit Sub
            End If
            If txtBillableAdditivesAmount.Text = "" Then
                sValidarBA = False
                MsgBox("Error in Amount", MsgBoxStyle.Critical, "Error")
                txtBillableAdditivesAmount.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub


    Private Sub btnBillableAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBillableAdd.Click
        Try
            ValidarBillableAdditives()
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Drilling") = MsgBoxResult.Yes Then
                If sValidarBA = True Then
                    If sEditBA = "0" Then
                        oPlatform.sOpcion = "1"
                        oPlatform.sIdDc = sID
                        oPlatform.sIdBA = cmbBillableAddit.SelectedValue.ToString
                        oPlatform.sAmount = txtBillableAdditivesAmount.Text.ToString
                        oPlatform.sID = ""

                        'Registra auditoria'
                        oRf.InsertTrans("Billable Additives", "Inserted", clsRf.sUser.ToString(), _
                        "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                        "Date: " + dtDateCD.Value + ". " + _
                        "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                        "Additive: " + cmbBillableAddit.SelectedValue.ToString() + ". " + _
                        "Amount: " + txtBillableAdditivesAmount.Text.ToString() + ". " + _
                        "Event Date " + Date.Now())

                        Dim sResp As String = oPlatform.DH_BillableAdditives_Add()
                        If sResp <> "" Then

                            sEditBA = "0"
                            FillBillableAdditives()
                            'MsgBox("Inserted")
                        End If

                    Else
                        oPlatform.sOpcion = "2"
                        oPlatform.sIdDc = sID
                        oPlatform.sIdBA = cmbBillableAddit.SelectedValue.ToString
                        oPlatform.sAmount = txtBillableAdditivesAmount.Text.ToString
                        oPlatform.sID = sRowEdit

                        'Registra auditoria'
                        oRf.InsertTrans("Billable Additives", "Updated", clsRf.sUser.ToString(), _
                        "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                        "Date: " + dtDateCD.Value + ". " + _
                        "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                        "Additive: " + cmbBillableAddit.SelectedValue.ToString() + ". " + _
                        "Amount: " + txtBillableAdditivesAmount.Text.ToString() + ". " + _
                        "Event Date " + Date.Now())

                        Dim sResp As String = oPlatform.DH_BillableAdditives_Add()
                        If sResp <> "" Then
                            sEditBA = "0"
                            FillBillableAdditives()
                            'MsgBox("Updated")
                        End If
                    End If
                End If
            End If
            cmbBillableAddit.SelectedValue = "-1"
            txtBillableAdditivesAmount.Text = ""
            cmbBillableAddit.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub

    Private Sub btnAddDrillingTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddDrillingTime.Click
        Try
            validarDrillingTime()
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Drilling") = MsgBoxResult.Yes Then

                If sValidarDrTm = True Then
                    If sEditDrTm = "0" Then
                        oPlatform.sOpcion = "1"
                        oPlatform.sIdDc = sID
                        oPlatform.sIdDt = cmbDownTimeCD.SelectedValue.ToString
                        oPlatform.sResTimeCont = txtDrillingTimeResTimeCon1.Text.ToString
                        oPlatform.sResTimeComp = txtDrillingTimeResTimeCom1.Text.ToString
                        oPlatform.sTimeReportDrill = txtDrillingTimeTimeReportDrill1.Text.ToString
                        oPlatform.sTimeApprovedInter = txtDrillingTimeTimeApprovedInt1.Text.ToString

                        'Registra auditoria'
                        oRf.InsertTrans("Drilling Time", "Insert", clsRf.sUser.ToString(), _
                        "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                        "Date: " + dtDateCD.Value + ". " + _
                        "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                        "Resp. Time Cont: " + txtDrillingTimeResTimeCon1.Text.ToString() + ". " + _
                        "Resp. Time Comp: " + txtDrillingTimeResTimeCom1.Text.ToString() + ". " + _
                        "Time Report Cont: " + txtDrillingTimeTimeReportDrill1.Text.ToString() + ". " + _
                        "Time Approved Int: " + txtDrillingTimeTimeApprovedInt1.Text.ToString() + ". " + _
                        "Event Date " + Date.Now())

                        oPlatform.sID = ""

                        Dim sResp As String = oPlatform.DH_DrillingTime_Add()
                        If sResp <> "" Then
                            sEditBA = "0"
                            FillDrillingTime()
                            MsgBox("Inserted", MsgBoxStyle.Information)
                        End If
                    Else
                        oPlatform.sOpcion = "2"
                        oPlatform.sID = sRowEdit
                        oPlatform.sIdDc = sID
                        oPlatform.sIdDt = cmbDownTimeCD.SelectedValue.ToString
                        oPlatform.sResTimeCont = txtDrillingTimeResTimeCon1.Text.ToString
                        oPlatform.sResTimeComp = txtDrillingTimeResTimeCom1.Text.ToString
                        oPlatform.sTimeReportDrill = txtDrillingTimeTimeReportDrill1.Text.ToString
                        oPlatform.sTimeApprovedInter = txtDrillingTimeTimeApprovedInt1.Text.ToString

                        'Registra auditoria'
                        oRf.InsertTrans("Drilling Time", "Update", clsRf.sUser.ToString(), _
                        "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                        "Date: " + dtDateCD.Value + ". " + _
                        "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                        "Resp. Time Cont: " + txtDrillingTimeResTimeCon1.Text.ToString() + ". " + _
                        "Resp. Time Comp: " + txtDrillingTimeResTimeCom1.Text.ToString() + ". " + _
                        "Time Report Cont: " + txtDrillingTimeTimeReportDrill1.Text.ToString() + ". " + _
                        "Time Approved Int: " + txtDrillingTimeTimeApprovedInt1.Text.ToString() + ". " + _
                        "Event Date " + Date.Now())

                        Dim sResp As String = oPlatform.DH_DrillingTime_Add()
                        If sResp = "OK" Then
                            MsgBox("Update", MsgBoxStyle.Information)
                            sEditBA = "0"
                            sEditDrTm = "0"
                            FillDrillingTime()
                        End If
                    End If
                End If
            End If
            cmbDownTimeCD.Enabled = True
            cmbDownTimeCD.SelectedValue = "-1"
            txtDrillingTimeResTimeCon1.Text = ""
            txtDrillingTimeResTimeCom1.Text = ""
            txtDrillingTimeTimeReportDrill1.Text = ""
            txtDrillingTimeTimeApprovedInt1.Text = ""
            cmbDownTimeCD.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub validarDrillingTime()
        Try
            sValidarDrTm = True
            If cmbDownTimeCD.SelectedValue = "-1" Then
                sValidarDrTm = False
                MsgBox("Error in Description", MsgBoxStyle.Critical, "Error")
                cmbDownTimeCD.Focus()
                Exit Sub
            End If
            'MsgBox(txtDrillingTimeResTimeCon1.Text)
            If txtDrillingTimeResTimeCon1.Text = "  :" Then
                sValidarDrTm = False
                MsgBox("Error in time contractor min value = 00:00", MsgBoxStyle.Critical, "Error")
                txtDrillingTimeResTimeCon1.Focus()
                Exit Sub
            End If

            If txtDrillingTimeResTimeCom1.Text = "  :" Then
                sValidarDrTm = False
                MsgBox("Error in time contractor, min value = 00:00", MsgBoxStyle.Critical, "Error")
                txtDrillingTimeResTimeCon1.Focus()
                Exit Sub
            End If

            If txtDrillingTimeTimeReportDrill1.Text = "  :" Then
                sValidarDrTm = False
                MsgBox("Error in Time Reported By Driller, min value = 00:00", MsgBoxStyle.Critical, "Error")
                txtDrillingTimeTimeReportDrill1.Focus()
                Exit Sub
            End If
            If txtDrillingTimeTimeApprovedInt1.Text = "  :" Then
                sValidarDrTm = False
                MsgBox("Error in Time Approved By Interventory, min value = 00:00", MsgBoxStyle.Critical, "Error")
                txtDrillingTimeTimeApprovedInt1.Focus()
                Exit Sub
            End If

            'If txtDrillingTimeResTimeCon.Text = "" Or txtDrillingTimeResTimeCon.Text < "0" Then
            '    sValidarDrTm = False
            '    MsgBox("Error in time contractor (> = 0)", MsgBoxStyle.Critical, "Error")
            '    txtDrillingTimeResTimeCon.Focus()
            '    Exit Sub
            'End If
            'If txtDrillingTimeResTimeCom.Text = "" Or txtDrillingTimeResTimeCom.Text < "0" Then
            '    sValidarDrTm = False
            '    MsgBox("Error in time company (> = 0)", MsgBoxStyle.Critical, "Error")
            '    txtDrillingTimeResTimeCom.Focus()
            '    Exit Sub
            'End If
            'If txtDrillingTimeTimeReportDrill.Text = "" Or txtDrillingTimeTimeReportDrill.Text < "0" Then
            '    sValidarDrTm = False
            '    MsgBox("Error in Time Reported By Driller (> = 0)", MsgBoxStyle.Critical, "Error")
            '    txtDrillingTimeTimeReportDrill.Focus()
            '    Exit Sub
            'End If
            'If txtDrillingTimeTimeApprovedInt.Text = "" Or txtDrillingTimeTimeApprovedInt.Text < "0" Then
            '    sValidarDrTm = False
            '    MsgBox("Error in Time Approved By Interventory(> = 0)", MsgBoxStyle.Critical, "Error")
            '    txtDrillingTimeTimeApprovedInt.Focus()
            '    Exit Sub
            'End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub txtDrillingTimeResTimeCon_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDrillingTimeResTimeCon.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub


    Private Sub txtDrillingTimeResTimeCom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDrillingTimeResTimeCom.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtDrillingTimeTimeReportDrill_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDrillingTimeTimeReportDrill.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtDrillingTimeTimeApprovedInt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDrillingTimeTimeApprovedInt.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub tabOpciones_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tabOpciones.Click
        Try
            If Me.tabOpciones.SelectedIndex = 2 Then
                ValidaFtAvanceDiario()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnExcel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel2.Click
        Try
            Dim oXL As Excel.Application
            Dim oWB As Excel._Workbook
            Dim oSheet As Excel._Worksheet
            Dim oRng As Excel.Range

            oXL = New Excel.Application()
            oXL.Visible = True
            'Get a new workbook.
            'oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            'oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            'oWB = oXL.Workbooks.Open(@"D:/Template_Shipment_Sgs.xls", 0, true, 5,


            oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings("Ruta_DrillingPlannedHistory").ToString(), 0, False, 5, Type.Missing, Type.Missing, _
             False, Type.Missing, Type.Missing, True, False, Type.Missing, _
             False, False, False)

            oSheet = DirectCast(oWB.ActiveSheet, Excel._Worksheet)

            oSheet.Cells(2, 5) = clsRf.sUser
            oSheet.Cells(3, 5) = Date.Now


            oPlatform.sPlatformID = txtPlatform.Text
            Dim dtPlatformPlanned As DataTable = oPlatform.getDHPlatformPlannedHistory
            'MsgBox(dtPlatformPlanned.Rows.Count)
            Dim iInicial As Integer = 6
            For i As Integer = 0 To dtPlatformPlanned.Rows.Count - 1
                'dtCollars.Rows(0)("CatastralFolioID").ToString()
                'MsgBox(dtPlatformPlanned.Rows(i)("PLATFORM").ToString())
                oSheet.Cells(iInicial, 1) = dtPlatformPlanned.Rows(i)("Platform").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformPlanned.Rows(i)("EastPlanned").ToString()
                oSheet.Cells(iInicial, 3) = dtPlatformPlanned.Rows(i)("NorthPlanned").ToString()
                oSheet.Cells(iInicial, 4) = dtPlatformPlanned.Rows(i)("ElevationPlanned").ToString()
                oSheet.Cells(iInicial, 5) = dtPlatformPlanned.Rows(i)("InclinationPlanned").ToString()
                oSheet.Cells(iInicial, 6) = dtPlatformPlanned.Rows(i)("LengthPlanned").ToString()
                oSheet.Cells(iInicial, 7) = dtPlatformPlanned.Rows(i)("Comments").ToString()
                oSheet.Cells(iInicial, 8) = dtPlatformPlanned.Rows(i)("UpdateDate").ToString()
                oSheet.Cells(iInicial, 9) = dtPlatformPlanned.Rows(i)("User").ToString()
                iInicial += 1
            Next

            oXL.Visible = True

            oXL.UserControl = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub FillRigDill()
        Try
            oPlatform.sRfRig = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtRig As DataTable = oPlatform.getRfRig()
            Dim drC As DataRow = dtRig.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtRig.Rows.Add(drC)
            cmbRigUsedDrill1.DataSource = dtRig
            cmbRigUsedDrill1.DisplayMember = "Description"
            cmbRigUsedDrill1.ValueMember = "RigID"
            cmbRigUsedDrill1.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtSection_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSection.TextChanged

    End Sub

    Private Sub txtLostToolsAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLostToolsAmount.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub txtBillableAdditivesAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillableAdditivesAmount.KeyPress
        Try
            Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
            KeyAscii = CShort(SoloNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim val1 As Double
        'Dim val2 As Double = 50
        'Dim Dif As Double
        'val1 = Val(txtDepth1.Text)

        'dif = val1 - val2

        'MsgBox(Dif)
        'IntToStr(DaysInMonth(Now
        'MsgBox(DateTime.DaysInMonth(Now.Year, Now.Month))

    End Sub

    Private Sub dgDrillingTime_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgDrillingTime.CellClick
        Try
            cmbDownTimeCD.SelectedValue = dgDrillingTime.CurrentRow.Cells.Item("iddt").Value
            'cmbDownTimeCD.Enabled = False
            txtDrillingTimeTimeReportDrill1.Focus()
            sEditDrTm = "1"
            sRowEdit = dgDrillingTime.CurrentRow.Cells.Item("id").Value

            txtDrillingTimeTimeReportDrill1.Text = dgDrillingTime.CurrentRow.Cells.Item("TimeReportCont").Value
            txtDrillingTimeTimeApprovedInt1.Text = dgDrillingTime.CurrentRow.Cells.Item("TimeApprovedInter").Value
            txtDrillingTimeResTimeCon1.Text = dgDrillingTime.CurrentRow.Cells.Item("ResTimeCont").Value
            txtDrillingTimeResTimeCom1.Text = dgDrillingTime.CurrentRow.Cells.Item("ResTimeComp").Value
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dgDrillingTime_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgDrillingTime.CellContentClick
 
    End Sub


    Private Sub dgCompanyDrill_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgCompanyDrill.CellDoubleClick
        Try
            oPlatform.sRegistro = dgCompanyDrill.CurrentRow.Cells.Item("id").Value
            sID = dgCompanyDrill.CurrentRow.Cells.Item("id").Value
            Dim dtRegistro As DataTable = oPlatform.getRegistroListtID
            dtDateCD.Value = dtRegistro.Rows(0)("date").ToString
            cmbTurnDC.SelectedValue = dtRegistro.Rows(0)("Turn").ToString
            dtRgNoDC.SelectedValue = dtRegistro.Rows(0)("Rig").ToString
            txtCommentsDC.Text = dtRegistro.Rows(0)("Comments").ToString
            sEditCd = "1"
            btnAddDC.Enabled = True
            tbcOpciones.Enabled = True

            'FillMeterTurn()
            'FillDownTime()
            'FillChangeCrown()
            'FillTurnSuppplies()
            'FillBiabilityCon()
            'FillBiabilityCom()
            FillDrillingTime()
            FillLostTools()
            FillBillableAdditives()

            dtDateCD.Enabled = True
            cmbTurnDC.Enabled = True
            dtRgNoDC.Enabled = True

            cmbDownTimeCD.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub cmbSurface_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSurface.SelectedIndexChanged

    End Sub

    Private Sub dgDrillingTime_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgDrillingTime.CellContentDoubleClick
        Try
            If MsgBox("Remove the item?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oPlatform.sIdDel = dgDrillingTime.CurrentRow.Cells.Item("id").Value
                'MsgBox(dgDrillingTime.CurrentRow.Cells.Item("id").Value)
                Dim sResp As String = oPlatform.DelDHDrillingTime()
                If sResp = "OK" Then

                    oRf.InsertTrans("Drilling Time", "Delete", clsRf.sUser.ToString(), _
                    "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                    "Date: " + dtDateCD.Value + ". " + _
                    "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                    "Resp. Time Cont: " + txtDrillingTimeResTimeCon1.Text.ToString() + ". " + _
                    "Resp. Time Comp: " + txtDrillingTimeResTimeCom1.Text.ToString() + ". " + _
                    "Time Report Cont: " + txtDrillingTimeTimeReportDrill1.Text.ToString() + ". " + _
                    "Time Approved Int: " + txtDrillingTimeTimeApprovedInt1.Text.ToString() + ". " + _
                    "ID: " + oPlatform.sIdDel + ". " + _
                    "Event Date " + Date.Now())

                    MsgBox("Delete.", MsgBoxStyle.Information)
                    FillDrillingTime()
                End If
            End If
            cmbDownTimeCD.Enabled = True
            cmbDownTimeCD.SelectedValue = "-1"
            txtDrillingTimeResTimeCon1.Text = ""
            txtDrillingTimeResTimeCom1.Text = ""
            txtDrillingTimeTimeReportDrill1.Text = ""
            txtDrillingTimeTimeApprovedInt1.Text = ""
            cmbDownTimeCD.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dgDrillingTime_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgDrillingTime.CellDoubleClick

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnReportDrillingTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReportDrillingTime.Click
        'frmDrillingTimeReport.MdiParent = frmPpal
        frmDrillingTimeReport.ShowDialog()
    End Sub

    Private Sub txtCommentsPlanned_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCommentsPlanned.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'oPlatform.sHoleID = txtHoleIDDrill1.Text.ToString

        'Dim dtDepth As DataTable = oPlatform.getDH_DrillingTime_Depth()
        'Dim Depth1, Depth2, Depth3, Mail1, Mail2, Mail3 As String
        'Dim VeinConst As Decimal
        'Dim ValorTo As String = txtTo1.Text.ToString
        'Dim ValorAproxMayor As Decimal
        'Dim ValorAproxMenor As Decimal

        'VeinConst = ConfigurationSettings.AppSettings("MinValSampleMail").ToString()

        'Depth1 = dtDepth.Rows(0)("Depth1").ToString
        'Mail1 = dtDepth.Rows(0)("Mail1").ToString
        ''Se debe cambiar
        'ValorTo = "320.22"

        'If Mail1 = "" Then
        '    Mail1 = False
        'End If
        'If Depth1 <> "" Then
        '    ValorAproxMayor = Depth1 + VeinConst
        '    ValorAproxMenor = Depth1 - VeinConst
        '    If (ValorTo >= ValorAproxMenor And ValorTo <= ValorAproxMayor) Then

        '        sMetros = Depth1 - ValorTo
        '        If sMetros < 0 Then
        '            sMensaje = "Superamos por "
        '        Else
        '            sMensaje = "Estamos a "
        '        End If
        '        sDepth = Depth1
        '        sVeta = dtDepth.Rows(0)("Vein1").ToString

        '        EnviarCorreoAlertaVein()
        '    End If
        'End If

        'Depth2 = dtDepth.Rows(0)("Depth2").ToString
        'Mail2 = dtDepth.Rows(0)("Mail2").ToString



        'If Depth2 <> "" Then
        '    ValorAproxMayor = Depth2 + VeinConst
        '    ValorAproxMenor = Depth2 - VeinConst
        '    If (ValorTo >= ValorAproxMenor And ValorTo <= ValorAproxMayor) Then
        '        sMetros = Depth2 - ValorTo
        '        If sMetros < 0 Then
        '            sMensaje = "Superamos por "
        '        Else
        '            sMensaje = "Estamos a "
        '        End If
        '        sDepth = Depth2
        '        sVeta = dtDepth.Rows(0)("Vein2").ToString

        '        EnviarCorreoAlertaVein()
        '    End If
        'End If

        'Depth3 = dtDepth.Rows(0)("Depth3").ToString
        'Mail3 = dtDepth.Rows(0)("Mail3").ToString

        'If Depth3 <> "" Then
        '    ValorAproxMayor = Depth3 + VeinConst
        '    ValorAproxMenor = Depth3 - VeinConst
        '    If (ValorTo >= ValorAproxMenor And ValorTo <= ValorAproxMayor) Then
        '        sMetros = Depth3 - ValorTo
        '        If sMetros < 0 Then
        '            sMensaje = "Superamos por "
        '        Else
        '            sMensaje = "Estamos a "
        '        End If
        '        sDepth = Depth3
        '        sVeta = dtDepth.Rows(0)("Vein3").ToString

        '        EnviarCorreoAlertaVein()
        '    End If
        'End If

    End Sub

    Public Sub EnviarCorreoAlertaVein()
        Try
            FillMailSend()
            Dim UsuariosEnvio As String
            UsuariosEnvio = ConfigurationSettings.AppSettings("VeinAproximation").ToString()


            correo.From = New System.Net.Mail.MailAddress(sFrom)
            correo.To.Clear()

            correo.To.Add(UsuariosEnvio)

            correo.Subject = "Expected Cut Approximation: Platform : " & txtPlatform.Text.ToString & ", HoleID : " & txtHoleIDDrill1.Text.ToString
            correo.Body = "Notas :" & vbNewLine
            correo.Body = correo.Body & "--------" & vbNewLine
            correo.Body = correo.Body & sMensaje & sMetros & " metros de la veta: " & sVeta & " ,la cual esta planeada interceptar a una profundidad destinada de: " & sDepth & " metros"
            correo.Body = correo.Body & vbNewLine
            correo.Body = correo.Body & "Tomar medidas necesarias para asegurar recuperación."
            correo.Body = correo.Body & vbNewLine
            correo.Body = correo.Body & "----------------------------------------------------------------------------------------------------" & vbNewLine
            correo.Body = correo.Body & "Correo Enviado Automaticamente por el Sistema de Perforaciones (Drilling)" & vbNewLine
            correo.Body = correo.Body & "SOPORTE TECNICO: Edwin O. Londoño G. - edwin.londono@grancolombiagold.com.co" & vbNewLine
            correo.IsBodyHtml = False
            correo.Priority = System.Net.Mail.MailPriority.High

            Dim smtp As New System.Net.Mail.SmtpClient
            smtp.Host = sServer
            smtp.Credentials = New System.Net.NetworkCredential(sUserSend, sPassSend)


            Try
                smtp.Send(correo)
            Catch exI As Exception
                MsgBox(exI.Message)
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dgLostTools_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgLostTools.CellClick
        Try
            'cmbLostTools.SelectedValue = "-1"
            'txtLostToolsAmount.Text = ""
            'cmbPorcentajeTool.SelectedItem = 0
            'cmbPorcentajeAdmon.SelectedItem = 0
            txtLostToolsComments.Text = ""

            sEditLt = 1
            sRowEdit = dgLostTools.CurrentRow.Cells.Item("id").Value
            cmbLostTools.SelectedValue = dgLostTools.CurrentRow.Cells.Item("IDLT").Value
            txtLostToolsAmount.Text = dgLostTools.CurrentRow.Cells.Item("Amount").Value
            cmbPorcentajeTool.Text = dgLostTools.CurrentRow.Cells.Item("%Pay").Value
            cmbPorcentajeAdmon.Text = dgLostTools.CurrentRow.Cells.Item("%Pay_Admon").Value
            txtLostToolsComments.SelectedText = dgLostTools.CurrentRow.Cells.Item("Comments").Value
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        


    End Sub

    Private Sub dgLostTools_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgLostTools.CellContentDoubleClick
        Try
            If MsgBox("Remove the item?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oPlatform.sIdDel = dgLostTools.CurrentRow.Cells.Item("id").Value

                Dim sResp As String = oPlatform.DelDHlostTools()
                If sResp = "OK" Then

                    oRf.InsertTrans("Lost Tools", "Insert", clsRf.sUser.ToString(), _
                    "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                    "Date: " + dtDateCD.Value + ". " + _
                    "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                    "Amount: " + txtLostToolsAmount.Text.ToString() + ". " + _
                    "%Pay: " + cmbPorcentajeTool.SelectedItem.ToString() + ". " + _
                    "%Pay Admon: " + cmbPorcentajeAdmon.SelectedItem.ToString() + ". " + _
                    "Event Date " + Date.Now())

                    MsgBox("Delete.", MsgBoxStyle.Information)
                    FillLostTools()
                End If

            End If
            cmbLostTools.SelectedValue = "-1"
            txtLostToolsAmount.Text = ""
            cmbPorcentajeTool.SelectedItem = 0
            cmbPorcentajeAdmon.SelectedItem = 0
            txtLostToolsComments.Text = ""

            cmbLostTools.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dgLostTools_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgLostTools.CellContentClick

    End Sub

    Private Sub dbBillableAdditives_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dbBillableAdditives.CellClick
        Try
            txtBillableAdditivesAmount.Text = ""

            sEditBA = "1"
            sRowEdit = dbBillableAdditives.CurrentRow.Cells.Item("id").Value

            cmbBillableAddit.SelectedValue = dbBillableAdditives.CurrentRow.Cells.Item("IDBA").Value
            txtBillableAdditivesAmount.Text = dbBillableAdditives.CurrentRow.Cells.Item("Amount").Value
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        

    End Sub


    Private Sub dbBillableAdditives_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dbBillableAdditives.CellContentDoubleClick
        Try
            If MsgBox("Remove the item?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oPlatform.sIdDel = dbBillableAdditives.CurrentRow.Cells.Item("id").Value

                Dim sResp As String = oPlatform.DelDHBillAdditives()
                If sResp = "OK" Then

                    oRf.InsertTrans("Billable Additives", "Insert", clsRf.sUser.ToString(), _
                    "Rig: " + dtRgNoDC.SelectedValue.ToString + ". " + _
                    "Date: " + dtDateCD.Value + ". " + _
                    "Turn: " + cmbTurnDC.SelectedValue.ToString() + ". " + _
                    "Additive: " + cmbBillableAddit.SelectedValue.ToString() + ". " + _
                    "Amount: " + txtBillableAdditivesAmount.Text.ToString() + ". " + _
                    "Event Date " + Date.Now())

                    MsgBox("Delete.", MsgBoxStyle.Information)
                    FillBillableAdditives()
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dbBillableAdditives_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dbBillableAdditives.CellContentClick

    End Sub

    Private Sub btnPdf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPdf.Click
        Try
            fDialog.Filter = "Archivo PDF|*.pdf"
            fDialog.Title = "Seleccione Facturacion"
            fDialog.ShowDialog()
            txtPdf.Text = fDialog.FileName

            Dim oFi As New FileInfo(fDialog.FileName)
            Dim sExt As String = oFi.Extension.ToString()
            sFile = oFi.Name.Substring(0, oFi.Name.ToString().Length)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Private Sub btnAddPdf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPdf.Click
        Try
            If txtHoleIDDrill1.Text <> "" Then
                lstPdfFile.Items.Clear()
                Dim Source As String
                Dim Destino As String
                Dim ArchivosT As New DirectoryInfo(ConfigurationSettings.AppSettings("Ruta_PDF").ToString & cmbContractorDrill1.Text & "\Invoices\Scan\" & txtHoleIDDrill1.Text.ToString & "\InterventoryApproved\" & sFile)

                Source = txtPdf.Text
                Destino = ConfigurationSettings.AppSettings("Ruta_PDF").ToString & cmbContractorDrill1.Text & "\Invoices\Scan\" & txtHoleIDDrill1.Text.ToString & "\InterventoryApproved\" & sFile
                'MsgBox(Destino)

                System.IO.File.Copy(Source, Destino, True)

                'Registra auditoria'
                oRf.InsertTrans("Billing", "Insert", clsRf.sUser.ToString(), _
                "Hole ID: " + txtHoleIDDrill1.Text.ToString + ". " + _
                "File: " + sFile + ". " + _
                "Event Date " + Date.Now())

                For Each file As FileInfo In ArchivosT.GetFiles()
                    lstPdfFile.Items.Add(file.Name)
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            lstPdfFile.Items.Clear()
        End Try

        txtPdf.Text = ""
    End Sub

    Private Sub lst_ImagenesTopo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lst_ImagenesTopo.SelectedIndexChanged

    End Sub

    Private Sub lstPdfFile_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstPdfFile.DoubleClick
        Try
            Dim ruta As String
            ruta = ConfigurationSettings.AppSettings("Ruta_PDF").ToString & cmbContractorDrill1.Text & "\Invoices\Scan\" & txtHoleIDDrill1.Text.ToString & "\InterventoryApproved\" & lstPdfFile.SelectedItem
            System.Diagnostics.Process.Start(ruta)
        Catch ex As Exception
            MsgBox("Error: " + ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub lstPdfFile_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPdfFile.SelectedIndexChanged

    End Sub

    Private Sub GroupBox7_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGroup.SelectedIndexChanged

    End Sub

    Private Sub cmbGroup_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGroup.SelectionChangeCommitted
        Try
            FillSubGroup()
            'cm()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmbSubGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSubGroup.SelectedIndexChanged

    End Sub

    Private Sub cmbSubGroup_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSubGroup.SelectionChangeCommitted
        Try
            dgQuestion.Columns(0).Visible = False
            dgQuestion.Columns(1).Visible = False

            FillQuestion()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    Private Sub ValidarEnviromentanPollH()
        Try
            sValidarPoll = True
            If txtWaterColPoint.Text.ToString = "" Then
                MsgBox("Error in Water Colletion Point", MsgBoxStyle.Critical)
                sValidarPoll = False
                txtWaterColPoint.Focus()
                Exit Sub

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub ValidarEnvironmentPollC()
        Try
            svalidarPollC = False
            Dim i As Integer
            For i = 0 To dgQuestion.Rows.Count - 1
                If dgQuestion.Rows(i).Cells(3).Value Is Nothing Then
                    MsgBox("Error in: " & dgQuestion.Rows(i).Cells(2).Value.ToString)
                    svalidarPollC = False
                    Exit Sub
                Else
                    svalidarPollC = True
                End If

            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
       
    End Sub

    Private Sub btnAddEnvironmental_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddEnvironmental.Click
        Try
            ValidarEnviromentanPollH()

            If sValidarPoll = True Then
                If sEditPoll = "0" Then
                    oPlatform.sOpcion = "1"
                    oPlatform.sIDPH = ""
                    oPlatform.sPlatform = txtPlatform.Text.ToString
                    oPlatform.sWaterColPoint = txtWaterColPoint.Text.ToString
                    oPlatform.sDateDevelopment = dtpDevelopment.Value.ToString
                    oPlatform.sDateReview = dtpReview.Value.ToString
                    oPlatform.sConclutions = txtRecommendations.Text.ToString

                    oPlatform.dEastWC = Double.Parse(txtEastWC.Text.ToString())
                    oPlatform.dElevationWC = Double.Parse(txtElevationWC.Text.ToString())
                    oPlatform.dNorthWC = Double.Parse(txtNorthWC.Text.ToString())
                    oPlatform.sCoordinateSystemWC = txtCoordSystemWC.Text.ToString()

                    oRf.InsertTrans("Enviromental", "Insert", clsRf.sUser.ToString(), _
                    "Platform: " + txtPlatform.Text.ToString + ". " + _
                    "WaterColPoint : " + txtWaterColPoint.Text.ToString + ". " + _
                    "Date Development: " + dtpDevelopment.Value + ". " + _
                    "Date Review: " + dtpReview.Value + ". " + _
                    "Event Date " + Date.Now())


                    Dim sResp As String = oPlatform.DH_EnvironmentPollH_Insert()
                    If sResp = "NOK" Then
                        MsgBox("Poll Exists.(Platform and Date Of Development)", MsgBoxStyle.Information)
                        Exit Sub
                    End If
                    If sResp > 0 Then
                        sIDn = sResp
                        MsgBox("Poll created.", MsgBoxStyle.Information)
                        sEditPoll = "0"
                        'FillCompanyDrill()
                    End If
                Else

                    oPlatform.sOpcion = "2"
                    'oPlatform.sID = sIDn
                    oPlatform.sPlatform = txtPlatform.Text.ToString
                    oPlatform.sWaterColPoint = txtWaterColPoint.Text.ToString
                    oPlatform.sDateDevelopment = dtpDevelopment.Value.ToString
                    oPlatform.sDateReview = dtpReview.Value.ToString
                    oPlatform.sConclutions = txtRecommendations.Text.ToString


                    If txtEastWC.Text.ToString() <> "" Then
                        oPlatform.dEastWC = Double.Parse(txtEastWC.Text.ToString())
                    Else
                        oPlatform.dEastWC = Nothing
                    End If

                    If txtElevationWC.Text.ToString() <> "" Then
                        oPlatform.dElevationWC = Double.Parse(txtElevationWC.Text.ToString())
                    Else
                        oPlatform.dElevationWC = Nothing
                    End If

                    If txtNorthWC.Text.ToString() <> "" Then
                        oPlatform.dNorthWC = Double.Parse(txtNorthWC.Text.ToString())
                    Else
                        oPlatform.dNorthWC = Nothing
                    End If

                    If txtCoordSystemWC.Text.ToString() <> "" Then
                        oPlatform.sCoordinateSystemWC = txtCoordSystemWC.Text.ToString()
                    Else
                        oPlatform.sCoordinateSystemWC = Nothing
                    End If


                    oRf.InsertTrans("Enviromental", "Insert", clsRf.sUser.ToString(), _
                    "Platform: " + txtPlatform.Text.ToString + ". " + _
                    "WaterColPoint : " + txtWaterColPoint.Text.ToString + ". " + _
                    "Date Development: " + dtpDevelopment.Value + ". " + _
                    "Date Review: " + dtpReview.Value + ". " + _
                    "Event Date " + Date.Now())

                    Dim sResp As String = oPlatform.DH_EnvironmentPollH_Insert()
                    If sResp = "OK" Then
                        MsgBox("Poll Update.", MsgBoxStyle.Information)
                        sEditPoll = "1"
                    End If

                End If

            End If
            FillEnvironmentPoll()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dgEnvironmentPollQuery_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgEnvironmentPollQuery.CellClick
        Try
            If MsgBox("Press [YES] to edit, Press [NO] to add information in the poll", MsgBoxStyle.YesNo, "Edit?") = MsgBoxResult.No Then
                sValorPollH = dgEnvironmentPollQuery.CurrentRow.Cells.Item("id").Value
                dgQuestion.Rows.Clear()
                sEditPollC = "0"

                btnAdd.Enabled = True
                tbcEnvironment.SelectedTab = TabPage16
                cmbGroup.Enabled = True
                cmbSubGroup.Enabled = True
                btnSaveImpact.Enabled = True
            Else
                If dgEnvironmentPollQuery.CurrentRow.Cells(1).Value.ToString <> "" Then

                    sEditPoll = "1"

                    oPlatform.sIDPH = dgEnvironmentPollQuery.CurrentRow.Cells.Item("id").Value
                    sValorPollH = dgEnvironmentPollQuery.CurrentRow.Cells.Item("id").Value
                    oPlatform.sIDG = dgEnvironmentPollQuery.CurrentRow.Cells.Item("IDG").Value
                    oPlatform.sIDSG = dgEnvironmentPollQuery.CurrentRow.Cells.Item("IDSG").Value
                    Dim dtPoll As DataTable = oPlatform.getDH_Environment_Poll_Select()

                    txtWaterColPoint.Text = dtPoll.Rows(0)("WaterColPoint").ToString
                    dtpDevelopment.Value = dtPoll.Rows(0)("DateDevelopment").ToString
                    dtpReview.Value = dtPoll.Rows(0)("DateReview").ToString
                    txtRecommendations.Text = dtPoll.Rows(0)("Conclusions").ToString

                    txtEastWC.Text = dtPoll.Rows(0)("East").ToString
                    txtNorthWC.Text = dtPoll.Rows(0)("North").ToString
                    txtElevationWC.Text = dtPoll.Rows(0)("Elevation").ToString
                    txtCoordSystemWC.Text = dtPoll.Rows(0)("CoordinateSystem").ToString

                    cmbGroup.SelectedValue = dtPoll.Rows(0)("IDG").ToString
                    FillSubGroup()
                    cmbSubGroup.SelectedValue = dtPoll.Rows(0)("IDSG").ToString
                    FillQuestion()

                    For i = 0 To dtPoll.Rows.Count - 1
                        If dgQuestion.Rows(i).Cells(0).Value = dtPoll.Rows(i)("IDQ").ToString Then
                            If dtPoll.Rows(i)("OPT").ToString = 1 Then
                                dgQuestion.Rows(i).Cells(3).Value = "YES"
                            Else
                                If dtPoll.Rows(i)("OPT").ToString = 0 Then
                                    dgQuestion.Rows(i).Cells(3).Value = "NO"
                                Else
                                    dgQuestion.Rows(i).Cells(3).Value = "N/A"
                                End If
                            End If
                        End If
                        dgQuestion.Rows(i).Cells(4).Value = dtPoll.Rows(i)("Comments").ToString

                    Next

                    ' Limpiar

                    For i = 0 To dgImpact.Rows.Count - 1
                        dgImpact.Rows(0).Cells(2).Value = ""
                    Next


                    For i = 0 To dgImpact.Rows.Count - 1
                        oPlatform.sIDI = dgImpact.Rows(i).Cells(0).Value
                        oPlatform.sIDH = sValorPollH

                        Dim dtRegistro As DataTable = oPlatform.getDH_Environment_Poll_Impact_Select()

                        If dtRegistro.Rows.Count > 0 Then
                            sEditPollImpact = "1"
                            If dtRegistro.Rows(0)("IDO").ToString = 1 Then

                                dgImpact.Rows(i).Cells(2).Value = "Leve"
                            Else
                                If dtRegistro.Rows(0)("IDO").ToString = 2 Then
                                    dgImpact.Rows(i).Cells(2).Value = "Severo"
                                Else

                                    If dtRegistro.Rows(0)("IDO").ToString = 3 Then
                                        dgImpact.Rows(i).Cells(2).Value = "Crítico"
                                    Else
                                        dgImpact.Rows(i).Cells(2).Value = ""
                                    End If

                                End If
                            End If
                        Else
                            sEditPollImpact = "0"
                            dgImpact.Rows(i).Cells(2).Value = ""
                        End If
                    Next

                    btnAdd.Enabled = True
                    sEditPollC = "1"

                    tbcEnvironment.SelectedTab = TabPage16
                    cmbGroup.Enabled = False
                    cmbSubGroup.Enabled = False
                    btnSaveImpact.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        

    End Sub

    Private Sub dgQuestion_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgQuestion.CellClick

    End Sub

    Private Sub dgQuestion_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgQuestion.CellContentClick

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            ValidarEnvironmentPollC()
            Dim i As Integer
            Dim OK As Integer
            'Dim ValorOpt As Integer
            For i = 0 To dgQuestion.Rows.Count - 1
                If svalidarPollC = True Then
                    If sEditPollC = "0" Then
                        oPlatform.sOpcion = "1"
                        oPlatform.sID = ""
                        oPlatform.sIDPH = sValorPollH
                        oPlatform.sIDSG = cmbSubGroup.SelectedValue.ToString
                        oPlatform.sIDQ = dgQuestion.Rows(i).Cells(0).Value
                        If dgQuestion.Rows(i).Cells(3).Value = "YES" Then
                            oPlatform.sOpt = 1
                        Else
                            If dgQuestion.Rows(i).Cells(3).Value = "NO" Then
                                oPlatform.sOpt = 0
                            Else
                                oPlatform.sOpt = 2
                            End If
                        End If
                        'oPlatform.sOpt =gQuestion.Rows(i).Cells(3).Value
                        If dgQuestion.Rows(i).Cells(4).Value Is Nothing Then
                            oPlatform.sConclutions = ""
                        Else
                            oPlatform.sConclutions = dgQuestion.Rows(i).Cells(4).Value
                        End If


                        oRf.InsertTrans("Enviromental", "Insert", clsRf.sUser.ToString(), _
                        "Group : " + cmbGroup.SelectedText.ToString() + ". " + _
                        "SubGroup : " + cmbSubGroup.SelectedText.ToString() + ". " + _
                        "Question: " + dgQuestion.CurrentRow.Cells.Item("Question").Value + ". " + _
                        "Event Date " + Date.Now())

                        Dim sResp As String = oPlatform.DH_EnvironmentPollC_Insert()

                        If sResp = "NOK" Then
                            MsgBox("Question exist in this poll")
                            Exit Sub
                        End If

                        If sResp = "OK" Then
                            OK = 1
                        End If


                    Else
                        oPlatform.sOpcion = "2"
                        oPlatform.sID = ""
                        oPlatform.sIDPH = sValorPollH
                        oPlatform.sIDSG = cmbSubGroup.SelectedValue.ToString
                        oPlatform.sIDQ = dgQuestion.Rows(i).Cells(0).Value
                        If dgQuestion.Rows(i).Cells(3).Value = "YES" Then
                            oPlatform.sOpt = 1
                        Else
                            If dgQuestion.Rows(i).Cells(3).Value = "NO" Then
                                oPlatform.sOpt = 0
                            Else
                                oPlatform.sOpt = 2
                            End If
                        End If
                        'oPlatform.sOpt =gQuestion.Rows(i).Cells(3).Value
                        If dgQuestion.Rows(i).Cells(4).Value Is Nothing Then
                            oPlatform.sConclutions = ""
                        Else
                            oPlatform.sConclutions = dgQuestion.Rows(i).Cells(4).Value
                        End If

                        oRf.InsertTrans("Enviromental", "Insert", clsRf.sUser.ToString(), _
                        "Group : " + cmbGroup.SelectedText.ToString() + ". " + _
                        "SubGroup : " + cmbSubGroup.SelectedText.ToString() + ". " + _
                        "Question: " + dgQuestion.CurrentRow.Cells.Item("Question").Value + ". " + _
                        "Event Date " + Date.Now())

                        Dim sResp As String = oPlatform.DH_EnvironmentPollC_Insert()
                        OK = 2
                    End If
                    'OK = 3
                End If
            Next

            If OK = 1 Then
                MsgBox("Inserted OK")
            End If
            If OK = 2 Then
                MsgBox("Updated OK")
            End If
            FillEnvironmentPoll()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub dgEnvironmentPollQuery_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgEnvironmentPollQuery.CellContentClick

    End Sub
    Private Sub ValidarImpact()
        Try
            svalidarPollImpact = False
            Dim i As Integer
            For i = 0 To dgImpact.Rows.Count - 1
                If dgImpact.Rows(i).Cells(2).Value Is Nothing Then
                    MsgBox("Error in: " & dgImpact.Rows(i).Cells(1).Value.ToString)
                    svalidarPollImpact = False
                    Exit Sub
                Else
                    svalidarPollImpact = True
                End If

            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub btnSaveImpact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveImpact.Click
        Try
            ValidarImpact()
            sPollImpAdd = False
            sPollImpEdt = False
            If svalidarPollImpact = True Then

                If sEditPollImpact = "0" Then
                    For i = 0 To dgImpact.Rows.Count - 1
                        oPlatform.sOpcion = "1"
                        oPlatform.sIDI = dgImpact.Rows(i).Cells(0).Value
                        oPlatform.sIDPH = sValorPollH
                        If dgImpact.Rows(i).Cells(2).Value = "Leve" Then
                            oPlatform.sIDO = 1
                        Else
                            If dgImpact.Rows(i).Cells(2).Value = "Severo" Then
                                oPlatform.sIDO = 2
                            Else
                                oPlatform.sIDO = 3
                            End If
                        End If


                        Dim sResp As String = oPlatform.getDH_Environment_Poll_Impact_Add()
                        If sResp = "OK" Then
                            sPollImpAdd = True
                        End If
                    Next

                End If

                If sEditPollImpact = "1" Then
                    For i = 0 To dgImpact.Rows.Count - 1
                        oPlatform.sOpcion = "2"
                        oPlatform.sIDI = dgImpact.Rows(i).Cells(0).Value
                        oPlatform.sIDPH = sValorPollH
                        If dgImpact.Rows(i).Cells(2).Value = "Leve" Then
                            oPlatform.sIDO = 1
                        Else
                            If dgImpact.Rows(i).Cells(2).Value = "Severo" Then
                                oPlatform.sIDO = 2
                            Else
                                oPlatform.sIDO = 3
                            End If
                        End If


                        Dim sResp As String = oPlatform.getDH_Environment_Poll_Impact_Add()
                        If sResp = "OK" Then
                            sPollImpEdt = True
                        End If

                    Next
                End If

            End If

            If sPollImpAdd = True Then
                MsgBox("Inserted OK")
            End If
            If sPollImpEdt = True Then
                MsgBox("Updated OK")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Try
            frmEnvironmentReport.Platform = txtPlatform.Text.ToString
            frmEnvironmentReport.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    

    Private Sub btnPdfEnv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilePict.Click
        Try
            fDialog.Filter = "Archivo JPG|*.jpg"
            fDialog.Title = "Seleccione imagen"
            fDialog.ShowDialog()
            txtFilePict.Text = fDialog.FileName

            Dim oFi As New FileInfo(fDialog.FileName)
            Dim sExt As String = oFi.Extension.ToString()
            sFile = oFi.Name.Substring(0, oFi.Name.ToString().Length)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnAdFilePictEnv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdFilePictEnv.Click

        Try
            If txtHoleIDDrill1.Text <> "" Then

                If txtFilePict.Text <> "" Then


                    lstPictureEnv.Items.Clear()
                    Dim Source As String
                    Dim Destino As String
                    Dim ArchivosT As New DirectoryInfo(ConfigurationSettings.AppSettings("Ruta_PictEnv").ToString & txtHoleIDDrill1.Text.ToString)

                    Source = txtFilePict.Text
                    Destino = ConfigurationSettings.AppSettings("Ruta_PictEnv").ToString & txtHoleIDDrill1.Text.ToString & "\" & sFile
                    'MsgBox(Destino)

                    System.IO.File.Copy(Source, Destino, True)

                    'Registra auditoria'
                    oRf.InsertTrans("Environment", "Insert", clsRf.sUser.ToString(), _
                    "Hole ID: " + txtHoleIDDrill1.Text.ToString + ". " + _
                    "File: " + sFile + ". " + _
                    "Event Date " + Date.Now())

                    For Each file As FileInfo In ArchivosT.GetFiles()
                        If file.Name <> "Thumbs.db" Then
                            lstPictureEnv.Items.Add(file.Name)
                        End If
                    Next

                    txtFilePict.Text = ""

                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            lstPictureEnv.Items.Clear()
        End Try
    End Sub

    Private Sub ListarPictureEnv()
        Try
            lstPictureEnv.Items.Clear()
            Dim Source As String
            Dim Destino As String
            Dim ArchivosT As New DirectoryInfo(ConfigurationSettings.AppSettings("Ruta_PictEnv").ToString & txtHoleIDDrill1.Text.ToString)

            For Each file As FileInfo In ArchivosT.GetFiles()
                If file.Name <> "Thumbs.db" Then
                    lstPictureEnv.Items.Add(file.Name)
                End If
            Next
        Catch ex As Exception
            lstPictureEnv.Items.Clear()
        End Try

    End Sub


    Private Sub lstPictureEnv_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPictureEnv.DoubleClick

        

    End Sub

    Private Sub lstPictureEnv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPictureEnv.Click
        Try
            Dim ruta As String
            ruta = ConfigurationSettings.AppSettings("Ruta_PictEnv").ToString & txtHoleIDDrill1.Text.ToString & "\" & lstPictureEnv.SelectedItem
            pbPreviewPicRec.ImageLocation = ruta
            'System.Diagnostics.Process.Start(ruta)
        Catch ex As Exception
            MsgBox("Error: " + ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub


    Private Sub txtNorthWC_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNorthWC.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtEastWC_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEastWC.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtElevationWC_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtElevationWC.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub



End Class