Imports System.Configuration.ConfigurationSettings

Public Class frmAcctRpt
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmbRptType As System.Windows.Forms.ComboBox
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents crvAcctRpt As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbCompCode As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAcctRpt))
        Me.dtpStart = New System.Windows.Forms.DateTimePicker()
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker()
        Me.crvAcctRpt = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.cmbRptType = New System.Windows.Forms.ComboBox()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbCompCode = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'dtpStart
        '
        Me.dtpStart.CalendarTitleBackColor = System.Drawing.Color.SlateGray
        Me.dtpStart.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStart.Location = New System.Drawing.Point(16, 80)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(144, 23)
        Me.dtpStart.TabIndex = 1
        '
        'dtpEnd
        '
        Me.dtpEnd.CalendarTitleBackColor = System.Drawing.Color.SlateGray
        Me.dtpEnd.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEnd.Location = New System.Drawing.Point(16, 128)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.Size = New System.Drawing.Size(144, 23)
        Me.dtpEnd.TabIndex = 2
        '
        'crvAcctRpt
        '
        Me.crvAcctRpt.ActiveViewIndex = -1
        Me.crvAcctRpt.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvAcctRpt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvAcctRpt.Cursor = System.Windows.Forms.Cursors.Default
        Me.crvAcctRpt.Location = New System.Drawing.Point(176, 8)
        Me.crvAcctRpt.Name = "crvAcctRpt"
        Me.crvAcctRpt.Size = New System.Drawing.Size(832, 696)
        Me.crvAcctRpt.TabIndex = 4
        Me.crvAcctRpt.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'cmbRptType
        '
        Me.cmbRptType.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbRptType.Items.AddRange(New Object() {"[Select Report]", "Cash Receipt", "Sales Register"})
        Me.cmbRptType.Location = New System.Drawing.Point(16, 32)
        Me.cmbRptType.Name = "cmbRptType"
        Me.cmbRptType.Size = New System.Drawing.Size(144, 24)
        Me.cmbRptType.TabIndex = 0
        Me.cmbRptType.Text = "[Select Report]"
        '
        'btnPreview
        '
        Me.btnPreview.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.btnPreview.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPreview.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPreview.Location = New System.Drawing.Point(16, 216)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(144, 26)
        Me.btnPreview.TabIndex = 3
        Me.btnPreview.Text = "&Display"
        Me.btnPreview.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Report Type"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Start Date"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(16, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "End Date"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(16, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Company Code"
        '
        'cmbCompCode
        '
        Me.cmbCompCode.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCompCode.Items.AddRange(New Object() {"ALL", "SBITC", "ISI"})
        Me.cmbCompCode.Location = New System.Drawing.Point(16, 176)
        Me.cmbCompCode.Name = "cmbCompCode"
        Me.cmbCompCode.Size = New System.Drawing.Size(144, 24)
        Me.cmbCompCode.TabIndex = 10
        Me.cmbCompCode.Text = "ALL"
        '
        'frmAcctRpt
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightSlateGray
        Me.ClientSize = New System.Drawing.Size(1016, 710)
        Me.Controls.Add(Me.cmbCompCode)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnPreview)
        Me.Controls.Add(Me.cmbRptType)
        Me.Controls.Add(Me.crvAcctRpt)
        Me.Controls.Add(Me.dtpEnd)
        Me.Controls.Add(Me.dtpStart)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAcctRpt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Accounting Reports"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Objects 
    'Sales Invoice
    Private dtabINVICT As dsAcctRpt.INVICTDataTable
    Private dtabINVCYB As DataTable
    Private dtabPayDtl As DataTable
    Private dtabSales As dsAcctRpt.SalesDataTable
    'Cash Register
    Private dtabCCRPay As DataTable
    Private dtabCCRCyx As DataTable
    Private dtabCCRDtl As DataTable
    Private dtabCYMPay As DataTable
    Private dtabCYMGps As DataTable
    Private dtabInvPayHdr As DataTable
    Private dtabInvPayDtl As DataTable
    Private dtabInvDtl As DataTable
    Private dtabCash As dsAcctRpt.CashDataTable

    'Configuration Settings
    Private strServer As String = CType(AppSettings("Server"), String)
    Private strDataBase As String = CType(AppSettings("Database"), String)

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Dim clsAcctRpt As New clsAcctRpt(strServer, strDataBase)
        Dim strDocNo As String = ""
        Dim strChkNo As String = ""

        Cursor = Cursors.WaitCursor
#Region "Cash Register"
        '''''''''''''CASH REGISTER''''''''''''''''
        If Trim(cmbRptType.Text) = "Cash Receipt" Then
#Region "Export and Special Services"
            Dim lngExpAmt As Double = 0
            Dim dblWghAmt As Double
            Dim lngSpcAmt As Double = 0
            Dim dblExpTotal As Double = 0
            Dim dblExpPay As Double = 0
            Dim dblExpCsh As Double = 0
            Dim dblExpChk As Double = 0
            Dim dblExpAdr As Double = 0
            Dim charADR As Char = ""
            Dim ADRno As String
            dtabCash = New dsAcctRpt.CashDataTable

            'Export and Special Services
            dtabCCRPay = New DataTable
            dtabCCRPay = clsAcctRpt.Get_CCRPay(dtpStart.Text & " 00:00:00 AM", dtpEnd.Text & " 11:58:59 PM", cmbCompCode.Text.Trim)

            If dtabCCRPay.Rows.Count > 0 Then
                Dim lngCCRPay As Long = 0

                Do While lngCCRPay < dtabCCRPay.Rows.Count
                    'If clsAcctRpt.Chk_CAN_UG(dtabCCRPay.Rows(lngCCRPay)("ccrtyp"), dtabCCRPay.Rows(lngCCRPay)("refnum")) = False Then
                    If dtabCCRPay.Rows(lngCCRPay)("ccrtyp") = 1 Then
                            Dim intCCRCyx As Integer = 0

                            'Export
                            dtabCCRCyx = clsAcctRpt.Get_CCRCyx(dtabCCRPay.Rows(lngCCRPay)("refnum"))
                            If dtabCCRCyx.Rows.Count > 0 Then
                                strChkNo = ""
                                'Get Cheque Nos.
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno1").ToString) <> "" Then
                                    strChkNo = Trim(dtabCCRPay.Rows(lngCCRPay)("chkno1").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno2").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno2").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno3").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno3").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno4").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno4").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno5").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno5").ToString)
                                End If

                                strDocNo = ""
                                'Get CCR No. series
                                If dtabCCRCyx.Rows.Count = 1 Then
                                    strDocNo = "CCMR " & dtabCCRCyx.Rows(0)("ccrnum").ToString
                                ElseIf Trim(dtabCCRCyx.Rows(0)("ccrnum")) = Trim(dtabCCRCyx.Rows(dtabCCRCyx.Rows.Count - 1)("ccrnum")) Then
                                    strDocNo = "CCMR " & Trim(dtabCCRCyx.Rows(0)("ccrnum").ToString)
                                Else
                                    strDocNo = "CCMR " & Trim(dtabCCRCyx.Rows(0)("ccrnum").ToString) & " - " & Trim(dtabCCRCyx.Rows(dtabCCRCyx.Rows.Count - 1)("ccrnum").ToString)
                                End If


                                'Get Export Amount
                                lngExpAmt = 0
                                dblExpCsh = 0
                                dblExpChk = 0
                                dblWghAmt = 0
                                charADR = ""
                                Do While intCCRCyx < dtabCCRCyx.Rows.Count
                                    lngExpAmt = lngExpAmt +
                                                dtabCCRCyx.Rows(intCCRCyx)("whfamt") +
                                                dtabCCRCyx.Rows(intCCRCyx)("arramt") +
                                                dtabCCRCyx.Rows(intCCRCyx)("ovzamt") +
                                                dtabCCRCyx.Rows(intCCRCyx)("dgramt") +
                                                dtabCCRCyx.Rows(intCCRCyx)("arrvat") -
                                                dtabCCRCyx.Rows(intCCRCyx)("arrtax")

                                    dblWghAmt += dtabCCRCyx.Rows(intCCRCyx)("wghamt") 'Sum Weighing
                                    intCCRCyx += 1
                                Loop
                            End If
                            'Get Amount Paid Total
                            dblExpAdr = dtabCCRPay.Rows(lngCCRPay)("adramt")
                            dblExpCsh = dtabCCRPay.Rows(lngCCRPay)("cshamt")

                            dblExpChk = dtabCCRPay.Rows(lngCCRPay)("chkamt1") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt2") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt3") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt4") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt5")

                            dblExpTotal = dblExpCsh + dblExpChk


                            'Check ADR
                            If dtabCCRPay.Rows(lngCCRPay)("adramt") > 0 Then charADR = "*"

                            'Add cash data

                            'PRNH 08102017 - Retrieve Exporter if Cusnam is blank

                            Dim cusNameExp As String = Trim(dtabCCRPay.Rows(lngCCRPay)("cusnam"))
                            If cusNameExp = "" Then
                                cusNameExp = dtabCCRCyx.Rows(0)("exprtr")
                            End If

                            On Error Resume Next
                            With dtabCCRPay
                                ADRno = ""
                                Dim ADR1 As String = ZeroVoid(.Rows(lngCCRPay)("adrnum"))
                                Dim ADR2 As String = ZeroVoid(.Rows(lngCCRPay)("adrnum2"))
                                Dim ADR3 As String = ZeroVoid(.Rows(lngCCRPay)("adrnum3"))


                                If Len(ADR1) > 0 Then ADRno &= ADR1
                                If Len(ADR2) > 0 Then ADRno &= IIf(Len(ADR1) > 0, " / ", "") & ADR2
                                If Len(ADR3) > 0 Then ADRno &= IIf(Len(ADR2) > 0, " / ", "") & ADR3
                            End With
                            On Error GoTo 0

                            'Add_CashData(dtabCCRPay.Rows(lngCCRPay)("sysdttm"), _
                            '             strDocNo, "", Trim(strChkNo), _
                            '             dtabCCRPay.Rows(lngCCRPay)("cusnam"), _
                            '             lngExpAmt, 0, lngExpAmt, 0, 0, 0, 0)
                            Add_CashData(dtabCCRPay.Rows(lngCCRPay)("sysdttm"),
                                        strDocNo, ADRno, Trim(strChkNo),
                                        cusNameExp, charADR, dblExpTotal, dblExpCsh, dblExpChk, dblExpAdr, 0, lngExpAmt, 0, 0, dblWghAmt, 0, 0, 0, dtabCCRCyx.Rows(0)("CompanyCode"))


                        Else 'Special Services
                            Dim dblSPLTotal As Double = 0
                            Dim dblSPLCsh As Double = 0
                            Dim dblSPLChk As Double = 0
                            Dim dblSPLAdr As Double = 0
                            Dim intCCRCys As Integer = 0
                            Dim dblCIM As Double = 0
                            Dim dblCEX As Double = 0
                            Dim dblST As Double = 0
                            Dim dblRFR As Double = 0
                            Dim dblEQP As Double = 0
                            Dim dblSS As Double = 0
                            Dim dblSV As Double = 0
                            Dim ctr As Integer = 0

                            dtabCCRDtl = clsAcctRpt.Get_CCRDtl(dtabCCRPay.Rows(lngCCRPay)("refnum"))
                            'If dtabCCRPay.Rows(lngCCRPay)("refnum") = 107491 Then Stop

                            If dtabCCRDtl.Rows.Count > 0 Then

                                strChkNo = ""
                                'Get Cheque Nos.
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno1").ToString) <> "" Then
                                    strChkNo = Trim(dtabCCRPay.Rows(lngCCRPay)("chkno1").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno2").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno2").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno3").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno3").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno4").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno4").ToString)
                                End If
                                If Trim(dtabCCRPay.Rows(lngCCRPay)("chkno5").ToString) <> "" Then
                                    strChkNo += "," & Trim(dtabCCRPay.Rows(lngCCRPay)("chkno5").ToString)
                                End If

                                strDocNo = ""
                                'Get CCR No. series
                                If dtabCCRDtl.Rows.Count = 1 Then
                                    strDocNo = "CCMR " & dtabCCRDtl.Rows(0)("ccrnum").ToString
                                ElseIf Trim(dtabCCRDtl.Rows(0)("ccrnum")) = Trim(dtabCCRDtl.Rows(dtabCCRDtl.Rows.Count - 1)("ccrnum")) Then
                                    strDocNo = "CCMR " & Trim(dtabCCRDtl.Rows(0)("ccrnum").ToString)
                                Else
                                    strDocNo = "CCMR " & Trim(dtabCCRDtl.Rows(0)("ccrnum").ToString) & " - " & Trim(dtabCCRDtl.Rows(dtabCCRDtl.Rows.Count - 1)("ccrnum").ToString)
                                End If

                                'Get Special Services Amount
                                lngSpcAmt = 0
                                charADR = ""
                                Do While intCCRCys < dtabCCRDtl.Rows.Count
                                    lngSpcAmt = dtabCCRDtl.Rows(intCCRCys)("amt") +
                                                dtabCCRDtl.Rows(intCCRCys)("vatamt") +
                                                dtabCCRDtl.Rows(intCCRCys)("ovzamt") +
                                                dtabCCRDtl.Rows(intCCRCys)("dgramt") -
                                                dtabCCRDtl.Rows(intCCRCys)("wtax")

                                    'Add cash data 'Jaspher by Bill Type Type
                                    Select Case dtabCCRDtl.Rows(intCCRCys)("cyr_biltyp").ToString
                                        Case "CB" 'Cargo Billing IMP/EsXP
                                            If InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "IM") > 0 Then
                                                dblCIM += lngSpcAmt
                                            ElseIf InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "EX") > 0 Then
                                                dblCEX += lngSpcAmt
                                            Else
                                                dblCIM += lngSpcAmt
                                            End If
                                        Case "ST" 'Storage
                                            dblST += lngSpcAmt
                                        Case "MC" 'Miscellaneous Charges
                                            If InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "RF") > 0 Then 'Reefer
                                                dblRFR += lngSpcAmt
                                            ElseIf InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "V") = 1 Then 'Vessel Billing
                                                dblSV += lngSpcAmt
                                            Else
                                                dblEQP += lngSpcAmt
                                            End If
                                        Case "SS" 'Stripping/Stuffing
                                            dblSS += lngSpcAmt
                                        Case "VB" 'Vessel Billing
                                            dblSV += lngSpcAmt
                                        Case Else 'Check in Chargetype
                                            If InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "CB") = 1 Then 'Cargo Billing IMP/EXP
                                                If InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "IM") > 0 Then
                                                    dblCIM += lngSpcAmt
                                                ElseIf InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "EX") > 0 Then
                                                    dblCEX += lngSpcAmt
                                                Else
                                                    dblCIM += lngSpcAmt
                                                End If
                                            ElseIf InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "STRIP") = 3 Then 'Stripping / Stuffing
                                                dblSS += lngSpcAmt
                                            ElseIf InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "ST") = 1 Then 'Storage
                                                dblST += lngSpcAmt
                                            ElseIf InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "V") = 1 Then 'Vessel Billing
                                                dblSV += lngSpcAmt
                                            ElseIf InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "MC") = 1 Then 'Miscellaneous Charges
                                                If InStr(dtabCCRDtl.Rows(intCCRCys)("chargetyp").ToString, "RF") > 0 Then 'Reefer
                                                    dblRFR += lngSpcAmt
                                                Else
                                                    dblEQP += lngSpcAmt
                                                End If
                                            End If

                                    End Select
                                    intCCRCys += 1
                                Loop
                            End If


                            'Get Amount Paid in Cash and Cheque
                            dblSPLAdr = dtabCCRPay.Rows(lngCCRPay)("adramt")

                            dblSPLCsh = dtabCCRPay.Rows(lngCCRPay)("cshamt")

                            dblSPLChk = dtabCCRPay.Rows(lngCCRPay)("chkamt1") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt2") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt3") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt4") +
                                         dtabCCRPay.Rows(lngCCRPay)("chkamt5")

                            dblSPLTotal = dblSPLCsh + dblSPLChk

                            'Check ADR
                            If dtabCCRPay.Rows(lngCCRPay)("adramt") > 0 Then charADR = "*"

                            On Error Resume Next
                            With dtabCCRPay
                                ADRno = ""
                                Dim ADR1 As String = ZeroVoid(.Rows(lngCCRPay)("adrnum"))
                                Dim ADR2 As String = ZeroVoid(.Rows(lngCCRPay)("adrnum2"))
                                Dim ADR3 As String = ZeroVoid(.Rows(lngCCRPay)("adrnum3"))


                                If Len(ADR1) > 0 Then ADRno &= ADR1
                                If Len(ADR2) > 0 Then ADRno &= IIf(Len(ADR1) > 0, " / ", "") & ADR2
                                If Len(ADR3) > 0 Then ADRno &= IIf(Len(ADR2) > 0, " / ", "") & ADR3
                            End With
                            On Error GoTo 0

                            'add Cash Data
                            Add_CashData(dtabCCRPay.Rows(lngCCRPay)("sysdttm"),
                            strDocNo, ADRno, Trim(strChkNo),
                            dtabCCRPay.Rows(lngCCRPay)("cusnam"), charADR, dblSPLTotal, dblSPLCsh, dblSPLChk, dblSPLAdr, dblCIM, dblCEX, dblST, dblRFR, dblEQP, dblSS, 0, dblSV, dtabCCRDtl.Rows(0)("CompanyCode"))

                        End If
                    'End If
                    lngCCRPay += 1
                Loop
            End If
#End Region
#Region "Import"
            'Import
            dtabCYMPay = New DataTable
            dtabCYMPay = clsAcctRpt.Get_CYMPay(dtpStart.Text & " 00:00:00 AM", dtpEnd.Text & " 11:58:59 PM", cmbCompCode.Text.Trim)

            If dtabCYMPay.Rows.Count > 0 Then
                Dim lngCYMAmt As Double = 0
                Dim dblCYMSTO As Double = 0
                Dim dblCYMRFR As Double = 0
                Dim lngCYMPay As Long = 0
                Dim dblCYMTotal As Double = 0
                Dim dblCYMCsh As Double = 0
                Dim dblCYMChk As Double = 0
                Dim dblCYMAdr As Double = 0

                Do While lngCYMPay < dtabCYMPay.Rows.Count
                    If clsAcctRpt.Chk_CAN_UG(3, dtabCYMPay.Rows(lngCYMPay)("refnum")) = False Then
                        Dim intCCRCym As Integer = 0

                        dtabCYMGps = clsAcctRpt.Get_CYMGps(dtabCYMPay.Rows(lngCYMPay)("refnum"))
                        If dtabCYMGps.Rows.Count > 0 Then
                            strChkNo = ""
                            'Get Cheque Nos.
                            If Trim(dtabCYMPay.Rows(lngCYMPay)("chkno1").ToString) <> "0" Then
                                strChkNo = Trim(dtabCYMPay.Rows(lngCYMPay)("chkno1").ToString)
                            End If
                            If Trim(dtabCYMPay.Rows(lngCYMPay)("chkno2").ToString) <> "0" Then
                                strChkNo += "," & Trim(dtabCYMPay.Rows(lngCYMPay)("chkno2").ToString)
                            End If
                            If Trim(dtabCYMPay.Rows(lngCYMPay)("chkno3").ToString) <> "0" Then
                                strChkNo += "," & Trim(dtabCYMPay.Rows(lngCYMPay)("chkno3").ToString)
                            End If
                            If Trim(dtabCYMPay.Rows(lngCYMPay)("chkno4").ToString) <> "0" Then
                                strChkNo += "," & Trim(dtabCYMPay.Rows(lngCYMPay)("chkno4").ToString)
                            End If
                            If Trim(dtabCYMPay.Rows(lngCYMPay)("chkno5").ToString) <> "0" Then
                                strChkNo += "," & Trim(dtabCYMPay.Rows(lngCYMPay)("chkno5").ToString)
                            End If
                            strDocNo = ""
                            'Get Gps No. series
                            If dtabCYMGps.Rows.Count = 1 Then
                                strDocNo = "CCR " & dtabCYMGps.Rows(0)("gpsnum").ToString
                            ElseIf Trim(dtabCYMGps.Rows(0)("gpsnum")) = Trim(dtabCYMGps.Rows(dtabCYMGps.Rows.Count - 1)("gpsnum")) Then
                                strDocNo = "CCR " & Trim(dtabCYMGps.Rows(0)("gpsnum").ToString)
                            Else
                                strDocNo = "CCR " & Trim(dtabCYMGps.Rows(0)("gpsnum").ToString) & " - " & Trim(dtabCYMGps.Rows(dtabCYMGps.Rows.Count - 1)("gpsnum").ToString)
                            End If
                            'Get Import Amount
                            lngCYMAmt = 0
                            dblCYMSTO = 0
                            dblCYMRFR = 0

                            charADR = ""
                            Do While intCCRCym < dtabCYMGps.Rows.Count
                                lngCYMAmt = lngCYMAmt +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("udstoamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstoamt")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("udstovat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstovat")) -
                                        IIf(dtabCYMGps.Rows(intCCRCym)("udstotax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstotax")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("arramt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arramt")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("whfamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("whfamt")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("wghamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghamt")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("stovat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stovat")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("arrvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arrvat")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("wghvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghvat")) +
                                        IIf(dtabCYMGps.Rows(intCCRCym)("rfrvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rfrvat")) -
                                        IIf(dtabCYMGps.Rows(intCCRCym)("stotax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stotax")) -
                                        IIf(dtabCYMGps.Rows(intCCRCym)("arrtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arrtax")) -
                                        IIf(dtabCYMGps.Rows(intCCRCym)("wghtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghtax")) -
                                        IIf(dtabCYMGps.Rows(intCCRCym)("rfrtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rfrtax"))

                                dblCYMSTO = dblCYMSTO + IIf(dtabCYMGps.Rows(intCCRCym)("stoamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stoamt"))
                                dblCYMRFR = dblCYMRFR + IIf(dtabCYMGps.Rows(intCCRCym)("rframt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rframt"))
                                'IIf(dtabCYMGps.Rows(intCCRCym)("dgramt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("dgramt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("udstoamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstoamt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("udstovat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstovat")) - _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("udstotax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstotax")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("ovzamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("ovzamt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("stoamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stoamt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("arramt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arramt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("whfamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("whfamt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("wghamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghamt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("rframt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rframt")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("stovat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stovat")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("arrvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arrvat")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("wghvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghvat")) + _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("rfrvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rfrvat")) - _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("stotax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stotax")) - _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("arrtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arrtax")) - _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("wghtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghtax")) - _
                                'IIf(dtabCYMGps.Rows(intCCRCym)("rfrtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rfrtax"))
                                intCCRCym += 1
                            Loop
                            'Get Amount paid in Cash and Cheque
                            dblCYMAdr = dtabCYMPay.Rows(lngCYMPay)("adramt")
                            dblCYMCsh = dtabCYMPay.Rows(lngCYMPay)("cshamt")

                            dblCYMChk = dtabCYMPay.Rows(lngCYMPay)("chkamt1") +
                                         dtabCYMPay.Rows(lngCYMPay)("chkamt2") +
                                         dtabCYMPay.Rows(lngCYMPay)("chkamt3") +
                                         dtabCYMPay.Rows(lngCYMPay)("chkamt4") +
                                         dtabCYMPay.Rows(lngCYMPay)("chkamt5")
                            dblCYMTotal = dblCYMCsh + dblCYMChk
                            'Check ADR
                            If dtabCYMPay.Rows(lngCYMPay)("adramt") > 0 Then charADR = "*"
                            'Add cash data

                            'PRNH 08102017 - Retrieve Consignee if Cusnam is blank
                            Dim custImp As String = dtabCYMPay.Rows(lngCYMPay)("cusnam")
                            If custImp = "" Then
                                custImp = dtabCYMGps.Rows(0)("cnsgne")
                            End If

                            On Error Resume Next
                            With dtabCYMPay
                                ADRno = ""
                                Dim ADR1 As String = ZeroVoid(.Rows(lngCYMPay)("adrnum"))
                                Dim ADR2 As String = ZeroVoid(.Rows(lngCYMPay)("adrnum2"))
                                Dim ADR3 As String = ZeroVoid(.Rows(lngCYMPay)("adrnum3"))


                                If Len(ADR1) > 0 Then ADRno &= ADR1
                                If Len(ADR2) > 0 Then ADRno &= IIf(Len(ADR1) > 0, " / ", "") & ADR2
                                If Len(ADR3) > 0 Then ADRno &= IIf(Len(ADR2) > 0, " / ", "") & ADR3
                            End With
                            On Error GoTo 0

                            'Add_CashData(dtabCYMPay.Rows(lngCYMPay)("sysdttm"), _
                            '             strDocNo, "", Trim(strChkNo), _
                            '             dtabCYMPay.Rows(lngCYMPay)("cusnam"), lngCYMAmt, lngCYMAmt, 0, 0, 0, 0, 0)

                            Add_CashData(dtabCYMPay.Rows(lngCYMPay)("sysdttm"),
                                         strDocNo, ADRno, Trim(strChkNo),
                                         custImp, charADR, dblCYMTotal, dblCYMCsh, dblCYMChk, dblCYMAdr, lngCYMAmt, 0, dblCYMSTO, dblCYMRFR, 0, 0, 0, 0, dtabCYMGps.Rows(0)("CompanyCode"))
                        End If
                    End If
                    lngCYMPay += 1
                Loop
            End If
#End Region
#Region "Invoice"
            'Invoice
            dtabInvPayHdr = New DataTable
            dtabInvPayHdr = clsAcctRpt.Get_InvPayHdr(dtpStart.Text & " 00:00:00 AM", dtpEnd.Text & " 11:58:59 PM")
            If dtabInvPayHdr.Rows.Count > 0 Then
                Dim lngInvAmt As Double = 0
                Dim rowInvPayHdr As Integer = 0

                Do While rowInvPayHdr < dtabInvPayHdr.Rows.Count
                    strChkNo = ""
                    'Get Cheque Nos.
                    If Trim(dtabInvPayHdr.Rows(rowInvPayHdr)("checkno1").ToString) <> "0" Then
                        strChkNo = Trim(dtabInvPayHdr.Rows(rowInvPayHdr)("checkno1").ToString)
                    End If
                    If Trim(dtabInvPayHdr.Rows(rowInvPayHdr)("checkno2").ToString) <> "0" Then
                        strChkNo += "," & Trim(dtabInvPayHdr.Rows(rowInvPayHdr)("checkno2").ToString)
                    End If

                    Dim intInvPayDtl As Integer = 0
                    Dim strCusName As String = ""

                    'If dtabInvPayHdr.Rows(rowInvPayHdr)("ornum") = 1579 Then Stop

                    dtabInvPayDtl = clsAcctRpt.Get_InvPayDtl(dtabInvPayHdr.Rows(rowInvPayHdr)("ornum"))

                    If dtabInvPayDtl.Rows.Count > 0 Then
                        Do While intInvPayDtl < dtabInvPayDtl.Rows.Count
                            'If clsAcctRpt.Chk_CAN_UG(4, dtabInvPayHdr.Rows(rowInvPayHdr)("ornum")) = False Then
                            If clsAcctRpt.Chk_CAN_UG(4, dtabInvPayDtl.Rows(intInvPayDtl)("invnum")) = False Then
                                'Get Invoice Amount
                                lngInvAmt = 0
                                lngInvAmt = IIf(dtabInvPayDtl.Rows(intInvPayDtl)("payamt") Is DBNull.Value, 0, dtabInvPayDtl.Rows(intInvPayDtl)("payamt"))
                                'Get Invoice Number
                                strDocNo = ""
                                strDocNo = Trim(dtabInvPayDtl.Rows(intInvPayDtl)("invnum").ToString)
                                'Get Customer Name
                                strCusName = ""
                                strCusName = clsAcctRpt.Get_CustomerName(dtabInvPayHdr.Rows(rowInvPayHdr)("cuscde"))

                                'Check if account is an AR(Accounts Receivablle)
                                If clsAcctRpt.Is_Invoice_AR(strDocNo, dtpStart.Text & " 00:00:00 AM", dtpEnd.Text & " 11:58:59 PM") = True Then
                                    'Add cash data
                                    Add_CashData(dtabInvPayHdr.Rows(rowInvPayHdr)("ORDate"),
                                                 Trim("INV " & strDocNo), dtabInvPayHdr.Rows(rowInvPayHdr)("ORNum"), Trim(strChkNo),
                                                 strCusName, charADR, lngInvAmt, 0, 0, 0, 0, 0, 0, 0, 0, 0, lngInvAmt, 0, dtabInvPayDtl.Rows(intInvPayDtl)("CompanyCode"))
                                Else
                                    'Get invoice data
                                    Dim lngInvRefNo As Long = 0
                                    lngInvRefNo = clsAcctRpt.Get_InvRefNum(strDocNo)

                                    'Charges
                                    Dim dblVC, dblVC1 As Double  'Vessel Charges
                                    Dim dblCIM, dblCIM1 As Double 'Import Cargoes
                                    Dim dblCEX, dblCEX1 As Double 'Export Cargoes
                                    Dim dblST, dblST1 As Double 'Storage
                                    Dim dblMC, dblMC1 As Double 'Reefer
                                    Dim dblSS, dblSS1 As Double 'Stripping/Stuffing
                                    Dim dblOth, dblOth1 As Double 'Other Charges
                                    'Invoice Amount
                                    Dim dblInvAmt As Double = 0
                                    'Invoice PayDate

                                    'Initialize charges
                                    dblVC = 0 : dblVC1 = 0
                                    dblCIM = 0 : dblCIM1 = 0
                                    dblCEX = 0 : dblCEX1 = 0
                                    dblST = 0 : dblST1 = 0
                                    dblMC = 0 : dblMC1 = 0
                                    dblSS = 0 : dblSS1 = 0
                                    dblOth = 0 : dblOth1 = 0

                                    'Get the invoice amount on a per Billing Type basis------------
                                    dtabINVCYB = New DataTable
                                    dtabINVCYB = clsAcctRpt.Get_INVCYB(lngInvRefNo)


                                    If dtabINVCYB.Rows.Count > 0 Then
                                        Dim intCybCtr As Integer = 0

                                        Do While intCybCtr < dtabINVCYB.Rows.Count
                                            Dim CCRNUM As String = dtabINVCYB.Rows(intCybCtr)("rtedsc").ToString
                                            Dim CNTNUM As String = CCRNUM

                                            Select Case Trim(dtabINVCYB.Rows(intCybCtr)("cyr_biltyp").ToString)
                                                Case "CB" 'Cargoes
                                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "EX") > 0 Then 'Export Cargoes

                                                        CCRNUM = Mid(CCRNUM, 'Get CCRNUM
                                                                 InStr(CCRNUM, "R#") + 2, 'Start Position at 2 Characters from R
                                                                 InStr(CCRNUM, " ") - (InStr(CCRNUM, "R#") + 2)) 'Get Length by finding the Space ' ' and then Substracting the Starting Position

                                                        CNTNUM = Mid(CNTNUM, 'Get CCRNUM
                                                                 InStr(CNTNUM, " ") + 1, 'Start Position at 2 Characters from Space
                                                                 InStr(CNTNUM, "-") - (InStr(CNTNUM, " ") + 1)) 'Get Length by finding the Dash '-' and then Substracting the Starting Position

                                                        dtabInvDtl = clsAcctRpt.Get_InvDtl(CCRNUM, dtabINVCYB.Rows(intCybCtr)("rtecde"), CNTNUM)

                                                        dblCEX += CDbl(dtabInvDtl.Rows(0)("ExportCharges").ToString)
                                                        dblOth += CDbl(dtabInvDtl.Rows(0)("wghamt").ToString)

                                                        'ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "IM") > 0 Then 'Import Cargoes
                                                    Else
                                                        dblCIM += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    End If
                                                Case "ST" 'Storage
                                                    dblST += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                Case "MC" 'Miscellaneous Charges
                                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "RF") > 0 Then 'Reefer
                                                        dblMC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "V") = 1 Then 'Vessel Charge
                                                        dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    Else
                                                        dblOth += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString) 'Equipment Rental
                                                    End If
                                                Case "SS" 'Stripping/Stuffing
                                                    dblSS += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                Case "VB" 'Vessel Charges
                                                    dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                Case Else 'Other Charges
                                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "CB") = 1 Then 'Cargo Billing IMP/EXP
                                                        If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "IM") > 0 Then
                                                            dblCIM += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                            'ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "EX") > 0 Then
                                                        Else
                                                            dblCEX += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                        End If
                                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "STRIP") = 1 Then 'Stripping/Stuffing
                                                        dblSS += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "ST") = 1 Then 'Storage
                                                        dblST += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "MC") = 1 Then 'Miscellaneous Charges
                                                        If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "RF") > 0 Then 'Reefer
                                                            dblMC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                        Else
                                                            dblOth += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString) 'Equipment Rental
                                                        End If
                                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "V") = 1 Then 'Vessel Billing
                                                        dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    End If
                                            End Select
                                            intCybCtr += 1
                                        Loop
                                    End If
                                    '--------------------------------------------------------------

                                    Dim dblPayAmt As Double = 0

                                    dblPayAmt = lngInvAmt

                                    'Recompute charges
                                    If dblVC > 0 Then 'Vessel Charges
                                        If dblVC >= dblPayAmt Then
                                            dblVC1 = dblPayAmt
                                        Else
                                            dblVC1 = dblVC
                                        End If
                                    End If
                                    If dblCEX > 0 Then 'Export Cargoes
                                        If (dblVC + dblCEX) >= dblPayAmt Then
                                            dblCEX1 = dblPayAmt - (dblVC1)
                                        Else
                                            dblCEX1 = dblCEX
                                        End If
                                    End If
                                    If dblCIM > 0 Then 'Import Cargoes
                                        If (dblVC + dblCEX + dblCIM) >= dblPayAmt Then
                                            dblCIM1 = dblPayAmt - (dblVC1 + dblCEX1)
                                        Else
                                            dblCIM1 = dblCIM
                                        End If
                                    End If
                                    If dblST > 0 Then 'Storage
                                        If (dblVC + dblCEX + dblCIM + dblST) >= dblPayAmt Then
                                            dblST1 = dblPayAmt - (dblVC1 + dblCEX1 + dblCIM1)
                                        Else
                                            dblST1 = dblST
                                        End If
                                    End If
                                    If dblMC > 0 Then 'Reefer
                                        If (dblVC + dblCEX + dblCIM + dblST + dblMC) >= dblPayAmt Then
                                            dblMC1 = dblPayAmt - (dblVC1 + dblCEX1 + dblCIM1 + dblST1)
                                        Else
                                            dblMC1 = dblMC
                                        End If
                                    End If
                                    If dblSS > 0 Then 'Stripping/Stuffing
                                        If (dblVC + dblCEX + dblCIM + dblST + dblMC + dblSS) >= dblPayAmt Then
                                            dblSS1 = dblPayAmt - (dblVC1 + dblCEX1 + dblCIM1 + dblST1 + dblMC1)
                                        Else
                                            dblSS1 = dblSS
                                        End If
                                    End If
                                    If dblOth > 0 Then 'Others
                                        If (dblVC + dblCEX + dblCIM + dblST + dblMC + dblSS + dblOth) >= dblPayAmt Then
                                            dblOth1 = dblPayAmt - (dblVC1 + dblCEX1 + dblCIM1 + dblST1 + dblMC1 + dblSS1)
                                        Else
                                            dblOth1 = dblOth
                                        End If
                                    End If

                                    'Add cash data
                                    Add_CashData(dtabInvPayHdr.Rows(rowInvPayHdr)("ORDate"),
                                                 Trim("INV " & strDocNo), dtabInvPayHdr.Rows(rowInvPayHdr)("ORNum"), Trim(strChkNo),
                                                 strCusName, charADR, lngInvAmt, 0, 0, 0, dblCIM1, dblCEX1, dblST1, dblMC1, dblOth1, dblSS1, 0, dblVC1,
                                                 dtabINVCYB.Rows(0)("CompanyCode"))
                                End If
                            End If
                            intInvPayDtl += 1
                        Loop
                    End If
                    rowInvPayHdr += 1
                Loop
            End If

            clsAcctRpt.DisConnect()
            clsAcctRpt = Nothing

            'Call Display Report
            Dim rptCash As New rptCash

            rptCash.SetDataSource(dataTable:=dtabCash)
            rptCash.SetParameterValue("StartDte", dtpStart.Value)
            rptCash.SetParameterValue("EndDte", dtpEnd.Value)
            crvAcctRpt.ReportSource = rptCash
#End Region
#End Region
#Region "Sales Register"
            ''''''''''''''SALES REGISTER'''''''''''''''''
        ElseIf Trim(cmbRptType.Text) = "Sales Register" Then
            dtabSales = New dsAcctRpt.SalesDataTable

            'Get invoice data
            dtabINVICT = New dsAcctRpt.INVICTDataTable
            dtabINVICT = clsAcctRpt.Get_INVICT(dtpStart.Text & " 00:00:00 AM", dtpEnd.Text & " 11:58:59 PM", cmbCompCode.Text.Trim)

            If dtabINVICT.Rows.Count > 0 Then
                Dim intInvCtr As Integer = 0

                Do While intInvCtr < dtabINVICT.Rows.Count
                    'Charges
                    Dim dblVC, dblVC1 As Double  'Vessel Charges
                    Dim dblCIM, dblCIM1 As Double 'Import Cargoes
                    Dim dblCEX, dblCEX1 As Double 'Export Cargoes
                    Dim dblST, dblST1 As Double 'Storage
                    Dim dblMC, dblMC1 As Double 'Reefer
                    Dim dblSS, dblSS1 As Double 'Stripping/Stuffing
                    Dim dblOth, dblOth1 As Double 'Other Charges
                    'Invoice Amount
                    Dim dblInvAmt As Double = 0
                    'Invoice PayDate
                    Dim strPayDte As String

                    'Initialize charges
                    dblVC = 0 : dblVC1 = 0
                    dblCIM = 0 : dblCIM1 = 0
                    dblCEX = 0 : dblCEX1 = 0
                    dblST = 0 : dblST1 = 0
                    dblMC = 0 : dblMC1 = 0
                    dblSS = 0 : dblSS1 = 0
                    dblOth = 0 : dblOth1 = 0

                    'Get the invoice amount on a per Billing Type basis------------
                    dtabINVCYB = New DataTable
                    dtabINVCYB = clsAcctRpt.Get_INVCYB(CLng(dtabINVICT.Rows(intInvCtr)("refnum").ToString))
                    If dtabINVCYB.Rows.Count > 0 Then
                        Dim intCybCtr As Integer = 0

                        Do While intCybCtr < dtabINVCYB.Rows.Count
                            Dim CCRNUM As String = dtabINVCYB.Rows(intCybCtr)("rtedsc").ToString
                            Dim CNTNUM As String = CCRNUM

                            Select Case Trim(dtabINVCYB.Rows(intCybCtr)("cyr_biltyp").ToString)
                                Case "CB" 'Cargoes
                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "EX") > 0 Then 'Export Cargoes

                                        CCRNUM = Mid(CCRNUM, 'Get CCRNUM
                                         InStr(CCRNUM, "R#") + 2, 'Start Position at 2 Characters from R
                                         InStr(CCRNUM, " ") - (InStr(CCRNUM, "R#") + 2)) 'Get Length by finding the Space ' ' and then Substracting the Starting Position

                                        CNTNUM = Mid(CNTNUM, 'Get CCRNUM
                                         InStr(CNTNUM, " ") + 1, 'Start Position at 2 Characters from Space
                                         InStr(CNTNUM, "-") - (InStr(CNTNUM, " ") + 1)) 'Get Length by finding the Dash '-' and then Substracting the Starting Position

                                        dtabInvDtl = clsAcctRpt.Get_InvDtl(CCRNUM, dtabINVCYB.Rows(intCybCtr)("rtecde"), CNTNUM)

                                        dblCEX += CDbl(dtabInvDtl.Rows(0)("ExportCharges").ToString)
                                        dblOth += CDbl(dtabInvDtl.Rows(0)("wghamt").ToString)

                                        'ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "IM") > 0 Then 'Import Cargoes
                                    Else
                                        dblCIM += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    End If
                                Case "ST" 'Storage
                                    dblST += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                Case "MC" 'Miscellaneous Charges
                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "RF") > 0 Then 'Reefer
                                        dblMC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "V") = 1 Then 'Vessel Charge
                                        dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    Else
                                        dblOth += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString) 'Equipment Rental
                                    End If
                                Case "SS" 'Stripping/Stuffing
                                    dblSS += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                Case "VB" 'Vessel Charges
                                    dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                Case Else 'Other Charges
                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "CB") = 1 Then 'Cargo Billing IMP/EXP
                                        If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "IM") > 0 Then
                                            dblCIM += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                            'ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "EX") > 0 Then
                                        Else
                                            dblCEX += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                        End If
                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "STRIP") = 1 Then 'Stripping/Stuffing
                                        dblSS += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "ST") = 1 Then 'Storage
                                        dblST += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "MC") = 1 Then 'Miscellaneous Charges
                                        If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "RF") > 0 Then 'Reefer
                                            dblMC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                        Else
                                            dblOth += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString) 'Equipment Rental
                                        End If
                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "V") = 1 Then 'Vessel Billing
                                        dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    End If
                            End Select
                            intCybCtr += 1
                        Loop
                    End If
                    '--------------------------------------------------------------

                    If Trim(dtabINVICT.Rows(intInvCtr)("Status").ToString) = "CAN" Then
                        'Add Sales data to temporary table
                        Call Add_SalesData(dtabINVICT.Rows(intInvCtr)("invdttm"),
                                           dtabINVICT.Rows(intInvCtr)("invnum"),
                                           "- - - - - C A N C E L L E D - - - - -",
                                           0, 0, 0, 0, 0, 0, 0,
                                           0, dtpEnd.MaxDate)
                    Else
                        'Check if invoice has payment record,get paydate and invoice balance
                        Dim blnSkip As Boolean = False

                        dtabPayDtl = New DataTable
                        dtabPayDtl = clsAcctRpt.Get_PayDtl(CLng(dtabINVICT.Rows(intInvCtr)("invnum").ToString), CDate(dtpEnd.Text & " 11:58:59 PM")) 'Get Payment Details
                        dblInvAmt = CDbl(dtabINVICT.Rows(intInvCtr)("InvAmt").ToString) 'Initialize Invoice Amount

                        If dtabPayDtl.Rows.Count > 0 Then 'No Payment
                            If CDbl(dtabPayDtl.Rows(0)("RBalance").ToString) <> 0 Then ' Has Balance 
                                Dim dblPayAmt As Double = 0

                                dblPayAmt = CDbl(dtabINVICT.Rows(intInvCtr)("InvAmt").ToString) - CDbl(dtabPayDtl.Rows(0)("RBalance").ToString)
                                dblInvAmt = CDbl(dtabPayDtl.Rows(0)("RBalance").ToString)

                                'Recompute charges
                                If dblVC > 0 Then 'Vessel Charges
                                    If dblVC > dblPayAmt Then
                                        dblVC1 = dblVC - dblPayAmt
                                    Else
                                        dblVC1 = dblVC
                                    End If
                                End If
                                If dblCEX > 0 Then 'Export Cargoes
                                    If (dblVC + dblCEX) > dblPayAmt Then
                                        dblCEX1 = (dblVC + dblCEX) - (dblVC1) - dblPayAmt
                                    Else
                                        dblCEX1 = dblCEX
                                    End If
                                End If
                                If dblCIM > 0 Then 'Import Cargoes
                                    If (dblVC + dblCEX + dblCIM) > dblPayAmt Then
                                        dblCIM1 = (dblVC + dblCEX + dblCIM) - (dblVC1 + dblCEX1) - dblPayAmt
                                    Else
                                        dblCIM1 = dblCIM
                                    End If
                                End If
                                If dblST > 0 Then 'Storage
                                    If (dblVC + dblCEX + dblCIM + dblST) > dblPayAmt Then
                                        dblST1 = (dblVC + dblCEX + dblCIM + dblST) - (dblVC1 + dblCEX1 + dblCIM1) - dblPayAmt
                                    Else
                                        dblST1 = dblST
                                    End If
                                End If
                                If dblMC > 0 Then 'Reefer
                                    If (dblVC + dblCEX + dblCIM + dblST + dblMC) > dblPayAmt Then
                                        dblMC1 = (dblVC + dblCEX + dblCIM + dblST + dblMC) - (dblVC1 + dblCEX1 + dblCIM1 + dblST1) - dblPayAmt
                                    Else
                                        dblMC1 = dblMC
                                    End If
                                End If
                                If dblSS > 0 Then 'Stripping/Stuffing
                                    If (dblVC + dblCEX + dblCIM + dblST + dblMC + dblSS) > dblPayAmt Then
                                        dblSS1 = (dblVC + dblCEX + dblCIM + dblST + dblMC + dblSS) - (dblVC1 + dblCEX1 + dblCIM1 + dblST1 + dblMC1) - dblPayAmt
                                    Else
                                        dblSS1 = dblSS
                                    End If
                                End If
                                If dblOth > 0 Then 'Others
                                    If (dblVC + dblCEX + dblCIM + dblST + dblSS + dblOth) > dblPayAmt Then
                                        dblOth1 = (dblVC + dblCEX + dblCIM + dblST + dblMC + dblSS + dblOth) - (dblVC1 + dblCEX1 + dblCIM1 + dblST1 + dblMC1 + dblSS1) - dblPayAmt
                                    Else
                                        dblOth1 = dblOth
                                    End If
                                End If
                            Else
                                dblVC1 = dblVC
                                dblCEX1 = dblCEX
                                dblCIM1 = dblCIM
                                dblST1 = dblST
                                dblMC1 = dblMC
                                dblSS1 = dblSS
                                dblOth1 = dblOth

                            End If

                            Call Add_SalesData(dtabINVICT.Rows(intInvCtr)("invdttm"),
                                                   dtabINVICT.Rows(intInvCtr)("invnum"),
                                                   dtabINVICT.Rows(intInvCtr)("cusnam"),
                                                   dblInvAmt, dblVC1, dblCEX1, dblCIM1, dblST1, dblMC1,
                                                   dblSS1, dblOth1, dtabPayDtl.Rows(0)("Paydate"))

                        Else
                            strPayDte = ""

                            'Add Sales data to temporary table
                            Call Add_SalesData(dtabINVICT.Rows(intInvCtr)("invdttm"),
                                               dtabINVICT.Rows(intInvCtr)("invnum"),
                                               dtabINVICT.Rows(intInvCtr)("cusnam"),
                                               dblInvAmt, dblVC, dblCEX, dblCIM, dblST, dblMC,
                                               dblSS, dblOth, dtpEnd.MaxDate)
                        End If
                    End If
                    intInvCtr += 1
                Loop
            End If

            clsAcctRpt.DisConnect()
            clsAcctRpt = Nothing

            'Call Display Report
            Dim rptSales As New rptSales

            dtabSales.DefaultView.Sort = "Invdttm Asc,Invnum Asc"
            rptSales.SetDataSource(dataTable:=dtabSales)
            rptSales.SetParameterValue("StartDte", dtpStart.Value)
            rptSales.SetParameterValue("EndDte", dtpEnd.Value)
            crvAcctRpt.ReportSource = rptSales
        Else
            Cursor = Cursors.Default
            MsgBox("Please select a valid report type!", MsgBoxStyle.Exclamation, "Display Restriction")
        End If
#End Region
        Cursor = Cursors.Default
    End Sub
    Private Function ZeroVoid(ByVal NumVar As Object) As String
        Try
            ZeroVoid = CLng(NumVar) : If ZeroVoid = "0" Then ZeroVoid = ""
        Catch ex As Exception
            ZeroVoid = ""
        End Try

    End Function

    Private Sub Add_SalesData(ByVal dteInvDte As Date, ByVal lngInvNum As Long,
                              ByVal strCustomer As String, ByVal dblInvAmt As Double,
                              ByVal dblVC As Double, ByVal dblCEX As Double,
                              ByVal dblCIM As Double, ByVal dblST As Double,
                              ByVal dblMC As Double, ByVal dblSS As Double,
                              ByVal dblOth As Double, ByVal dtePayDte As Date)

        Dim rowSales As dsAcctRpt.SalesRow

        rowSales = dtabSales.NewSalesRow

        With rowSales
            .Invdttm = dteInvDte
            .invnum = lngInvNum
            .cusnam = Trim(strCustomer)
            .invamt = dblInvAmt
            .VC = dblVC
            .CEX = dblCEX
            .CIM = dblCIM
            .ST = dblST
            .MC = dblMC
            .SS = dblSS
            .Others = dblOth
            If dtePayDte <> Date.MaxValue Then
                .PayDate = dtePayDte
            End If
        End With

        dtabSales.Rows.Add(rowSales)
    End Sub

    Private Sub Add_CashData(ByVal dtePd As Date, ByVal strDocno As String,
                             ByVal strOR As String, ByVal strChkNo As String,
                             ByVal strPayor As String, ByVal charADR As Char,
                             ByVal dblAmt As Double, ByVal dblCshAmt As Double, ByVal dblChkAmt As Double, ByVal dblAdrAmt As Double,
                             ByVal dblImp As Double, ByVal dblExp As Double,
                             ByVal dblST As Double, ByVal dblMC As Double, ByVal dblOth As Double,
                             ByVal dblSS As Double, ByVal dblAR As Double,
                             ByVal dblSV As Double, ByVal dblCompCode As String)

        Dim rowCash As dsAcctRpt.CashRow

        rowCash = dtabCash.NewCashRow

        With rowCash
            .Pddte = dtePd
            .DocNo = Trim(strDocno)
            .ORNo = Trim(strOR)
            .ChkNo = Trim(strChkNo)
            .Payor = Trim(strPayor)
            .ADR = charADR
            .Amt = dblAmt
            .CshAMT = dblCshAmt
            .ChkAmt = dblChkAmt
            .AdrAmt = dblAdrAmt
            .ImpAmt = dblImp
            .ExpAmt = dblExp
            .ST = dblST
            .McAmt = dblMC
            .OthAmt = dblOth
            .SS = dblSS
            .AR = dblAR
            .SV = dblSV
            .CompanyCode = dblCompCode
        End With

        dtabCash.Rows.Add(rowCash)
    End Sub

    Private Sub frmAcctRpt_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text &= " v" & Application.ProductVersion
    End Sub
End Class
