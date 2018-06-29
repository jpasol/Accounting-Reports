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
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAcctRpt))
        Me.dtpStart = New System.Windows.Forms.DateTimePicker
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker
        Me.crvAcctRpt = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.cmbRptType = New System.Windows.Forms.ComboBox
        Me.btnPreview = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmbCompCode = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'dtpStart
        '
        Me.dtpStart.CalendarTitleBackColor = System.Drawing.Color.SlateGray
        Me.dtpStart.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpStart.Location = New System.Drawing.Point(16, 80)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(144, 23)
        Me.dtpStart.TabIndex = 1
        '
        'dtpEnd
        '
        Me.dtpEnd.CalendarTitleBackColor = System.Drawing.Color.SlateGray
        Me.dtpEnd.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Short
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
        Me.crvAcctRpt.DisplayGroupTree = False
        Me.crvAcctRpt.Location = New System.Drawing.Point(176, 8)
        Me.crvAcctRpt.Name = "crvAcctRpt"
        Me.crvAcctRpt.ReportSource = Nothing
        Me.crvAcctRpt.Size = New System.Drawing.Size(832, 696)
        Me.crvAcctRpt.TabIndex = 4
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
    Private dtabCash As dsAcctRpt.CashDataTable

    'Configuration Settings
    Private strServer As String = CType(AppSettings("Server"), String)
    Private strDataBase As String = CType(AppSettings("Database"), String)

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Dim clsAcctRpt As New clsAcctRpt(strServer, strDataBase)
        Dim strDocNo As String = ""
        Dim strChkNo As String = ""

        Cursor = Cursors.WaitCursor

        If Trim(cmbRptType.Text) = "Cash Receipt" Then
            Dim lngExpAmt As Double = 0
            Dim lngSpcAmt As Double = 0
            
            dtabCash = New dsAcctRpt.CashDataTable

            'Export and Special Services
            dtabCCRPay = New DataTable
            dtabCCRPay = clsAcctRpt.Get_CCRPay(dtpStart.Text & " 00:00:00 AM", dtpEnd.Text & " 11:58:59 PM", cmbCompCode.Text.Trim)

            If dtabCCRPay.Rows.Count > 0 Then
                Dim lngCCRPay As Long = 0

                Do While lngCCRPay < dtabCCRPay.Rows.Count
                    If clsAcctRpt.Chk_CAN_UG(dtabCCRPay.Rows(lngCCRPay)("ccrtyp"), dtabCCRPay.Rows(lngCCRPay)("refnum")) = False Then
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
                                Do While intCCRCyx < dtabCCRCyx.Rows.Count
                                    lngExpAmt = lngExpAmt + _
                                                dtabCCRCyx.Rows(intCCRCyx)("whfamt") + _
                                                dtabCCRCyx.Rows(intCCRCyx)("arramt") + _
                                                dtabCCRCyx.Rows(intCCRCyx)("ovzamt") + _
                                                dtabCCRCyx.Rows(intCCRCyx)("dgramt") + _
                                                dtabCCRCyx.Rows(intCCRCyx)("arrvat") - _
                                                dtabCCRCyx.Rows(intCCRCyx)("arrtax")
                                    intCCRCyx += 1
                                Loop
                            End If


                            'Add cash data

                            'PRNH 08102017 - Retrieve Exporter if Cusnam is blank

                            Dim cusNameExp As String = Trim(dtabCCRPay.Rows(lngCCRPay)("cusnam"))
                            If cusNameExp = "" Then
                                cusNameExp = dtabCCRCyx.Rows(0)("exprtr")
                            End If

                            'Add_CashData(dtabCCRPay.Rows(lngCCRPay)("sysdttm"), _
                            '             strDocNo, "", Trim(strChkNo), _
                            '             dtabCCRPay.Rows(lngCCRPay)("cusnam"), _
                            '             lngExpAmt, 0, lngExpAmt, 0, 0, 0, 0)
                            Add_CashData(dtabCCRPay.Rows(lngCCRPay)("sysdttm"), _
                                        strDocNo, "", Trim(strChkNo), _
                                        cusNameExp, _
                                        lngExpAmt, 0, lngExpAmt, 0, 0, 0, 0, dtabCCRCyx.Rows(0)("CompanyCode"))


                        Else 'Special Services
                            Dim intCCRCys As Integer = 0

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
                                Do While intCCRCys < dtabCCRDtl.Rows.Count
                                    lngSpcAmt = lngSpcAmt + _
                                                dtabCCRDtl.Rows(intCCRCys)("amt") + _
                                                dtabCCRDtl.Rows(intCCRCys)("vatamt") + _
                                                dtabCCRDtl.Rows(intCCRCys)("ovzamt") + _
                                                dtabCCRDtl.Rows(intCCRCys)("dgramt") - _
                                                dtabCCRDtl.Rows(intCCRCys)("wtax")
                                    intCCRCys += 1
                                Loop
                            End If

                            'Add cash data
                            Add_CashData(dtabCCRPay.Rows(lngCCRPay)("sysdttm"), _
                                         strDocNo, "", Trim(strChkNo), _
                                         dtabCCRPay.Rows(lngCCRPay)("cusnam"), lngSpcAmt, 0, 0, lngSpcAmt, 0, 0, 0, dtabCCRDtl.Rows(0)("CompanyCode"))
                        End If
                    End If
                    lngCCRPay += 1
                Loop
            End If

            'Import
            dtabCYMPay = New DataTable
            dtabCYMPay = clsAcctRpt.Get_CYMPay(dtpStart.Text & " 00:00:00 AM", dtpEnd.Text & " 11:58:59 PM", cmbCompCode.Text.Trim)

            If dtabCYMPay.Rows.Count > 0 Then
                Dim lngCYMAmt As Double = 0
                Dim lngCYMPay As Long = 0

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
                            Do While intCCRCym < dtabCYMGps.Rows.Count
                                lngCYMAmt = lngCYMAmt + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("udstoamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstoamt")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("udstovat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstovat")) - _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("udstotax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("udstotax")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("stoamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stoamt")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("arramt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arramt")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("whfamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("whfamt")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("wghamt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghamt")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("rframt") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rframt")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("stovat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stovat")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("arrvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arrvat")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("wghvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghvat")) + _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("rfrvat") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rfrvat")) - _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("stotax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("stotax")) - _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("arrtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("arrtax")) - _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("wghtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("wghtax")) - _
                                        IIf(dtabCYMGps.Rows(intCCRCym)("rfrtax") Is DBNull.Value, 0, dtabCYMGps.Rows(intCCRCym)("rfrtax"))
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

                            'Add cash data

                            'PRNH 08102017 - Retrieve Consignee if Cusnam is blank
                            Dim custImp As String = dtabCYMPay.Rows(lngCYMPay)("cusnam")
                            If custImp = "" Then
                                custImp = dtabCYMGps.Rows(0)("cnsgne")
                            End If

                            'Add_CashData(dtabCYMPay.Rows(lngCYMPay)("sysdttm"), _
                            '             strDocNo, "", Trim(strChkNo), _
                            '             dtabCYMPay.Rows(lngCYMPay)("cusnam"), lngCYMAmt, lngCYMAmt, 0, 0, 0, 0, 0)

                            Add_CashData(dtabCYMPay.Rows(lngCYMPay)("sysdttm"), _
                                         strDocNo, "", Trim(strChkNo), _
                                         custImp, lngCYMAmt, lngCYMAmt, 0, 0, 0, 0, 0, dtabCYMGps.Rows(0)("CompanyCode"))
                        End If
                    End If
                    lngCYMPay += 1
                Loop
            End If

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
                                    Add_CashData(dtabInvPayHdr.Rows(rowInvPayHdr)("ORDate"), _
                                                 Trim("INV " & strDocNo), dtabInvPayHdr.Rows(rowInvPayHdr)("ORNum"), Trim(strChkNo), _
                                                 strCusName, lngInvAmt, 0, 0, 0, 0, lngInvAmt, 0, dtabInvPayDtl.Rows(intInvPayDtl)("CompanyCode"))
                                Else
                                    'Get invoice data
                                    Dim lngInvRefNo As Long = 0
                                    lngInvRefNo = clsAcctRpt.Get_InvRefNum(strDocNo)

                                    'Charges
                                    Dim dblVC, dblVC1 As Double  'Vessel Charges
                                    Dim dblCIM, dblCIM1 As Double 'Import Cargoes
                                    Dim dblCEX, dblCEX1 As Double 'Export Cargoes
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
                                    dblSS = 0 : dblSS1 = 0
                                    dblOth = 0 : dblOth1 = 0

                                    'Get the invoice amount on a per Billing Type basis------------
                                    dtabINVCYB = New DataTable
                                    dtabINVCYB = clsAcctRpt.Get_INVCYB(lngInvRefNo)
                                    If dtabINVCYB.Rows.Count > 0 Then
                                        Dim intCybCtr As Integer = 0

                                        Do While intCybCtr < dtabINVCYB.Rows.Count
                                            Select Case Trim(dtabINVCYB.Rows(intCybCtr)("cyr_biltyp").ToString)
                                                Case "VB" 'Vessel Charges
                                                    dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                Case "CB" 'Cargoes
                                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "EXP") > 0 Then
                                                        'Export Cargoes
                                                        dblCEX += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "IMP") > 0 Then
                                                        'Import Cargoes
                                                        dblCIM += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                    Else
                                                        'Do nothing
                                                    End If
                                                Case "SS" 'Stripping/Stuffing
                                                    dblSS += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                                Case Else 'Other Charges
                                                    dblOth += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
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
                                    If dblSS > 0 Then 'Stripping/Stuffing
                                        If (dblVC + dblCEX + dblCIM + dblSS) >= dblPayAmt Then
                                            dblSS1 = dblPayAmt - (dblVC1 + dblCEX1 + dblCIM1)
                                        Else
                                            dblSS1 = dblSS
                                        End If
                                    End If
                                    If dblOth > 0 Then 'Others
                                        If (dblVC + dblCEX + dblCIM + dblSS + dblOth) >= dblPayAmt Then
                                            dblOth1 = dblPayAmt - (dblVC1 + dblCEX1 + dblCIM1 + dblSS1)
                                        Else
                                            dblOth1 = dblOth
                                        End If
                                    End If

                                    'Add cash data
                                    Add_CashData(dtabInvPayHdr.Rows(rowInvPayHdr)("ORDate"), _
                                                 Trim("INV " & strDocNo), dtabInvPayHdr.Rows(rowInvPayHdr)("ORNum"), Trim(strChkNo), _
                                                 strCusName, lngInvAmt, dblCIM1, dblCEX1, dblOth1, dblSS1, 0, dblVC1, _
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

            rptCash.SetDataSource(dtabCash)
            rptCash.SetParameterValue("StartDte", dtpStart.Value)
            rptCash.SetParameterValue("EndDte", dtpEnd.Value)
            crvAcctRpt.ReportSource = rptCash
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
                    dblSS = 0 : dblSS1 = 0
                    dblOth = 0 : dblOth1 = 0

                    'Get the invoice amount on a per Billing Type basis------------
                    dtabINVCYB = New DataTable
                    dtabINVCYB = clsAcctRpt.Get_INVCYB(CLng(dtabINVICT.Rows(intInvCtr)("refnum").ToString))
                    If dtabINVCYB.Rows.Count > 0 Then
                        Dim intCybCtr As Integer = 0

                        Do While intCybCtr < dtabINVCYB.Rows.Count
                            Select Case Trim(dtabINVCYB.Rows(intCybCtr)("cyr_biltyp").ToString)
                                Case "VB" 'Vessel Charges
                                    dblVC += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                Case "CB" 'Cargoes
                                    If InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "EXP") > 0 Then
                                        'Export Cargoes
                                        dblCEX += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    ElseIf InStr(Trim(dtabINVCYB.Rows(intCybCtr)("rtecde").ToString), "IMP") > 0 Then
                                        'Import Cargoes
                                        dblCIM += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                    Else
                                        'Do nothing
                                    End If
                                Case "SS" 'Stripping/Stuffing
                                    dblSS += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                                Case Else 'Other Charges
                                    dblOth += CDbl(dtabINVCYB.Rows(intCybCtr)("invamt").ToString)
                            End Select
                            intCybCtr += 1
                        Loop
                    End If
                    '--------------------------------------------------------------

                    If Trim(dtabINVICT.Rows(intInvCtr)("Status").ToString) = "CAN" Then
                        'Add Sales data to temporary table
                        Call Add_SalesData(dtabINVICT.Rows(intInvCtr)("invdttm"), _
                                           dtabINVICT.Rows(intInvCtr)("invnum"), _
                                           "- - - - - C A N C E L L E D - - - - -", _
                                           0, 0, 0, 0, 0, _
                                           0, dtpEnd.MaxDate, dtabINVCYB.Rows(0)("invamt"))
                    Else
                        'Check if invoice has payment record,get paydate and invoice balance
                        Dim blnSkip As Boolean = False

                        dtabPayDtl = New DataTable
                        dtabPayDtl = clsAcctRpt.Get_PayDtl(CLng(dtabINVICT.Rows(intInvCtr)("invnum").ToString), CDate(dtpEnd.Text & " 11:58:59 PM"))
                        If dtabPayDtl.Rows.Count > 0 Then
                            If CDbl(dtabPayDtl.Rows(0)("RBalance").ToString) = 0 Then
                                blnSkip = True
                            Else
                                Dim dblPayAmt As Double = 0

                                dblPayAmt = CDbl(dtabINVICT.Rows(intInvCtr)("InvAmt").ToString) - CDbl(dtabPayDtl.Rows(0)("RBalance").ToString)
                                dblInvAmt = CDbl(dtabPayDtl.Rows(0)("RBalance").ToString)

                                'Recompute charges
                                If dblVC > 0 Then 'Vessel Charges
                                    If dblVC > dblPayAmt Then
                                        dblVC1 = dblVC - dblPayAmt
                                    End If
                                End If
                                If dblCEX > 0 Then 'Export Cargoes
                                    If (dblVC + dblCEX) > dblPayAmt Then
                                        dblCEX1 = (dblVC + dblCEX) - (dblVC1) - dblPayAmt
                                    End If
                                End If
                                If dblCIM > 0 Then 'Import Cargoes
                                    If (dblVC + dblCEX + dblCIM) > dblPayAmt Then
                                        dblCIM1 = (dblVC + dblCEX + dblCIM) - (dblVC1 + dblCEX1) - dblPayAmt
                                    End If
                                End If
                                If dblSS > 0 Then 'Stripping/Stuffing
                                    If (dblVC + dblCEX + dblCIM + dblSS) > dblPayAmt Then
                                        dblSS1 = (dblVC + dblCEX + dblCIM + dblSS) - (dblVC1 + dblCEX1 + dblCIM1) - dblPayAmt
                                    End If
                                End If
                                If dblOth > 0 Then 'Others
                                    If (dblVC + dblCEX + dblCIM + dblSS + dblOth) > dblPayAmt Then
                                        dblOth1 = (dblVC + dblCEX + dblCIM + dblSS + dblOth) - (dblVC1 + dblCEX1 + dblCIM1 + dblSS1) - dblPayAmt
                                    End If
                                End If
                            End If

                            If blnSkip = False Then
                                'Add Sales data to temporary table
                                Call Add_SalesData(dtabINVICT.Rows(intInvCtr)("invdttm"), _
                                                   dtabINVICT.Rows(intInvCtr)("invnum"), _
                                                   dtabINVICT.Rows(intInvCtr)("cusnam"), _
                                                   dblInvAmt, dblVC1, dblCEX1, dblCIM1, dblSS1, _
                                                   dblOth1, dtabPayDtl.Rows(0)("Paydate"), _
                                                   "")

                            End If
                        Else
                            dblInvAmt = CDbl(dtabINVICT.Rows(intInvCtr)("InvAmt").ToString)
                            strPayDte = ""

                            'Add Sales data to temporary table
                            Call Add_SalesData(dtabINVICT.Rows(intInvCtr)("invdttm"), _
                                               dtabINVICT.Rows(intInvCtr)("invnum"), _
                                               dtabINVICT.Rows(intInvCtr)("cusnam"), _
                                               dblInvAmt, dblVC, dblCEX, dblCIM, dblSS, _
                                               dblOth, dtpEnd.MaxDate, "")
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
            rptSales.SetDataSource(dtabSales)
            rptSales.SetParameterValue("StartDte", dtpStart.Value)
            rptSales.SetParameterValue("EndDte", dtpEnd.Value)
            crvAcctRpt.ReportSource = rptSales
        Else
            Cursor = Cursors.Default
            MsgBox("Please select a valid report type!", MsgBoxStyle.Exclamation, "Display Restriction")
        End If

        Cursor = Cursors.Default
    End Sub

    Private Sub Add_SalesData(ByVal dteInvDte As Date, ByVal lngInvNum As Long, _
                              ByVal strCustomer As String, ByVal dblInvAmt As Double, _
                              ByVal dblVC As Double, ByVal dblCEX As Double, _
                              ByVal dblCIM As Double, ByVal dblSS As Double, _
                              ByVal dblOth As Double, ByVal dtePayDte As Date, _
                              ByVal dblCompCode As String)

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
            .SS = dblSS
            .Others = dblOth
            If dtePayDte <> Date.MaxValue Then
                .PayDate = dtePayDte
            End If
        End With

        dtabSales.Rows.Add(rowSales)
    End Sub

    Private Sub Add_CashData(ByVal dtePd As Date, ByVal strDocno As String, _
                             ByVal strOR As String, ByVal strChkNo As String, _
                             ByVal strPayor As String, ByVal dblAmt As Double, _
                             ByVal dblImp As Double, ByVal dblExp As Double, _
                             ByVal dblMC As Double, ByVal dblSS As Double, _
                             ByVal dblAR As Double, ByVal dblSV As Double, _
                             ByVal dblCompCode As String)

        Dim rowCash As dsAcctRpt.CashRow

        rowCash = dtabCash.NewCashRow

        With rowCash
            .Pddte = dtePd
            .DocNo = Trim(strDocno)
            .ORNo = Trim(strOR)
            .ChkNo = Trim(strChkNo)
            .Payor = Trim(strPayor)
            .Amt = dblAmt
            .ImpAmt = dblImp
            .ExpAmt = dblExp
            .McAmt = dblMC
            .SS = dblSS
            .AR = dblAR
            .SV = dblSV
            .CompanyCode = dblCompCode
        End With

        dtabCash.Rows.Add(rowCash)
    End Sub
End Class
