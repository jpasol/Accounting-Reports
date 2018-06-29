Imports System.Data.SqlClient

Public Class clsAcctRpt

    'Objects
    Private objConn As SqlConnection

    'Variables
    Private strConn As String

    Public Sub New(ByVal strSrvr As String, ByVal strDB As String)
        strConn = "Data Source='" & Trim(strSrvr) & "';Initial Catalog='" & Trim(strDB) & "';UID='SA_ICTSI';PWD=Ictsi123"
        'strConn = "Data Source='" & Trim(strSrvr) & "';Initial Catalog='" & Trim(strDB) & "';Integrated Security=SSPI"
    End Sub

    Private Function Connection() As Boolean
        Try
            If objConn Is Nothing Then
                objConn = New SqlConnection
                objConn.ConnectionString = Trim(strConn)
                objConn.Open()
            End If

            If objConn.State = ConnectionState.Closed Then
                objConn.Open()
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Connection Error")
            Return False
        End Try
    End Function

    Public Function Get_INVICT(ByVal dteStart As Date, ByVal dteEnd As Date, ByVal compCode As String) As dsAcctRpt.INVICTDataTable
        Dim cmdINVICT As New SqlCommand
        Dim daINVICT As New SqlDataAdapter
        Dim dtabINVICT As New dsAcctRpt.INVICTDataTable
        Dim sqlQuery As String = "Select a.* From INVICT a " & _
        "Inner join invcyb b on a.refnum = b.refnum " & _
        "Where a.invdttm >='" & dteStart & "' And a.invdttm <='" & dteEnd & "' "

        If compCode <> "ALL" Then
            sqlQuery &= "and b.CompanyCode = '" & compCode & "' "
        End If

        sqlQuery &= "Order By a.refnum Asc"

        If Connection() = True Then
            With cmdINVICT
                .Connection = objConn
                .CommandText = sqlQuery
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daINVICT.SelectCommand = cmdINVICT
                daINVICT.Fill(dtabINVICT)

                Return dtabINVICT
            End With
        End If
    End Function

    Public Function Get_InvRefNum(ByVal lngInvNum As Long) As Long
        Dim cmdINVICT As New SqlCommand
        Dim lngRefNum As Long = 0

        If Connection() = True Then
            With cmdINVICT
                .Connection = objConn
                .CommandText = "Select refnum From INVICT Where invnum=" & lngInvNum
                .CommandType = CommandType.Text

                lngRefNum = .ExecuteScalar

                Return lngRefNum
            End With
        End If
    End Function

    Public Function Get_INVCYB(ByVal lngRefNum As Long) As DataTable
        Dim cmdINVCYB As New SqlCommand
        Dim daINVCYB As New SqlDataAdapter
        Dim dtabINVCYB As New DataTable

        If Connection() = True Then
            With cmdINVCYB
                .Connection = objConn
                .CommandText = "SELECT * FROM dbo.INVCYB INNER JOIN dbo.CYRate ON dbo.INVCYB.rtecde = dbo.CYRate.cyr_rtecde Where dbo.INVCYB.refnum=" & lngRefNum
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daINVCYB.SelectCommand = cmdINVCYB
                daINVCYB.Fill(dtabINVCYB)

                Return dtabINVCYB
            End With
        End If
    End Function

    Public Function Get_PayDtl(ByVal lngInvNum As Long, ByVal dtePay As Date) As DataTable
        Dim cmdPayDtl As New SqlCommand
        Dim daPayDtl As New SqlDataAdapter
        Dim dtabPayDtl As New DataTable

        If Connection() = True Then
            With cmdPayDtl
                .Connection = objConn
                .CommandText = "SELECT TOP 1* FROM InvPayDtl WHERE invnum=" & lngInvNum & " AND paydate <='" & dtePay & "' ORDER BY PAYdate DESC"
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daPayDtl.SelectCommand = cmdPayDtl
                daPayDtl.Fill(dtabPayDtl)

                Return dtabPayDtl
            End With
        End If
    End Function

    Public Function Get_CCRPay(ByVal dteStart As Date, ByVal dteEnd As Date, ByVal compCode As String) As DataTable
        Dim cmdCCRPay As New SqlCommand
        Dim daCCRPay As New SqlDataAdapter
        Dim dtabCCRPay As New DataTable

        Dim sqlQuery As String = "SELECT a.* FROM CCRPay a " & _
            "INNER JOIN CCRCyx b on a.refnum=b.refnum " & _
            "INNER JOIN CCRDtl c on a.refnum=c.refnum " & _
            "WHERE a.sysdttm >='" & dteStart & "' And a.sysdttm <='" & dteEnd & "' "

        If compCode <> "ALL" Then
            sqlQuery &= " AND (b.CompanyCode='" & compCode & "' or c.CompanyCode='" & compCode & "')"
        End If
        sqlQuery &= " Order By a.Refnum"


        If Connection() = True Then
            With cmdCCRPay
                .Connection = objConn
                .CommandText = sqlQuery
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCCRPay.SelectCommand = cmdCCRPay
                daCCRPay.Fill(dtabCCRPay)

                Return dtabCCRPay
            End With
        End If
    End Function

    Public Function Get_CYMPay(ByVal dteStart As Date, ByVal dteEnd As Date, ByVal compCode As String) As DataTable
        Dim cmdCYMPay As New SqlCommand
        Dim daCYMPay As New SqlDataAdapter
        Dim dtabCYMPay As New DataTable
        Dim sqlQuery As String = "SELECT a.* FROM CYMPay a " & _
            "INNER JOIN CYMGps b on a.refnum=b.refnum " & _
            "WHERE a.sysdttm >='" & dteStart & "' And sysdttm <='" & dteEnd & "' "

        If compCode <> "ALL" Then
            sqlQuery &= " AND b.CompanyCode='" & compCode & "'"
        End If
        sqlQuery &= " Order By a.Refnum"

        If Connection() = True Then
            With cmdCYMPay
                .Connection = objConn
                .CommandText = sqlQuery
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCYMPay.SelectCommand = cmdCYMPay
                daCYMPay.Fill(dtabCYMPay)

                Return dtabCYMPay
            End With
        End If
    End Function

    Public Function Get_InvPayHdr(ByVal dteStart As Date, ByVal dteEnd As Date) As DataTable
        Dim cmdInvPayHdr As New SqlCommand
        Dim daInvPayHdr As New SqlDataAdapter
        Dim dtabInvPayHdr As New DataTable

        If Connection() = True Then
            With cmdInvPayHdr
                .Connection = objConn
                .CommandText = "SELECT * FROM INVPayHdr WHERE ORDate >='" & dteStart & "' And ORDate <='" & dteEnd & "'"
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daInvPayHdr.SelectCommand = cmdInvPayHdr
                daInvPayHdr.Fill(dtabInvPayHdr)

                Return dtabInvPayHdr
            End With
        End If
    End Function

    Public Function Chk_CAN_UG(ByVal intCCRTyp As Integer, ByVal lngRefNum As Long) As Boolean
        Dim cmdCanUg As New SqlCommand
        Dim intCount As Integer = 0
        Dim daCanUg As New SqlDataAdapter
        Dim dtabCanUg As New DataTable

        If Connection() = True Then
            With cmdCanUg
                .Connection = objConn
                If intCCRTyp = 1 Then
                    'Export
                    .CommandText = "SELECT * FROM CCRCyx WHERE (status = 'CAN' Or guarntycde = 'Y') And refnum = " & lngRefNum
                ElseIf intCCRTyp = 2 Then
                    'Special Services
                    .CommandText = "SELECT * FROM CCRDtl WHERE (status = 'CAN' Or guarntycde = 'Y') And refnum = " & lngRefNum
                ElseIf intCCRTyp = 3 Then
                    'Import
                    .CommandText = "SELECT * FROM CYMGps WHERE (status = 'CAN' Or gtycde <> ' ') And refnum = " & lngRefNum
                Else
                    'Invoice
                    .CommandText = "SELECT * FROM InvIct WHERE (status = 'CAN') And invnum = " & lngRefNum
                End If
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCanUg.SelectCommand = cmdCanUg
                daCanUg.Fill(dtabCanUg)

                intCount = dtabCanUg.Rows.Count

                If intCount > 0 Then
                    Return True
                Else
                    Return False
                End If
            End With
        End If
    End Function

    Public Function Get_CCRCyx(ByVal lngRefNum As Long) As DataTable
        Dim cmdCCRCyx As New SqlCommand
        Dim daCCRCyx As New SqlDataAdapter
        Dim dtabCCRCyx As New DataTable

        If Connection() = True Then
            With cmdCCRCyx
                .Connection = objConn
                .CommandText = "SELECT * FROM CCRCyx WHERE refnum =" & lngRefNum
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCCRCyx.SelectCommand = cmdCCRCyx
                daCCRCyx.Fill(dtabCCRCyx)

                Return dtabCCRCyx
            End With
        End If
    End Function

    Public Function Get_CCRDtl(ByVal lngRefNum As Long) As DataTable
        Dim cmdCCRDtl As New SqlCommand
        Dim daCCRDtl As New SqlDataAdapter
        Dim dtabCCRDtl As New DataTable

        If Connection() = True Then
            With cmdCCRDtl
                .Connection = objConn
                .CommandText = "SELECT * FROM CCRDtl WHERE refnum =" & lngRefNum
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCCRDtl.SelectCommand = cmdCCRDtl
                daCCRDtl.Fill(dtabCCRDtl)

                Return dtabCCRDtl
            End With
        End If
    End Function

    Public Function Get_CYMGps(ByVal lngRefNum As Long) As DataTable
        Dim cmdCYMGps As New SqlCommand
        Dim daCYMGps As New SqlDataAdapter
        Dim dtabCYMGps As New DataTable

        If Connection() = True Then
            With cmdCYMGps
                .Connection = objConn
                .CommandText = "SELECT * FROM CYMGps WHERE refnum =" & lngRefNum
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCYMGps.SelectCommand = cmdCYMGps
                daCYMGps.Fill(dtabCYMGps)

                Return dtabCYMGps
            End With
        End If
    End Function

    Public Function Get_InvPayDtl(ByVal lngORNum As Long) As DataTable
        Dim cmdInvPayDtl As New SqlCommand
        Dim daInvPayDtl As New SqlDataAdapter
        Dim dtabInvPayDtl As New DataTable

        If Connection() = True Then
            With cmdInvPayDtl
                .Connection = objConn
                .CommandText = "SELECT * FROM InvPayDtl WHERE ornum =" & lngORNum
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daInvPayDtl.SelectCommand = cmdInvPayDtl
                daInvPayDtl.Fill(dtabInvPayDtl)

                Return dtabInvPayDtl
            End With
        End If
    End Function

    Public Function Get_CustomerName(ByVal lngCusCde As Long) As String
        Dim cmdCusName As New SqlCommand
        Dim strCusNam As String = ""

        If Connection() = True Then
            With cmdCusName
                .Connection = objConn
                .CommandText = "SELECT cusnam FROM Customer WHERE cuscde =" & lngCusCde
                .CommandType = CommandType.Text

                strCusNam = .ExecuteScalar

                Return Trim(strCusNam)
            End With
        End If
    End Function

    Public Function Is_Invoice_AR(ByVal lngInvnum As Long, ByVal dteStart As Date, ByVal dteEnd As Date) As Boolean
        Dim cmdInvICT As New SqlCommand
        Dim dteInvIssuance As Date

        If Connection() = True Then
            With cmdInvICT
                .Connection = objConn
                .CommandText = "SELECT invdttm FROM InvICT WHERE invnum =" & lngInvnum
                .CommandType = CommandType.Text

                dteInvIssuance = .ExecuteScalar
            End With

            If dteInvIssuance >= dteStart And dteInvIssuance <= dteEnd Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Public Function DisConnect()
        If Not objConn Is Nothing Then
            objConn.Close()
            objConn = Nothing
        End If
    End Function
End Class
