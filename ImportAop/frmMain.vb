Imports System.Data.Odbc
Imports Microsoft.SqlServer
Public Class frmMain
    Dim mintWaitTime As Integer
    Private mblnQuit As Boolean

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Show()
        pobjSqlRas.ConnectionString = pstrConnectionRAS
        pobjSqlAop.ConnectionString = pstrConnectionAop
        If Not pobjSqlRas.Connect() Then
            MsgBox("Unable to connect SQL for RAS")
            Me.Dispose()
            Exit Sub
        End If
        Try
            pobjOdbc.Connection.ConnectionString = "DSN=Travel2020;OLE DB Services=-2; Connection Timeout=30;"
            'pobjOdbc.Connection.ConnectionString = "DSN=Travel2021;OLE DB Services=-2; Connection Timeout=30;"
            If pobjOdbc.Connection.State <> ConnectionState.Open Then
                pobjOdbc.Connection.Open()
            End If

        Catch ex As Exception
            MsgBox("Unable to connect ODBC for AOP" & vbNewLine & ex.Message)
            Me.Dispose()
            Exit Sub
        End Try

        ProcessQ("All", "ODBC")

        'If Not pobjSqlAop.Connect() Then
        '    MsgBox("Unable to connect SQL for AOP")
        '    Me.Dispose()
        '    Exit Sub
        'End If


    End Sub
    Private Sub barViewPendingQ_Click(sender As Object, e As EventArgs) Handles barViewPendingQ.Click
        frmViewPending.ShowDialog()
    End Sub

    Private Sub barRunAir_Click(sender As Object, e As EventArgs) Handles barRunAir.Click
        ProcessQ("Air", "SQL")
    End Sub

    Private Sub barRunNonAir_Click(sender As Object, e As EventArgs) Handles barRunNonAir.Click
        ProcessQ("NonAir", "SQL")
    End Sub

    Private Sub barRunAll_Click(sender As Object, e As EventArgs) Handles barRunAll.Click
        ProcessQ("All", "SQL")
    End Sub

    Private Function ProcessQ(strProd As String, strQuerryType As String) As Boolean
        Dim tblQueues As DataTable
        Dim dteStart As Date
        Dim intElapsed As Integer
        Dim dteEndTime As Date
        Dim strFilterLinkId As String
        Dim mAOPListID, mSQL, mQuerry As String  '^_^20221125 add by 7643
        Dim mReturn As New DataTable  '^_^20221125 add by 7643
        mblnQuit = False
        TxtFeedBack.Text = Now & " Started " & strProd
        Me.Refresh()
        Dim strQuerry As String

        strFilterLinkId = " and linkid=(select top 1 linkId from aopqueue where status='ok' and QuerryType='" & strQuerryType & "'"

        Select Case strProd
            Case "Air", "NonAir", "AOP"
                strFilterLinkId = strFilterLinkId & " and Prod='" & strProd & "' order by recid)"
            Case "All"
                strFilterLinkId = strFilterLinkId & " order by recid)"

            Case Else
                MsgBox("Invalid product:" & strProd)
                Return False
        End Select
        'strFilterLinkId = " and LinkId between 54484 and 54484"          'test
        strQuerry = "select * from AopQueue" _
                        & " where Status='OK' and QuerryType='" & strQuerryType & "'" & strFilterLinkId & " order by recid"

StartPosition:
        Try
            If pobjSqlRas.Connection.State <> ConnectionState.Open Then
                pobjSqlRas.Connect()
            End If
            tblQueues = pobjSqlRas.GetDataTable(strQuerry)
            If tblQueues.Rows.Count = 0 Then
                mintWaitTime = 60
            Else
                mintWaitTime = 1
                dteStart = Now
                Select Case strQuerryType
                    Case "ODBC"
                        If pobjOdbc.Connection.State <> ConnectionState.Open Then
                            pobjOdbc.Connection.Open()
                        End If

                        For i = 0 To tblQueues.Rows.Count - 1
                            Dim objRow As DataRow = tblQueues.Rows(i)
                            Select Case objRow("Prod")
                                Case "AOP"
                                    If objRow("B_I") = "B" Then
                                        DowloadBspBillData(objRow("Querry"), objRow("Counter"))
                                    ElseIf objRow("B_I") = "VC" Then
                                        DowloadBspVendorCreditData(objRow("Querry"), objRow("Counter"))
                                    End If
                                    intElapsed = DateDiff(DateInterval.Second, dteStart, Now)
                                    If pobjSqlRas.ExecuteNonQuerry("update AopQueue set Status='RR',LstUpdate=getdate()" _
                                                                   & ",Count=Count+1 where Status='OK' and RecId=" & objRow("RecId")) Then
                                        TxtFeedBack.Text = Now & " " & objRow("RecId") & " " & intElapsed _
                                                                    & " " & objRow("B_I") & " " & objRow("TrxCode") & " imported!" _
                                                                    & vbCrLf & TxtFeedBack.Text
                                    Else
                                        TxtFeedBack.Text = Now & " " & objRow("B_I") & " " & objRow("RecId") _
                                            & " downloaed, BUT can not change Status in RAS!"
                                    End If

                                Case Else
                                    If pobjOdbc.ExecuteNonQuerry(objRow("Querry")) Then
                                        If i = tblQueues.Rows.Count - 1 Then

                                            intElapsed = DateDiff(DateInterval.Second, dteStart, Now)
                                            If pobjSqlRas.ExecuteNonQuerry("update AopQueue set Status='RR',LstUpdate=getdate(),Count=Count+1 where Status='OK' and LinkId=" & objRow("LinkId")) Then
                                                TxtFeedBack.Text = Now & " " & objRow("LinkId") & " " & intElapsed & " " & objRow("B_I") & " " & objRow("TrxCode") & " imported!" & vbCrLf & TxtFeedBack.Text
                                            Else
                                                TxtFeedBack.Text = Now & " " & objRow("B_I") & " " & objRow("TrxCode") & " imported, BUT can not change Status in RAS!"
                                            End If
                                        End If
                                    Else
                                        If pobjOdbc.UpdtErr.Contains("Error parsing complete XML return string") _
                                            Or pobjOdbc.UpdtErr.Contains("Incorrectly built XML from Update Start") Then  '^_^20221125 add by 7643
                                            Application.Restart()
                                            Environment.Exit(0)
                                        End If

                                        If (objRow("B_I") = "P" AndAlso pobjOdbc.UpdtErr.Contains("specified in the request cannot be found")) _
                                            OrElse pobjOdbc.UpdtErr.Contains("before the closing date of the company") Then
                                            pobjSqlRas.ExecuteNonQuerry("update AopQueue set Status='ER',LstUpdate=getdate(),Count=Count+1 where Status='OK' and LinkId=" & objRow("LinkId"))
                                            mintWaitTime = 1
                                            '^_^20230224 add by 7643 -b-
                                        ElseIf objRow("Prod") = "UNC" Then
                                            If pobjOdbc.UpdtErr.Contains("There is an invalid reference to QuickBooks Item") Then
                                                mSQL = "insert into [42.117.5.70].AOP.dbo.tblSync_TourCode (TourCode,City) " &
                                                       "values ('" & objRow("Memo") & "','" & pstrCity & "')"
                                                pobjSqlRas.ExecuteNonQuerry(mSQL)

                                                mintWaitTime = 60
                                            ElseIf pobjOdbc.UpdtErr.Contains("The given object ID """" in the field ""list id"" is invalid") Then
                                                mSQL = "select ven.AOPListID " &
                                                 "from Vendor ven left join UNC_Payments unc on ven.RecID=unc.PayeeAccountID And unc.Status='OK' " &
                                                 "where ven.Status='OK' and unc.RefNo='" & objRow("TrxCode") & "'"
                                                mAOPListID = pobjSqlRas.GetScalarAsString(mSQL)
                                                If mAOPListID <> "" Then
                                                    mQuerry = Replace(objRow("Querry"), "'',", "'" & mAOPListID & "',")
                                                    mQuerry = Replace(mQuerry, "'", "''")

                                                    pobjSqlRas.ExecuteNonQuerry("update AopQueue set Querry='" & mQuerry & "' where recid=" & objRow("RecID"))

                                                    mintWaitTime = 60
                                                End If
                                            ElseIf pobjOdbc.UpdtErr.Contains("The currency of the account must be either in home currency or the transaction currency") Then
                                                pobjSqlRas.ExecuteNonQuerry("update AopQueue set Status='XX' where RefNumber='" & objRow("RefNumber") & "'")

                                                mintWaitTime = 1
                                            End If
                                            '^_^20230224 add by 7643 -e-
                                        Else
                                            mintWaitTime = 120
                                        End If
                                        intElapsed = DateDiff(DateInterval.Second, dteStart, Now)
                                        TxtFeedBack.Text = Now & " " & objRow("LinkId") & " " & intElapsed & " " & objRow("B_I") & " " & objRow("TrxCode") & " FAILED !" & vbCrLf & TxtFeedBack.Text
                                        'pobjSqlRas.ExecuteNonQuerry("exec dbo.web_syncAOP")  '^_^20221123 add by 7643
                                        Exit For
                                    End If
                            End Select



                        Next
                    Case Else
                        If pobjSqlAop.Connection.State <> ConnectionState.Open Then
                            pobjSqlAop.Connect()
                        End If
                        For i = 0 To tblQueues.Rows.Count - 1
                            Dim objRow As DataRow = tblQueues.Rows(i)

                            If pobjSqlAop.ExecuteNonQuerry(objRow("Querry")) Then
                                If i = tblQueues.Rows.Count - 1 Then
                                    intElapsed = DateDiff(DateInterval.Second, dteStart, Now)
                                    If pobjSqlRas.ExecuteNonQuerry("update AopQueue set Status='RR',LstUpdate=getdate(),Count=Count+1 where Status='OK' and LinkId=" & objRow("LinkId")) Then
                                        TxtFeedBack.Text = Now & " " & objRow("LinkId") & " " & " " & intElapsed & " " & objRow("B_I") & " " & objRow("TrxCode") & " imported!" & vbCrLf & TxtFeedBack.Text
                                    Else
                                        MsgBox(objRow("B_I") & " " & objRow("TrxCode") & " imported, BUT can not change Status in RAS!")
                                    End If

                                End If
                            Else
                                pobjSqlRas.ExecuteNonQuerry("update AopQueue set LstUpdate=getdate(),Count=Count+1,Error=N'" _
                                                            & pobjSqlAop.SqlError.Replace("'", "*") & "' where Status='OK' and RecId=" & objRow("RecId"))

                                intElapsed = DateDiff(DateInterval.Second, dteStart, Now)
                                TxtFeedBack.Text = Now & " " & intElapsed & " " & objRow("B_I") & " " & objRow("TrxCode") & " FAILED !" & vbCrLf & TxtFeedBack.Text
                                pobjSqlRas.ExecuteNonQuerry("exec dbo.web_syncAOP")  '^_^20221123 add by 7643
                                mintWaitTime = 120
                                Exit For
                            End If
                        Next
                End Select
            End If

        Catch ex As Exception
            TxtFeedBack.Text = Now & " " & ex.Message & vbCrLf & TxtFeedBack.Text
            pobjSqlRas.ExecuteNonQuerry("exec dbo.web_syncAOP")  '^_^20221123 add by 7643
            mintWaitTime = 300
        Finally
            TxtFeedBack.Text = Now & "  Wait " & mintWaitTime & " seconds" & vbCrLf & TxtFeedBack.Text

            dteEndTime = DateAdd(DateInterval.Second, mintWaitTime, Now)
            Do While DateDiff(DateInterval.Second, Now, dteEndTime) > 0
                If mblnQuit Then
                    Exit Do
                End If
                My.Application.DoEvents()
                Threading.Thread.Sleep(500)
            Loop

        End Try
        If mblnQuit Then
            TxtFeedBack.Text = Now & " Started " & strProd & vbCrLf & TxtFeedBack.Text
            Return False
        Else
            GoTo StartPosition
        End If

        Return True
    End Function

    Private Sub barStop_Click(sender As Object, e As EventArgs) Handles barStop.Click
        mblnQuit = True
        TxtFeedBack.Text = Now & " Stopped!" & vbCrLf & TxtFeedBack.Text
    End Sub

    Private Sub barAllOdbc_Click(sender As Object, e As EventArgs) Handles barAllOdbc.Click
        ProcessQ("All", "ODBC")
    End Sub

    Private Function DowloadBspBillData(strQuerry As String, strCity As String) As Boolean
        Dim tblBIll As DataTable
        Dim dteStart As Date = Now
        tblBIll = pobjOdbc.GetDataTable(strQuerry)
        Dim objBulkCopy As New SqlClient.SqlBulkCopy(pobjSqlRas.Connection)
        'Set the database table name
        objBulkCopy.DestinationTableName = "lib.dbo.AOPBill"
        objBulkCopy.ColumnMappings.Add("TxnID", "TxnID")
        objBulkCopy.ColumnMappings.Add("TimeCreated", "TimeCreated")
        'objBulkCopy.ColumnMappings.Add("TimeModified", "TimeModified")
        objBulkCopy.ColumnMappings.Add("TxnNumber", "TxnNumber")
        'objBulkCopy.ColumnMappings.Add("VendorRefListID", "VendorRefListID")
        objBulkCopy.ColumnMappings.Add("VendorRefFullName", "VendorRefFullName")
        'objBulkCopy.ColumnMappings.Add("APAccountRefListID", "APAccountRefListID")
        objBulkCopy.ColumnMappings.Add("APAccountRefFullName", "APAccountRefFullName")
        objBulkCopy.ColumnMappings.Add("TxnDate", "TxnDate")
        objBulkCopy.ColumnMappings.Add("DueDate", "DueDate")
        objBulkCopy.ColumnMappings.Add("AmountDue", "AmountDue")
        objBulkCopy.ColumnMappings.Add("RefNumber", "RefNumber")
        objBulkCopy.ColumnMappings.Add("Memo", "Memo")
        'objBulkCopy.ColumnMappings.Add("IsPaid", "IsPaid")
        objBulkCopy.ColumnMappings.Add("City", "City")

        Try
            pobjSqlRas.ExecuteNonQuerry("delete lib.dbo.AopBill where City='" & strCity & "'")
            objBulkCopy.WriteToServer(tblBIll)
            Return True
        Catch ex As Exception
            Append2TextFile(ex.Message)
            Return False
        End Try

        'MsgBox(DateDiff(DateInterval.Second, dteStart, Now))
    End Function
    Private Function DowloadBspVendorCreditData(strQuerry As String, strCity As String) As Boolean
        Dim tblBIll As DataTable
        Dim dteStart As Date = Now
        tblBIll = pobjOdbc.GetDataTable(strQuerry)
        Dim objBulkCopy As New SqlClient.SqlBulkCopy(pobjSqlRas.Connection)
        'Set the database table name
        objBulkCopy.DestinationTableName = "lib.dbo.AopVendorCredit"
        objBulkCopy.ColumnMappings.Add("TxnID", "TxnID")
        objBulkCopy.ColumnMappings.Add("TimeCreated", "TimeCreated")
        'objBulkCopy.ColumnMappings.Add("TimeModified", "TimeModified")
        objBulkCopy.ColumnMappings.Add("TxnNumber", "TxnNumber")
        'objBulkCopy.ColumnMappings.Add("VendorRefListID", "VendorRefListID")
        objBulkCopy.ColumnMappings.Add("VendorRefFullName", "VendorRefFullName")
        'objBulkCopy.ColumnMappings.Add("APAccountRefListID", "APAccountRefListID")
        objBulkCopy.ColumnMappings.Add("APAccountRefFullName", "APAccountRefFullName")
        objBulkCopy.ColumnMappings.Add("TxnDate", "TxnDate")
        objBulkCopy.ColumnMappings.Add("CreditAmount", "CreditAmount")
        objBulkCopy.ColumnMappings.Add("RefNumber", "RefNumber")
        objBulkCopy.ColumnMappings.Add("Memo", "Memo")
        objBulkCopy.ColumnMappings.Add("City", "City")
        Try
            pobjSqlRas.ExecuteNonQuerry("delete lib.dbo.AopVendorCredit where City='" & strCity & "'")
            objBulkCopy.WriteToServer(tblBIll)
            Return True
        Catch ex As Exception
            Append2TextFile(ex.Message)
            Return False
        End Try

        'MsgBox(DateDiff(DateInterval.Second, dteStart, Now))
    End Function

    Private Sub barDeleteDupCust4HAN_Click(sender As Object, e As EventArgs) Handles barDeleteDupCust4HAN.Click
        Dim tblHanCust As DataTable = pobjSqlRas.GetDataTable("Select * from Customer where Status='ok' and AOPTravelListID<>'' and City='HAN'")
        For Each objRow As DataRow In tblHanCust.Rows
            If pobjOdbc.ExecuteNonQuerry("delete from Customer where ListId='" & objRow("AOPTravelListID") & "'") Then
                pobjSqlRas.ExecuteNonQuerry("Update Customer set AOPTravelListID='' where RecId=" & objRow("RecId"))
            End If
        Next

    End Sub
End Class
