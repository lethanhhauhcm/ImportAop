Imports System.Data.Odbc

Public Class clsOdbc
    Dim mcnxConnection As New OdbcConnection
    Dim mstrCnxErr As String
    Dim mstrUpdtErr As String
    Dim mstrConnectionString As String
    Dim msglLastInsertedId As Single



    Public Function GetConnecttionString(ByVal strFileName As String) As Boolean
        Dim objFile As System.IO.StreamReader
        Dim strFullPath As String

        strFullPath = System.AppDomain.CurrentDomain.BaseDirectory() & "\" & strFileName
        Try
            objFile = New System.IO.StreamReader(strFullPath)
            mstrConnectionString = objFile.ReadLine()
            objFile.Close()
            objFile.Dispose()
            GetConnecttionString = True
        Catch ex As Exception
            mstrCnxErr = ex.Message
            GetConnecttionString = False
        End Try

    End Function


    Public Function Disconnect() As Boolean
        If mcnxConnection.State = ConnectionState.Open Then
            mcnxConnection.Dispose()
        End If
        Return True
    End Function

    Public Function ExecuteNonQuerry(ByVal strQuerry As String) As Boolean
        Try
            Dim cmdOdbc As New OdbcCommand
            cmdOdbc.Connection = mcnxConnection
            cmdOdbc.CommandText = strQuerry
            cmdOdbc.CommandTimeout = 128
            cmdOdbc.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            mstrUpdtErr = vbNewLine & Now & vbTab & ex.Message & vbNewLine & strQuerry
            Append2TextFile(mstrUpdtErr)
            Return False
        End Try


    End Function
    Public Function UpdateListOfQuerries(ByVal lstQuerries As List(Of String) _
                                     , Optional blnGetLastInsertedRecId As Boolean = False) As Boolean

        Dim trcSql As OdbcTransaction
        Dim i As Integer
        If mcnxConnection.State = ConnectionState.Closed Then
            mcnxConnection.Open()
        End If


        trcSql = mcnxConnection.BeginTransaction
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection
        cmdOdbc.Transaction = trcSql
        Try
            For i = 0 To lstQuerries.Count - 1
                cmdOdbc.CommandText = lstQuerries(i)
                cmdOdbc.CommandTimeout = 10000
                cmdOdbc.ExecuteNonQuery()
                If blnGetLastInsertedRecId AndAlso UCase(Mid(lstQuerries(i), 1, 6)) = "INSERT" Then
                    cmdOdbc.CommandText = "select SCOPE_IDENTITY()"
                    msglLastInsertedId = cmdOdbc.ExecuteScalar
                End If
            Next
            trcSql.Commit()
            Return True
        Catch ex As Exception
            mstrUpdtErr = vbNewLine & ex.Message & vbNewLine & lstQuerries(i)
            trcSql.Rollback()
            Append2TextFile(mstrUpdtErr)
            Return False
        End Try
    End Function
    Public Function Update(ByVal arrQuerries() As String) As Boolean

        Dim trcSql As OdbcTransaction
        Dim i As Integer
        If mcnxConnection.State = ConnectionState.Closed Then
            mcnxConnection.Open()
        End If

        trcSql = mcnxConnection.BeginTransaction
        Dim cmdOdbc As OdbcCommand
        cmdOdbc.Connection = mcnxConnection
        cmdOdbc.Transaction = trcSql
        Try
            For i = LBound(arrQuerries) To UBound(arrQuerries)
                If Not arrQuerries(i) Is Nothing And arrQuerries(i) <> "" Then
                    cmdOdbc.CommandText = arrQuerries(i)
                    cmdOdbc.ExecuteNonQuery()
                    If UCase(Mid(arrQuerries(i), 1, 6)) = "INSERT" Then
                        cmdOdbc.CommandText = "select SCOPE_IDENTITY()"
                        msglLastInsertedId = cmdOdbc.ExecuteScalar
                    End If
                End If
            Next
            trcSql.Commit()
            Update = True
        Catch ex As Exception
            mstrUpdtErr = ex.Message & vbCrLf & arrQuerries(i)
            Update = False
            trcSql.Rollback()
            Append2TextFile(mstrUpdtErr)
        End Try
    End Function

    Public Function GetRoe(ByVal strCur As String) As Decimal
        'Purpose: Get ROE in RAS
        Dim cmdOdbc As New OdbcCommand
        Dim strQuerry As String

        cmdOdbc.Connection = mcnxConnection
        strQuerry = "Select top 1 BSR from ForEX where IsActive='Y'"
        strQuerry = strQuerry & " and ApplyROETo like '%TS%'"
        strQuerry = strQuerry & " and Currency='" & strCur & "' order by EffectDate desc"

        cmdOdbc.CommandText = strQuerry
        GetRoe = cmdOdbc.ExecuteScalar

    End Function


    Public Function Apt2City(ByVal strAirport As String) As String

        'purpose: Get the City code
        'input: Airport code
        'Output: City code

        Dim cmdOdbc As New OdbcCommand
        Dim strQuerry As String
        cmdOdbc.Connection = mcnxConnection
        strQuerry = "Select City from City where Airport='" & strAirport & "'"

        cmdOdbc.CommandText = strQuerry
        Apt2City = cmdOdbc.ExecuteScalar
    End Function
    Public Function Apt2Country(ByVal strAirport As String) As String
        'purpose: Get the Country code
        'input: Airport code
        'Output: Country code

        Dim cmdOdbc As New OdbcCommand
        Dim strQuerry As String
        cmdOdbc.Connection = mcnxConnection
        strQuerry = "Select Country from City where Airport='" & strAirport & "'"

        cmdOdbc.CommandText = strQuerry
        Apt2Country = cmdOdbc.ExecuteScalar
    End Function
    Public Function GetCity(ByVal strAirportCode) As String
        'Tim ma thanh pho cho ma san bay
        'Input: Ma san bay
        'Output: Ma nuoc
        'Pre-requisite: Can ket noi TVCS
        Dim strQry As String = ""
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQry = "select City from City where Airport ='" & strAirportCode & "'"
        cmdOdbc.CommandText = strQry
        GetCity = cmdOdbc.ExecuteScalar
    End Function
    Public Function GetCityCodeByName(ByVal strCityName As String) As String
        'Tim ma thanh pho cho ma san bay
        'Input: Ma san bay
        'Output: Ma nuoc
        'Pre-requisite: Can ket noi TVCS
        Return GetScalarAsString("select top 1 City from CityCode where CityName ='" & strCityName & "'")
    End Function
    Public Function GetCityCodeByNameHotFile(ByVal strCityName As String) As String
        'Tim ma thanh pho cho ma san bay
        'Input: Ma san bay
        'Output: Ma nuoc
        'Pre-requisite: Can ket noi TVCS
        Return GetScalarAsString("select top 1 CityName from TblCityCode_LoadExcel where CityName ='" & strCityName & "'")
    End Function
    Public Function GetCityName(ByVal strAirportCode As String, Optional ByVal blnWzNewCnx As Boolean = False) As String
        'Tim ma thanh pho cho ma san bay
        'Input: Ma san bay
        'Output: Ma nuoc
        'Pre-requisite: Can ket noi TVCS
        Dim strQry As String = ""
        Dim cmdOdbc As New OdbcCommand

        If blnWzNewCnx Then
            cmdOdbc.Connection = mcnxConnection
            cmdOdbc.Connection.Open()
        Else
            cmdOdbc.Connection = mcnxConnection
        End If

        strQry = "select CityName from City where Airport ='" & strAirportCode & "'"
        cmdOdbc.CommandText = strQry
        GetCityName = UCase(Trim(cmdOdbc.ExecuteScalar))
    End Function
    Public Function GetCountryName(ByVal strCountryCode) As String
        'Tim ma nuoc cho ma thanh pho
        'Input: Ma san bay
        'Output: Ma nuoc
        'Pre-requisite: Can ket noi TVCS
        Dim strQuerry As String = ""
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select CountryName from Country where Country ='" & strCountryCode & "'"
        cmdOdbc.CommandText = strQuerry
        GetCountryName = cmdOdbc.ExecuteScalar
    End Function
    Public Function GetCountry(ByVal strAirportCode) As String
        'Tim ma nuoc cho ma thanh pho
        'Input: Ma san bay
        'Output: Ma nuoc
        'Pre-requisite: Can ket noi TVCS
        Dim strQuerry As String = ""
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select Country from City where Airport ='" & strAirportCode & "'"
        cmdOdbc.CommandText = strQuerry
        GetCountry = cmdOdbc.ExecuteScalar
    End Function
    Public Function GetRtgType(Optional ByVal strRtg As String = "") As String
        'Purpose: Tim loai hanh trinh tu PNR
        'Input: Hanh trinh chi co City va Carrier
        'Output: Loai hanh trinh INTL hoac XXDOM
        Dim ArrCity() As String
        Dim strRtgType As String = ""
        Dim strExCountry As String = ""
        Dim bytCityCount As Integer
        Dim i As Integer

        bytCityCount = (Len(strRtg) + 2) / 5
        ReDim ArrCity(0 To bytCityCount - 1)
        For i = 0 To bytCityCount - 1
            ArrCity(i) = Mid(strRtg, i * 5 + 1, 3)
            If i = 0 Then
                strExCountry = GetCountry(ArrCity(i))
                strRtgType = strExCountry & "DOM"
            ElseIf strExCountry <> GetCountry(ArrCity(i)) Then
                strRtgType = "INTL"
                Exit For
            End If
        Next
        GetRtgType = strRtgType
    End Function
    Public Function GetCar2C(ByVal strCar3D As String) As String

        'purpose: Get the 2-character code of airlines
        'input: 3 digit code of airline        
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection
        strQuerry = "select ALCode from Airline where DocCode='" & strCar3D & "'"
        cmdOdbc.CommandText = strQuerry
        GetCar2C = cmdOdbc.ExecuteScalar
    End Function
    Public Function GetCarName(ByVal strCarCode As String, Optional ByVal blnWzNewCnx As Boolean = False) As String

        'purpose: Get the NAME of airlines
        'input: 2-character code of airlines
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand

        If blnWzNewCnx Then
            cmdOdbc.Connection = mcnxConnection
            cmdOdbc.Connection.Open()
        Else
            cmdOdbc.Connection = mcnxConnection
        End If
        strQuerry = "select ALName from Airline where ALCode='" & strCarCode & "'"
        cmdOdbc.CommandText = strQuerry
        GetCarName = UCase(Trim(cmdOdbc.ExecuteScalar))
    End Function
    Public Function GetIsi(ByVal strAptCode As String) As String
        If GetCountry(strAptCode) = "VN" Then
            GetIsi = "SITI"
        Else
            GetIsi = "SOTO"
        End If
    End Function

    Public Function CreateEmail(ByVal intCus As Integer, ByVal strSubj As String, ByVal strMsg As String,
                            ByVal strFrom As String, ByVal strEmailGroup As String) As Boolean
        Dim strColumns As String
        Dim strValues As String = ""
        Dim strQuerry As String
        Dim intResult As Integer
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        'If intCus = 0 Then Exit Function
        strColumns = "CustID,Subj,Msg,Frm,Dept"
        strValues = strValues & "'" & intCus & "'"
        strValues = strValues & ",'" & strSubj & "'"
        strValues = strValues & ",'" & strMsg & "'"
        strValues = strValues & ",'" & strFrom & "'"
        strValues = strValues & ",'" & strEmailGroup & "'"
        strQuerry = "insert into EmailLog ("
        strQuerry = strQuerry & strColumns & ") values ("
        strQuerry = strQuerry & strValues & ")"

        cmdOdbc.CommandText = strQuerry
        intResult = cmdOdbc.ExecuteNonQuery()
        If intResult > 0 Then
            CreateEmail = True
        Else
            CreateEmail = False
        End If

    End Function
    Public Function GetTktEntry(ByVal strValCar As String, ByVal mstrTktBox As String _
                            , ByVal strLocaion As String) As String
        'Purpose: Find additional Ticket entry to be inserted into TST
        'Input: Validatint carrier, ISI
        'Output: Ticket entry

        Dim strQry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQry = "select Value from TktEntries"
        strQry = strQry & " where Status ='OK' and ValCar ='" & strValCar & "'"
        strQry = strQry & " and '" & Format(Now, "dd-mmm-yyyy hh:nn:ss") & "' between TktDateFrom and TktDateTo"
        strQry = strQry & " and Catergory ='" & mstrTktBox & "'"
        strQry = strQry & " and Location='" & strLocaion & "'"
        cmdOdbc.CommandText = strQry
        GetTktEntry = cmdOdbc.ExecuteScalar
    End Function




    Public Function GetCustId(ByVal strCustShortName As String) As String
        'Tim Customer Id
        'Input: Customer short name
        'Output: Customer Id
        'Pre-requisite: Can ket noi TVCS
        Dim strQry As String = ""
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQry = "select RecId from CustomerList where Status='OK' " _
            & " and CustShortName ='" & strCustShortName & "'"
        cmdOdbc.CommandText = strQry
        GetCustId = cmdOdbc.ExecuteScalar
    End Function
    Public Function GetUsdRoeInRas() As Decimal
        Dim strQuerry As String
        Dim strResult As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection
        strQuerry = "Select Details from MISC where CAT='RoeQuerry'"
        cmdOdbc.CommandText = strQuerry
        strResult = cmdOdbc.ExecuteScalar

        cmdOdbc.CommandText = strResult
        GetUsdRoeInRas = cmdOdbc.ExecuteScalar
    End Function

    Public Function DuplicateGO_Air(ByVal strValCar As String, ByVal strTKNO As String, ByVal strSRV As String) As Boolean
        'Purpose: Check if insert querry will create duplicate ticket nbrs in TKT_1A
        'Input: Ticket number & SRV
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection
        Dim sglResult As Single

        strTKNO = Replace(strTKNO, "-", "")
        strTKNO = Replace(strTKNO, " ", "")
        strTKNO = Mid(strTKNO, 4)

        strQuerry = "select * from GO_Air where Carrier='" & strValCar _
                & "' and TKNO='" & strTKNO & "' and SRV='" & strSRV & "'"
        cmdOdbc.CommandText = strQuerry
        sglResult = cmdOdbc.ExecuteScalar
        If sglResult = 0 Then
            DuplicateGO_Air = False
        Else
            DuplicateGO_Air = True
        End If
    End Function
    Public Function DeleteGO_Travel(ByVal sglRecId As Single) As Boolean
        'Purpose: Delete duplicate record in GO_Travel
        'Input: Record ID
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "DELETE from GO_Travel where RecId=" & sglRecId
        cmdOdbc.CommandText = strQuerry
        cmdOdbc.ExecuteNonQuery()
        DeleteGO_Travel = True
    End Function
    Public Function GetGO_COS(ByVal strBkgCls) As String
        'Purpose: Convert Bkg class into Global One's class of service
        'Input: Booking class
        'Output: GO's class of service

        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection


        strQuerry = "select Details from GO_MISC where CAT='COS' and VAL='" & strBkgCls & "'"
        cmdOdbc.CommandText = strQuerry
        GetGO_COS = cmdOdbc.ExecuteScalar()
        If GetGO_COS = "" Then
            GetGO_COS = "Y"
        End If
    End Function
    Public Function GetGO_ClassOfService(ByVal strBkgCls As String) As String
        'Purpose: Convert Bkg class into Global One's class of service
        'Input: Booking class
        'Output: GO's class of service

        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand


        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select RMK from GO_MISC where CAT='COS' and VAL='" & strBkgCls & "'"
        cmdOdbc.CommandText = strQuerry
        GetGO_ClassOfService = cmdOdbc.ExecuteScalar()
        If GetGO_ClassOfService = "" Then
            GetGO_ClassOfService = "ECONOMY"
        End If
    End Function



    Public Function GetStoredConditions() As String()
        'Purpose: Get stored condition
        'Input: 
        'Output: array

        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        Dim drResult As OdbcDataReader
        Dim arrResult(0 To 0) As String
        Dim i As Integer
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select Details from GO_MISC where CAT='Conditions'"

        cmdOdbc.CommandText = strQuerry

        drResult = cmdOdbc.ExecuteReader
        If Not drResult Is Nothing Then
            Do While drResult.Read
                ReDim Preserve arrResult(0 To i)
                arrResult(i) = Replace(drResult("Details"), vbCrLf, vbLf)
                i = i + 1
            Loop
        End If
        drResult.Close()
        GetStoredConditions = arrResult
    End Function
    Public Sub LoadValCarFt(ByRef cboInput As ComboBox)
        Dim strQuerry As String
        strQuerry = "select distinct Value from autoissue1 where Catergory='Valcar' and Status='OK' order by value"
        LoadCombo(cboInput, strQuerry)

    End Sub
    Public Function LoadCombo(ByRef cboInput As ComboBox, ByVal strQuerry As String) As ComboBox
        Dim daConditions As OdbcDataAdapter
        Dim dsConditions As New DataSet

        daConditions = New OdbcDataAdapter(strQuerry, mcnxConnection)
        daConditions.Fill(dsConditions, "RESULT")
        cboInput.DataSource = dsConditions.Tables("RESULT")
        cboInput.DisplayMember = "Value"
        cboInput.ValueMember = "Value"
        LoadCombo = cboInput
        dsConditions.Dispose()
        daConditions.Dispose()

    End Function
    Public Function LoadComboVal(ByVal cboInput As ComboBox, ByVal strQuerry As String) As ComboBox
        Dim daConditions As OdbcDataAdapter
        Dim dsConditions As New DataSet

        daConditions = New OdbcDataAdapter(strQuerry, mcnxConnection)
        daConditions.Fill(dsConditions, "RESULT")
        cboInput.DataSource = dsConditions.Tables("RESULT")
        cboInput.DisplayMember = "Display"
        cboInput.ValueMember = "Value"
        LoadComboVal = cboInput
        dsConditions.Dispose()
        daConditions.Dispose()

    End Function
    Public Function DeleteGridViewRow(ByRef dbInput As DataGridView, ByVal strQuerry As String) As Boolean
        Dim strMessage As String
        Dim i As Integer

        strMessage = "Do you want to delete the following record?" & vbCrLf
        With dbInput
            For i = 0 To dbInput.Columns.Count - 1
                If .Columns.Item(i).Visible Then
                    strMessage = strMessage & .Columns.Item(i).HeaderText & ": " _
                                & .CurrentRow.Cells.Item(i).Value & vbCrLf
                End If
            Next
        End With

        If MsgBox(strMessage, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Dim arrQuerry(0 To 0) As String
            arrQuerry(0) = strQuerry
            If Update(arrQuerry) Then
                Return True
            Else
                Return False
            End If
        End If
        Return True
    End Function

    Public Function UpdateDefaultHierachy(ByVal intCmc As Integer) As Boolean
        'Purpose: update hierachy in GO_DefaultHierachy for record with blank hierachy 2 - 5
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection
        Dim intHierNbr As Integer

        For intHierNbr = 2 To 5
            strQuerry = "update GO_Travel set Hierachy" & intHierNbr & "=(select top 1 Value" _
                            & " from GO_DefaultHierachy where Status='OK' and CMC=" & intCmc _
                            & " and HierachyNbr=" & intHierNbr _
                            & ") where Hierachy" & intHierNbr & "='' and CMC=" & intCmc _
                            & "and (select top 1 Value from GO_DefaultHierachy where Status='OK' and CMC=" _
                            & intCmc & " and HierachyNbr=" & intHierNbr & ") <>''"

            cmdOdbc.CommandText = strQuerry
            cmdOdbc.ExecuteNonQuery()
        Next

        UpdateDefaultHierachy = True
    End Function
    Public Function UpdateGoTravelDOI(ByVal intRecId As Integer) As Boolean
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "update GO_Travel set DOI=(select top 1 DOI from GO_Air where TravelId=" & intRecId _
                & " order by DOI) where RecId=" & intRecId

        cmdOdbc.CommandText = strQuerry
        cmdOdbc.ExecuteNonQuery()

        UpdateGoTravelDOI = True
    End Function
    Public Function UpdateGoTravelBkgTool(ByVal intRecId As Integer, ByVal strBkgTool As String) As Boolean
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "update GO_Travel set BkgTool='" & strBkgTool _
                & "' where BkgMethod='G' and RecId=" & intRecId

        cmdOdbc.CommandText = strQuerry
        cmdOdbc.ExecuteNonQuery()
        UpdateGoTravelBkgTool = True

    End Function
    Public Function UpdateGoTravelDefaultValues(ByVal intRecId As Integer, ByVal strBkgTool As String) As Boolean
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "update GO_Travel set BkgDate=DOI, BkgTool='" & strBkgTool _
                & "' where BkgMethod='G' and RecId=" & intRecId

        cmdOdbc.CommandText = strQuerry
        cmdOdbc.ExecuteNonQuery()
        UpdateGoTravelDefaultValues = True

    End Function
    Public Function UpdateGoAirDefaultValues(ByVal intRecId As Integer, ByVal strTkno As String _
                                        , ByVal strDepDates As String, ByVal strArrDates As String _
                                        , ByVal strFltNbrs As String, ByVal strETD As String _
                                        , ByVal strETA As String, ByVal strSOs As String) As Boolean
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "update GO_Air set RefFare=Fare,LowestFare=Fare,DepDates='" & strDepDates _
                & "', ArrDateIndicators='" & strArrDates _
                & "', FltNbrs='" & strFltNbrs _
                & "', etd='" & strETD _
                & "', eta='" & strETA _
                & "', SO='" & strSOs _
                & "' where Recid=" & intRecId _
                & "  and TKNO='" & strTkno & "'"

        cmdOdbc.CommandText = strQuerry
        cmdOdbc.ExecuteNonQuery()
        UpdateGoAirDefaultValues = True

    End Function
    Public Sub LoadDataGridView(ByRef dgInput As DataGridView, ByVal strQuerry As String)
        Dim daConditions As SqlClient.SqlDataAdapter
        Dim dsConditions As New DataSet

        daConditions = New SqlClient.SqlDataAdapter(strQuerry, mstrConnectionString)
        daConditions.Fill(dsConditions, "Result")
        dgInput.DataSource = dsConditions.Tables("Result")
        'dgInput.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dsConditions.Dispose()
        daConditions.Dispose()
    End Sub
    Public Function CreateCityPair(ByVal strCityPair As String) As Boolean
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand

        If CityPairExist(strCityPair) Then
            Return False
        End If
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "Insert into GO_CITYPAIR (Citypair) values('" & strCityPair & "')"

        cmdOdbc.CommandText = strQuerry
        cmdOdbc.ExecuteNonQuery()

        Return True
    End Function
    Public Function CityPairExist(ByVal strCityPair As String) As Boolean
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select CityPair from GO_CITYPAIR where CityPair='" & strCityPair & "'"

        cmdOdbc.CommandText = strQuerry
        If cmdOdbc.ExecuteScalar = "" Then
            CityPairExist = False
        Else
            CityPairExist = True
        End If
    End Function
    Public Function UpdateETAs(ByVal intRecId As Integer, ByVal strETAs As String) As Boolean
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "update go_air set ETA='" & strETAs & "' where Recid=" & intRecId

        cmdOdbc.CommandText = strQuerry
        cmdOdbc.ExecuteNonQuery()

        UpdateETAs = True
    End Function
    Public Function GetETA(ByVal strCityPair As String, ByVal strCar As String _
                        , ByVal strFltNbr As String) As String
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select ETA from GO_AirSC where CityPair='" & strCityPair _
                & "' and Car='" & strCar & "' and FltNbr='" & strFltNbr & "'"

        cmdOdbc.CommandText = strQuerry
        GetETA = cmdOdbc.ExecuteScalar()

    End Function
    Public Function GetElapsedTime(ByVal strCityPair As String, ByVal strCar As String
                        ) As Integer
        'Purpose:  
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select ElapsedTime from GO_AirSC where CityPair='" & strCityPair _
                & "' and Car='" & strCar & "'"

        cmdOdbc.CommandText = strQuerry
        GetElapsedTime = cmdOdbc.ExecuteScalar()

    End Function
    Public Function DefaultFareApplied(ByVal intCmc As Integer) As Boolean
        'Purpose:  check if RefFare and low fare will be the same with Paid fare
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select DefaultFare from GO_CompanyInfo where Cmc=" & intCmc

        cmdOdbc.CommandText = strQuerry
        DefaultFareApplied = cmdOdbc.ExecuteScalar()
    End Function
    Public Function RasShortNameExist(ByVal strShortName) As Boolean
        'Purpose:  check if RasShortName exists
        'Output: Y/N
        Dim strQuerry As String
        Dim cmdOdbc As New OdbcCommand
        cmdOdbc.Connection = mcnxConnection

        strQuerry = "select RecId from Customerlist where Status='OK' and CustShortName='" _
                & strShortName & "'"

        cmdOdbc.CommandText = strQuerry
        If cmdOdbc.ExecuteScalar <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function CitiesExist(ByVal strCities As String) As Boolean
        Dim arrCities() As String
        Dim i As Integer
        If Mid(strCities, 1, 1) = "-" Then
            strCities = Mid(strCities, 2)
        End If
        arrCities = Split(strCities, ",")
        For i = 0 To UBound(arrCities)
            If GetCityName(arrCities(i)) = "" Then
                Return False
            End If
        Next
        Return True
    End Function
    Public Function CountriesExist(ByVal strCountries As String) As Boolean
        Dim arrCountries() As String
        Dim i As Integer
        If Mid(strCountries, 1, 1) = "-" Then
            strCountries = Mid(strCountries, 2)
        End If
        arrCountries = Split(strCountries, ",")
        For i = 0 To UBound(arrCountries)
            If GetCountryName(arrCountries(i)) = "" Then
                Return False
            End If
        Next
        Return True
    End Function
    Public Function CarriersExist(ByVal strCarriers As String) As Boolean
        Dim arrCarriers() As String
        Dim i As Integer
        If Mid(strCarriers, 1, 1) = "-" Then
            strCarriers = Mid(strCarriers, 2)
        End If
        arrCarriers = Split(strCarriers, ",")
        For i = 0 To UBound(arrCarriers)
            If GetCarName(arrCarriers(i)) = "" Then
                Return False
            End If
        Next
        Return True
    End Function
    Public Function CheckFormatRtgType(ByVal strRtgType As String) As Boolean

        If strRtgType = "" Then
            Return True
        ElseIf strRtgType = "INTL" Then
            Return True
        ElseIf Mid(strRtgType, 3) = "DOM" AndAlso CountriesExist(Mid(strRtgType, 1, 2)) Then
            Return True
        Else
            CheckFormatRtgType = False
        End If
    End Function


    Public Function CheckFormatOffcId1A(ByVal strOffcId As String) As Boolean
        If Len(strOffcId) <> 9 Then
            Return False
        ElseIf Not CitiesExist(Mid(strOffcId, 1, 3)) Then
            Return False
        End If
        Return True
    End Function
    Public Function CheckFormatPaxType(ByVal strPaxType As String) As Boolean

        CheckFormatPaxType = True
        If strPaxType <> "" Then
            Dim arrPaxTypes() As String
            Dim i As Integer
            arrPaxTypes = Split(strPaxType, ",")
            For i = 0 To UBound(arrPaxTypes)
                Select Case arrPaxTypes(i)
                    Case "ADL", "CHD", "INF"
                    Case Else
                        CheckFormatPaxType = False
                        Exit Function
                End Select
            Next
        End If
    End Function
    Public Function CheckFormatCountryCode(ByVal strCountryCode As String _
                                        , ByVal blnAllowBlank As Boolean) As Boolean

        If strCountryCode = "" AndAlso blnAllowBlank Then
            Return True
        ElseIf CountriesExist(strCountryCode) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetLocationByRasTable() As String
        Dim strResult As String
        strResult = GetScalarAsString("Select top 1 VAL from POS")
        If strResult.StartsWith("0") Then
            Return "SGN"
        ElseIf strResult.StartsWith("3") Then
            Return "HAN"
        Else
            Return ""
        End If

    End Function
    Public Function GetScalarAsString(ByVal strQuerry As String) As String
        Dim cmdOdbc As New OdbcCommand
        Dim strResult As String = ""
        cmdOdbc.Connection = mcnxConnection
        cmdOdbc.CommandText = strQuerry
        cmdOdbc.CommandTimeout = 512
        Try
            strResult = cmdOdbc.ExecuteScalar()
        Catch ex As Exception
            mstrCnxErr = ex.Message
            Append2TextFile("SQL error:" & vbNewLine & strQuerry _
                        & vbNewLine & ex.Message)
        End Try

        Return strResult
    End Function
    Public Function GetScalarAsDecimal(ByVal strQuerry As String) As Decimal
        Dim decResult As Decimal

        Decimal.TryParse(GetScalarAsString(strQuerry), decResult)
        Return decResult
    End Function
    Public Function GetDataTable(ByVal strQuerry As String) As System.Data.DataTable
        Dim cmdOdbc As New OdbcCommand
        Dim tblResult As System.Data.DataTable
        cmdOdbc.Connection = mcnxConnection
        cmdOdbc.CommandText = strQuerry
        cmdOdbc.CommandTimeout = 512
        Dim daResult As New OdbcDataAdapter
        Dim dsResult As New DataSet
        daResult.SelectCommand = cmdOdbc
        daResult.Fill(dsResult)
        tblResult = dsResult.Tables(0)

        Return tblResult
    End Function
    Public Function CheckDupHotFile(strBAED As String, strCar As String) As Boolean
        'Output: Y/N - Khong bi dup/Co bi dup
        Dim strDupCheck As String

        strDupCheck = "select ID from RAS2K7.DBO.UA_HOT where RMED ='" & strBAED _
                & "' and substring(TDNR,1,3)='" & strCar & "'"

        If GetScalarAsString(strDupCheck) = 0 Then
            Return False
        Else
            Return True
        End If
    End Function


    Public Function IsIncentiveCalculatedByDate(dteInput As Date, strTimeFrame As String, strShortName As String) As Boolean
        Dim intYear As Integer = dteInput.Year
        Dim intPeriod As Integer
        Dim strQuerry As String

        Select Case strTimeFrame
            Case "Month"
                intPeriod = dteInput.Month
            Case "Quarter"
                If dteInput.Month > 9 Then
                    intPeriod = 4
                ElseIf dteInput.Month > 6 Then
                    intPeriod = 3
                ElseIf dteInput.Month > 3 Then
                    intPeriod = 2
                ElseIf dteInput.Month > 0 Then
                    intPeriod = 1
                End If
            Case "HalfYear"
                intPeriod = IIf(dteInput.Month > 6, 2, 1)
            Case "Year"
                intPeriod = 1
        End Select

        strQuerry = "Select Top 1 RecId from Data1A_IncentiveCalc where IncType='Auto' and BookYear=" _
                & intYear & " and Period>=" & intPeriod & " and TimeFrame='" & strTimeFrame _
                & "' and ShortName='" & strShortName & "'"
        If GetScalarAsString(strQuerry) <> "" Then
            Return True
        Else
            Return False
        End If
        Return False
    End Function


    Public Property CnxErr() As String
        Get
            CnxErr = mstrCnxErr
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public Property UpdtErr() As String
        Get
            UpdtErr = mstrUpdtErr
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public Property Connection() As OdbcConnection
        Get
            Connection = mcnxConnection
        End Get
        Set(ByVal value As OdbcConnection)

        End Set
    End Property

    Public Property ConnectionString() As String
        Get
            ConnectionString = mstrConnectionString
        End Get
        Set(ByVal value As String)
            mstrConnectionString = value
        End Set
    End Property
    Public Property LastInsertedId() As Single
        Get
            LastInsertedId = msglLastInsertedId
        End Get
        Set(ByVal value As Single)
            msglLastInsertedId = value
        End Set
    End Property

End Class
