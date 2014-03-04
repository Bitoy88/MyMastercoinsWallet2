'Bitoy (c) 2013 

Imports System
Imports System.IO
Imports System.Net
Imports System.Numerics
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Security.Cryptography
Imports System.Web

Class Form1
    Const SendType As Integer = 0
    Dim aUpdated(7) As DateTime

    Public Function FindBitcoind(ByVal targetDirectory As String) As String
        Dim PathName As String = ""
        Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
        ' Process the list of files found in the directory. 
        Dim lFound As Boolean = False
        Dim fileName As String
        For Each fileName In fileEntries
            If fileName.IndexOf("bitcoind.exe") > 0 Then
                lFound = True
                PathName = targetDirectory
                Exit For
            End If
        Next fileName
        If Not lFound Then
            Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            ' Recurse into subdirectories of this directory. 
            Dim subdirectory As String
            For Each subdirectory In subdirectoryEntries
                PathName = FindBitcoind(subdirectory)
                If PathName <> "" Then
                    Exit For
                End If
            Next subdirectory
        End If
        Return PathName
    End Function

    Private Sub Form1_Release(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.FormClosed
        Dim bitcoindprocess() As Process = Process.GetProcessesByName("bitcoind")
        For Each Process As Process In bitcoindprocess
            Process.Kill()
        Next
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Me.Text += " 2.01.00"
        If My.Settings.ConnectString = "" Then
            '            My.Settings.ConnectString = "Data Source=tcp:s09.winhost.com;Initial Catalog=DB_66808_mymsc;User ID=DB_66808_mymsc_user;Password=88BottlesofBeer;Integrated Security=False;"
            My.Settings.ConnectString = "Data Source=.\SQLEXPRESS;AttachDbFilename=""" + CurDir() + "\MYMASTERCOINS.MDF"";Integrated Security=True;User Instance=True"
        End If
        txtConnectString.Text = My.Settings.ConnectString

        If Not File.Exists(My.Settings.BitcoindExe) Then
            If lOK("Can't find bitcoind.exe.  Would you like me to look for it (If not, a file dialog box will open and you have to manually find bitcoind.exe)?") Then
                My.Settings.BitcoindExe = FindBitcoind(Path.GetPathRoot(CurDir) + "Program Files (x86)")
                If My.Settings.BitcoindExe = "" Then
                    MsgBox("Can't find bitcoind.exe. Please enter it in the Settings tab.")
                Else
                    My.Settings.BitcoindExe += "\bitcoind.exe"
                End If
            Else
                OpenFileDialog1.FileName = "bitcoind.exe"
                If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                    My.Settings.BitcoindExe = OpenFileDialog1.FileName
                End If
            End If
            My.Settings.Save()
        End If

        txtBitcoindexe.Text = Trim(My.Settings.BitcoindExe)
        txtDataDir.Text = Trim(My.Settings.DataDir)
        If txtDataDir.Text = "" Then
            Process.Start(txtBitcoindexe.Text, "-daemon")
        Else
            Process.Start(txtBitcoindexe.Text, "-datadir=" + txtDataDir.Text + " -daemon")
        End If

        ResetaUpdated()

        If My.Settings.DTAddress Is Nothing Then
            My.Settings.DTAddress = GetAddressTable()
            My.Settings.Save()
            ResetProgbar(5)
            updateprogbar()
            Timer2.Enabled = True
        End If

        GVWalletSummary.DataSource = My.Settings.DTAddress
        cboAddress.DataSource = My.Settings.DTAddress
        cboAddress.DisplayMember = "Address"
        cboAddress.ValueMember = "Address"


        lblTBTC.Text = My.Settings.TBTC
        lblTMSC.Text = My.Settings.TMSC
        lblTTMSC.Text = My.Settings.TTMSC

        If My.Settings.DTLog Is Nothing Then
            My.Settings.DTLog = GetLogTable()
            My.Settings.Save()
        End If
        GVSentLog.DataSource = My.Settings.DTLog

        If My.Settings.DTAddressBook Is Nothing Then
            My.Settings.DTAddressBook = GetAddressBookTable()
            My.Settings.Save()
        End If

        txtSendtoBTCA.DataSource = My.Settings.DTAddressBook
        txtSendtoBTCA.DisplayMember = "Address"
        txtSendtoBTCA.ValueMember = "Address"

        Dim p() As Process
        p = Process.GetProcessesByName("bitcoind")
        If p.Count = 0 Then
            MsgBox("Could not start " + txtBitcoindexe.Text + ".  Please enter the correct file location of bitcoind.exe in the settings tab and restart the application.")
        End If

    End Sub

    Function GetAddressBookTable() As DataTable
        ' Create new DataTable instance.
        Dim table As New DataTable
        table.TableName = "AddressBookTable"
        ' Create four typed columns in the DataTable.
        table.Columns.Add("Address", GetType(String))
        table.Columns.Add("Description", GetType(String))
        Return table
    End Function

    Sub updateDTAddress()
        If My.Settings.DTAddress.Rows.Count > 0 Then
            Dim ToMSC As Double = 0
            Dim TTMSC As Double = 0
            Dim TBTC As Double = 0
            ResetProgbar(My.Settings.DTAddress.Rows.Count)
            For i = 0 To My.Settings.DTAddress.Rows.Count - 1
                updateprogbar()
                With My.Settings.DTAddress.Rows(i)
                    Dim Address As String = Trim(.Item("Address"))
                    Dim MSC As Double = 0
                    Dim TMSC As Double = 0
                    Dim BTC As Double = 0
                    If Address.Length > 0 Then
                        Dim AddressID As Integer = (New mymastercoins).GetAddressID(Address)
                        If AddressID > 0 Then
                            MSC = (New mymastercoins).GetAddressBalance(AddressID, 1)
                            TMSC = (New mymastercoins).GetAddressBalance(AddressID, 2)
                        End If
                    End If
                    '                Dim PostURL As String = "http://www.blockchain.info/q/addressbalance/" + Address
                    '                Dim Result As String = (New Bitcoin).getjson(PostURL)

                    Dim PostURL As String = "http://blockexplorer.com/q/addressbalance/" + Address
                    Dim Result As Double = Val((New Bitcoin).getjson(PostURL))

                    BTC = Result
                    .Item("BTC") = BTC
                    .Item("MSC") = MSC
                    .Item("TMSC") = TMSC
                    TBTC += BTC
                    ToMSC += MSC
                    TTMSC += TMSC
                End With
            Next
            clearprogbar()
            lblTBTC.Text = FormatNumber(TBTC, 8)
            lblTMSC.Text = FormatNumber(ToMSC, 8)
            lblTTMSC.Text = FormatNumber(TTMSC, 8)
            My.Settings.TBTC = lblTBTC.Text
            My.Settings.TMSC = lblTMSC.Text
            My.Settings.TTMSC = lblTTMSC.Text
            My.Settings.Save()
        End If
    End Sub
    Function GetAddressTable() As DataTable
        ' Create new DataTable instance.
        Dim table As New DataTable
        table.TableName = "Datatable"
        ' Create four typed columns in the DataTable.
        table.Columns.Add("Address", GetType(String))
        table.Columns.Add("BTC", GetType(Double))
        table.Columns.Add("MSC", GetType(Double))
        table.Columns.Add("TMSC", GetType(Double))
        Return table
    End Function
    Sub logTxID(ByVal TxID As String, ByVal TransType As String, ByVal Sender As String, ByVal Recipient As String, ByVal Currency As String, ByVal Amount As Double)
        If TxID <> "" Then
            My.Settings.DTLog.Rows.Add(Now(), TxID, TransType, Sender, Recipient, Currency, Amount)
        End If
    End Sub
    Function GetLogTable() As DataTable
        ' Create new DataTable instance.
        Dim table As New DataTable
        table.TableName = "Datatable"
        ' Create four typed columns in the DataTable.
        table.Columns.Add("TransDate", GetType(Date))
        table.Columns.Add("TxID", GetType(String))
        table.Columns.Add("TransType", GetType(String))
        table.Columns.Add("Sender", GetType(String))
        table.Columns.Add("Recipient", GetType(String))
        table.Columns.Add("Currency", GetType(String))
        table.Columns.Add("Amount", GetType(Double))
        Return table
    End Function
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdSendCoins.Click
        Dim CurrencyID As Integer = GetCurrencyInt(cboCurrencySend.Text)
        If CurrencyID > 0 Then
            If Val(txtAmount.Text) > 0 Then
                If lOK("Are you sure you want to send " + txtAmount.Text + " " + cboCurrencySend.Text + " to " + txtSendtoBTCA.Text) Then
                    Dim SenderID As Integer = (New mymastercoins).GetAddressID(cboAddress.Text)
                    Dim Balance As Double = (New mymastercoins).GetAddressBalance(SenderID, CurrencyID)
                    Dim CurrencyForSale As Double = (New mymastercoins).GetCurrencyForSale(SenderID, CurrencyID)

                    If Balance - CurrencyForSale >= txtAmount.Value Then
                        Dim CoinType As Integer = GetCurrencyInt(cboCurrencySend.Text)
                        Dim SatAmount As BigInteger = (New mlib).AmounttoSat(txtAmount.Value)
                        Dim TxID As String = SendCoin(cboAddress.Text, txtSendtoBTCA.Text, CoinType, SatAmount)
                        If TxID <> "" Then
                            txtSendtoBTCA.Text = ""
                            txtAmount.Value = 0
                        End If
                    Else
                        MessageBox.Show("You can only send " + (Balance - CurrencyForSale).ToString + " " + cboCurrencySend.Text + ".  (Balance: " + Balance.ToString + " -  'For Sale':  " + CurrencyForSale.ToString + ")")
                    End If
                End If
            Else
                MessageBox.Show("Please enter an amount to send.")
            End If
        Else
            MessageBox.Show("Please select a Currency.")
        End If
    End Sub
    Function SendCoin(ByVal FromAddress As String, ByVal ToAddress As String, ByVal CoinType As String, ByVal SatAmount As BigInteger) As String
        Dim rawtx As String = (New mlib).encodetx(FromAddress, ToAddress, CoinType, SatAmount)
        Return signandsendrawtransaction(rawtx)
    End Function

    '                   RetVal = "Raw transaction is empty."

    Function signandsendrawtransaction(ByVal rawtx As String) As String
        Dim TxID As String = ""
        If rawtx <> "" Then
            ResetProgbar(4)
            updateprogbar()
            Call (New mlib).rpccalld("walletpassphrase", 2, Trim(txtBTCWalletPP.Text), "120", 0)
            Dim json2 As String = (New mlib).rpccalld("signrawtransaction", 1, rawtx, 0, 0)
            If json2 <> "" Then
                updateprogbar()
                Dim signedtxn As JObject = JsonConvert.DeserializeObject(json2)
                If LCase(signedtxn.Item("complete").ToString()) = "true" Then
                    updateprogbar()
                    Dim signedtx As String = signedtxn.Item("hex").ToString()
                    TxID = (New mlib).rpccalld("sendrawtransaction", 1, signedtx, 0, 0)
                    If TxID <> "" Then
                        updateprogbar()
                        MsgBox("Transaction sent. Tx ID: " & TxID)
                        Select Case tabMaster.SelectedIndex
                            Case 2
                                logTxID(TxID, "Simple Send", cboAddress.Text, txtSendtoBTCA.Text, cboCurrencySend.Text, txtAmount.Value)
                            Case 3
                                logTxID(TxID, "Sell Offer", cboAddress.Text, "", cboCurrencySell.Text, txtAmount2.Value)
                            Case 4
                                logTxID(TxID, "Purchase", cboAddress.Text, GVOfferstoSell.CurrentRow.Cells(6).Value, cboCurrencyBuy.Text, txtAmounttoPurchase.Value)
                            Case 5
                                logTxID(TxID, "Payment", cboAddress.Text, GVPurchasing.CurrentRow.Cells(6).Value, "BTC", txtAmounttoPay.Value)
                        End Select
                    Else
                        MsgBox("Error sending raw transaction.")
                    End If
                Else
                    MsgBox("Can't sign transaction (complete=false).")
                End If
            Else
                'Save Password temporarity.
                txtBTCWalletPP.Text = InputBox("Please enter your Wallet Passphrase and try to send again.  (You will have be requested re-enter everytime you restart.  Passphrase are not saved for your security.)", "Failed to sign transaction")
            End If
            clearprogbar()

        End If
        Return TxID
    End Function
    Function CheckValidCurrency(ByVal CurrencyName As String) As Boolean
        Dim IsValid As Boolean = False
        If Trim(CurrencyName) = "" Then
            MessageBox.Show("Please select a Currency.")
        Else
            If CurrencyName = "MSC" Then
                If Today >= CDate("03/15/2014") Then
                    IsValid = True
                Else
                    MessageBox.Show("MSC is not yet allowed in the Distributed Exchange")
                End If
            Else
                IsValid = True
            End If
        End If
        Return IsValid
    End Function


    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdSellOffer.Click
        If CheckValidCurrency(cboCurrencySell.Text) Then
            If lOK("Are you sure you want to sell " + txtAmount2.Value.ToString + " " + cboCurrencySell.Text + " for " + txtBTCSellPrice.Value.ToString + " btc") Then
                Dim CoinType As Integer = GetCurrencyInt(cboCurrencySell.Text)
                Dim SatTransactionFee As BigInteger = (New mlib).AmounttoSat(txtMinTransFee.Value)
                Dim SatAmount As BigInteger = (New mlib).AmounttoSat(txtAmount2.Value)
                Dim SatBTCAmount As BigInteger = (New mlib).AmounttoSat(txtBTCSellPrice.Value)

                Dim Action As Integer = 1
                If cmdSellOffer.Text = "Update Sell Offer" Then
                    Action = 2
                End If

                Dim RetVal As String = ""
                Dim rawtx As String = (New mlib).encodeselltx(cboAddress.Text, CoinType, _
                        SatAmount, SatBTCAmount, SatTransactionFee, Val(txtTimeLimit.Text), Action)
                If signandsendrawtransaction(rawtx) <> "" Then
                    txtAmount2.Value = 0
                    txtBTCUnitPrice.Value = 0
                    txtBTCSellPrice.Value = 0
                End If
            End If
        End If
    End Sub

    Function GetCurrencyInt(ByVal Currency As String) As Integer
        Dim CurrencyInt As Integer = 0
        Select Case Currency
            Case "MSC"
                CurrencyInt = 1
            Case "TMSC"
                CurrencyInt = 2
        End Select
        Return CurrencyInt
    End Function

    Structure BitcoinAddressAccount
        Dim strAddress As String
        Dim strAccount As String
        Dim strCusNum As String
        Public Overrides Function ToString() As String
            Return strCusNum
        End Function
    End Structure




    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        My.Settings.BitcoindExe = txtBitcoindexe.Text
        My.Settings.DataDir = txtDataDir.Text
        My.Settings.ConnectString = txtConnectString.Text
        My.Settings.Save()
        MsgBox("Settings updated.")

    End Sub

    Function IsMasterTabUpdated(ByVal tabno As Integer) As Boolean
        Return DateDiff(DateInterval.Minute, aUpdated(tabno), Now) < 10
    End Function
    Sub SettabMasterUpdated(ByVal tabno As Integer)
        aUpdated(tabno) = Now
    End Sub
    Sub GetTransactions()
        txtMessage.Text = ""
        If Not IsMasterTabUpdated(tabMaster.SelectedIndex) Then
            If tabMaster.SelectedIndex = 0 Or tabMaster.SelectedIndex = 1 Or tabMaster.SelectedIndex = 3 Or tabMaster.SelectedIndex = 4 Or tabMaster.SelectedIndex = 5 Then

                If Trim(cboAddress.Text).Length > 0 Then
                    SettabMasterUpdated(tabMaster.SelectedIndex)
                    Dim AddressID As Integer = (New mymastercoins).GetAddressID(cboAddress.Text)
                    Dim Sql As String = ""
                    Dim DS As DataSet
                    If AddressID > 0 Then
                        If tabMaster.SelectedIndex = 0 Then
                            updateDTAddress()
                        End If
                        If tabMaster.SelectedIndex = 1 Then
                            Getz_Transactions()
                        End If

                        If tabMaster.SelectedIndex = 3 Then
                            Getz_CurrencyforSale()
                        End If


                        If tabMaster.SelectedIndex = 4 Then
                            Getz_CurrencyforSaleAll()
                        End If


                        If tabMaster.SelectedIndex = 5 Then
                            Getz_PurchasingCurrencyforPayment()
                        End If
                    Else
                        If tabMaster.SelectedIndex = 1 Then
                            GVTrans.Visible = False
                            txtMessage.Text = "No transactions found"
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Sub ShowAvailable(ByVal AddressID, ByVal CurrencyID)
        txtTransactionBalance.Value = (New mymastercoins).GetAddressBalance(AddressID, CurrencyID)
        txtTransactionReserved.Value = (New mymastercoins).GetCurrencyForSale(AddressID, CurrencyID)
        txtTransactionAvailable.Value = txtTransactionBalance.Value - txtTransactionReserved.Value
        txtSellAvailable.Value = txtTransactionAvailable.Value
    End Sub

    Private Sub GVOfferstoSell_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GVOfferstoSell.RowEnter
        txtAmounttoPurchase.Text = Format(GVOfferstoSell.Rows(e.RowIndex).Cells(1).Value, "########.########")
        lblPurchaseOffer.Text = cboCurrencyBuy.Text + "  Total Btc " + Format(GVOfferstoSell.Rows(e.RowIndex).Cells(3).Value, "########.########") + " ( " + _
            Format(GVOfferstoSell.Rows(e.RowIndex).Cells(3).Value, "########.########") + " per " + cboCurrencyBuy.Text + " )  " + _
            "  Seller: " + GVOfferstoSell.Rows(e.RowIndex).Cells(6).Value
    End Sub

    Private Sub txtAmounttoPurchase_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not GVOfferstoSell.CurrentRow Is Nothing Then
            If Val(txtAmounttoPurchase.Text) < 0 Or Val(txtAmounttoPurchase.Text) > GVOfferstoSell.CurrentRow.Cells(1).Value Then
                lblPurchaseOffer.Text = "Amount must be less than or equal to " + Format(GVOfferstoSell.CurrentRow.Cells(1).Value, "########.########") + " " + cboCurrencyBuy.Text
            Else
                Dim UnitPrice As Double = GVOfferstoSell.CurrentRow.Cells(2).Value
                Dim TotalBTC As Double = Val(txtAmounttoPurchase.Text) * UnitPrice
                lblPurchaseOffer.Text = cboCurrencyBuy.Text + "  Total Btc " + Format(TotalBTC, "########.########") + " ( " + _
                    Format(UnitPrice, "########.########") + " per " + cboCurrencyBuy.Text + ") " + _
                    "  Seller: " + GVOfferstoSell.CurrentRow.Cells(6).Value
            End If
        End If
    End Sub



    Private Sub GVPurchasing_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GVPurchasing.RowEnter

        txtAmounttoPay.Text = Format(GVPurchasing.Rows(e.RowIndex).Cells(3).Value, "########.########")
        lblPay.Text = "btc to " + GVPurchasing.Rows(e.RowIndex).Cells(6).Value + ".  You get " + Format(GVPurchasing.Rows(e.RowIndex).Cells(1).Value, "########.########") + " " + cboCurrencyPayment.Text + " (" + _
            Format(GVPurchasing.Rows(e.RowIndex).Cells(2).Value, "########.########") + " per " + cboCurrencyPayment.Text + ")"
    End Sub
    Sub ResetProgbar(ByVal Num As Integer)
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = Num
        ProgressBar1.Visible = True
    End Sub
    Sub updateprogbar()
        ProgressBar1.Value += 1
    End Sub
    Sub clearprogbar()
        ProgressBar1.Visible = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If GVPurchasing.Visible = True Then
            Dim Recipient As String = GVPurchasing.CurrentRow.Cells(6).Value
            Dim txFee As String = GVPurchasing.CurrentRow.Cells(4).Value
            Dim CurrentBlockCount As Integer = (New Bitcoin).GetBlockCount()
            If CurrentBlockCount < GVPurchasing.CurrentRow.Cells(5).Value Then
                If lOK("Are you sure you want to send " + Str(txtAmounttoPay.Value) + " btc to " + Recipient) Then
                    Dim MinersFee As Double = 0.0002
                    If Val(txFee) > MinersFee Then
                        MinersFee = Val(txFee)
                    End If
                    Dim rawtx As String = (New mlib).encodepaymenttx(cboAddress.Text, Recipient, txtAmounttoPay.Value, MinersFee)
                    If signandsendrawtransaction(rawtx) <> "" Then
                        txtAmounttoPay.Value = 0
                    End If
                End If
            Else
                MsgBox("Time has expired.")
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim CurrencyID As Integer = GetCurrencyInt(cboCurrencyBuy.Text)
        If lOK("Are you sure you want to purchase " + txtAmounttoPurchase.Text + " " + cboCurrencyBuy.Text + "?") Then
            Dim Seller As String = GVOfferstoSell.CurrentRow.Cells(6).Value
            Dim rawtx As String = (New mlib).encodeaccepttx(cboAddress.Text, Seller, CurrencyID, (New mlib).AmounttoSat(Val(txtAmounttoPurchase.Text)))
            If signandsendrawtransaction(rawtx) <> "" Then
                txtAmounttoPurchase.Value = 0
            End If
        End If
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Interval = 600000
        txtMessage.Text = "Uploading Transactions "
        Call (New mymastercoins).GetExodusTransBlockExplorer()
        txtMessage.Text = "Processing Transactions "
        Call (New mymastercoins).ProcessTransactions()
        txtMessage.Text = ""
        GetTransactions()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As Object, ByVal e As EventArgs)
        Dim SQL As String = "select txid,count(txid) as ctr from z_exodus group by txid having count(txid)>1"
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
        If DS.Tables(0).DefaultView.Count > 0 Then
            For Each Row1 In DS.Tables(0).Rows
                SQL = "select * from z_Exodus where txid='" & Row1.item("TxID") & "'"
                Dim DS2 As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
                SQL = "delete from z_Exodus where ExodusID=" & DS2.Tables(0).Rows(0).Item("ExodusID")
                Call (New AWS.DB.ConnectDB).SQLdataset(SQL)
            Next
            txtMessage.Text = "Duplicates removed."
        Else
            txtMessage.Text = "No Duplicates found."
        End If

    End Sub


    Private Sub Button4_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim lFound As Boolean = False
        For Each Row1 In My.Settings.DTAddress.Rows
            If Row1("Address") = Trim(cboAddress.Text) Then
                lFound = True
                Exit For
            End If
        Next
        If Not lFound Then
            My.Settings.DTAddress.Rows.Add(cboAddress.Text, 0, 0, 0)
            My.Settings.Save()
            updateDTAddress()
            GetTransactions()
        End If
    End Sub


    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If lOK("Remove " + cboAddress.Text + "?") Then
            For Each Row1 In My.Settings.DTAddress.Rows
                If Row1("Address") = cboAddress.Text Then
                    My.Settings.DTAddress.Rows.Remove(Row1)
                    Exit For
                End If
            Next

            My.Settings.Save()
        End If
    End Sub
    Function lOK(ByVal Str As String) As Boolean
        Dim result As Integer = MessageBox.Show(Str, "Confirm", MessageBoxButtons.YesNo)
        Return result = 6
    End Function

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ResetProgbar(5)
        updateprogbar()
        Dim cLabel = Mid(txtPrivateKey.Text, 1, 10)
        Call (New mlib).rpccalld("walletpassphrase", 2, Trim(txtBTCWalletPP.Text), "30", 0)
        Dim json As String = (New mlib).rpccalld("importprivkey", 2, txtPrivateKey.Text, cLabel, 0)
        Call (New mlib).rpccalld("walletlock", 0, 0, 0, 0)
        Dim validater As JObject = JsonConvert.DeserializeObject(json)
        json = (New mlib).rpccalld("listreceivedbyaddress 0", 1, "true", 0, 0)
        If json.Length > 0 Then
            Dim obj As JArray = JsonConvert.DeserializeObject(json)
            For Each subitem In obj
                If Trim(subitem("account")) = Trim(cLabel) Then
                    Dim Address As String = subitem("address")
                    cboAddress.Items.Add(Address)
                    My.Settings.DTAddress.Rows.Add(Address, 0, 0, 0)
                    My.Settings.Save()
                    Exit For
                End If
            Next
            MsgBox("Private Keys imported.")
        Else
            MsgBox("Error importing Private Keys.")
        End If
        clearprogbar()
    End Sub
    Sub importaddressfromwallet()
        If lOK("Import all address in your Bitcoin Wallet?") Then
            '            Call (New mlib).rpccalld("walletlock", 0, 0, 0, 0)
            '           Call (New mlib).rpccalld("walletpassphrase", 2, Trim(txtBTCWalletPP.Text), "30", 0)
            Dim json As String = (New mlib).rpccalld("listreceivedbyaddress 0", 1, "true", 0, 0)
            If json.Length > 0 Then
                Dim obj As JArray = JsonConvert.DeserializeObject(json)
                For Each subitem In obj
                    Dim Address As String = Trim(subitem("address"))
                    Dim Found As Boolean = False
                    For Each Row1 In My.Settings.DTAddress.Rows
                        If Row1("Address") = Address Then
                            Found = True
                            Exit For
                        End If
                    Next
                    If Not Found Then
                        My.Settings.DTAddress.Rows.Add(Address, 0, 0, 0)
                    End If
                Next
                My.Settings.Save()
                updateDTAddress()
            Else
                MsgBox("Error importing Addresses.")
            End If
        End If

    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddress.Click
        importaddressfromwallet()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        updateDTAddress()
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabMaster.SelectedIndexChanged
        GetTransactions()
    End Sub

    Sub ResetaUpdated()
        For i = 1 To 7
            aUpdated(i) = DateAdd(DateInterval.Day, -1, Now)
        Next
    End Sub
    Private Sub cboAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAddress.SelectedIndexChanged
        GetTransactions()
        ResetaUpdated()
    End Sub

    Private Sub cmdCancelOffer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelOffer.Click
        Dim CoinType As Integer = GetCurrencyInt(cboCurrencySell.Text)
        If CoinType > 0 Then
            If lOK("Are you sure you want to cancel your " + cboCurrencySell.Text + " sell offer") Then
                Dim TxID As String = ""
                Dim rawtx As String = (New mlib).encodeselltx(cboAddress.Text, CoinType, 0, 0, 0, 0, 3)
                signandsendrawtransaction(rawtx)
            End If
        End If
    End Sub

    Private Sub Button11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim aLines As Array = Split(txtSendMany.Text + vbCr, vbCr)
        Dim TAmount As Double = 0
        Dim TAddress As Double = 0
        For Each Row1 In aLines
            Dim aRow As Array = Split(Row1, ",")
            If UBound(aRow) = 1 Then
                Dim Address As String = aRow(0)
                Dim Amount As Double = Val(aRow(1))
                If Amount > 0 Then
                    If IsValidAddress(Address) Then
                        TAmount += Amount
                        TAddress += 1
                    End If
                End If
            End If
        Next
        If lOK("Send " + TAmount.ToString + cboCurrencySend.Text + "  to " + TAddress.ToString + " Addresses?") Then
            Dim SenderID As Integer = (New mymastercoins).GetAddressID(cboAddress.Text)
            Dim CurrencyID As Integer = GetCurrencyInt(cboCurrencySend.Text)

            Dim Balance As Double = (New mymastercoins).GetAddressBalance(SenderID, CurrencyID)
            Dim CurrencyForSale As Double = (New mymastercoins).GetCurrencyForSale(SenderID, CurrencyID)

            If Balance - CurrencyForSale >= Val(TAmount) Then
                Dim Msg As String = ""
                For Each Row1 In aLines
                    Dim aRow As Array = Split(Row1, ",")
                    If UBound(aRow) = 1 Then
                        Dim Address As String = aRow(0)
                        Dim Amount As Double = Val(aRow(1))
                        If Amount > 0 Then
                            Dim CoinType As Integer = GetCurrencyInt(cboCurrencySend.Text)
                            Dim SatAmount As BigInteger = Amount * 100000000
                            Msg += SendCoin(cboAddress.Text, Address, CoinType, SatAmount) + vbCr
                        End If
                    End If
                Next
            Else
                MessageBox.Show("You can only send " + (Balance - CurrencyForSale).ToString + " " + cboCurrencySend.Text + ".  (Balance: " + Balance.ToString + " -  'For Sale':  " + CurrencyForSale.ToString + ")")
            End If
        End If
    End Sub
    Function IsValidAddress(ByVal Address As String) As Boolean
        Dim json As String = (New mlib).rpccalld("validateaddress", 1, Address, 0, 0)
        If json <> "" Then
            Dim validater As JObject = JsonConvert.DeserializeObject(json)
            Return LCase(validater.Item("isvalid").ToString) = "true"
        Else
            Return False
        End If
    End Function

    Sub RecalculateBTCSellPrice()
        txtBTCUnitPrice.Refresh()
        txtAmount2.Refresh()
        txtBTCSellPrice.Value = FormatNumber(Val(txtBTCUnitPrice.Value) * Val(txtAmount2.Value), 8)
        txtBTCSellPrice.Refresh()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount2.ValueChanged
        RecalculateBTCSellPrice()
    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBTCUnitPrice.ValueChanged
        RecalculateBTCSellPrice()
    End Sub


    Private Sub txtAmounttoPurchase_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmounttoPurchase.ValueChanged

        If Not GVOfferstoSell.CurrentRow Is Nothing Then
            If Val(txtAmounttoPurchase.Value) < 0 Or Val(txtAmounttoPurchase.Value) > GVOfferstoSell.CurrentRow.Cells(1).Value Then
                lblPurchaseOffer.Text = "Amount must be less than or equal to " + Format(GVOfferstoSell.CurrentRow.Cells(1).Value, "########.########") + " " + cboCurrencyBuy.Text
            Else
                Dim UnitPrice As Double = GVOfferstoSell.CurrentRow.Cells(2).Value
                Dim TotalBTC As Double = Val(txtAmounttoPurchase.Value) * UnitPrice
                lblPurchaseOffer.Text = cboCurrencyBuy.Text + "  Total Btc " + Format(TotalBTC, "########.########") + " ( " + _
                    Format(UnitPrice, "########.########") + " per " + cboCurrencyBuy.Text + ") " + _
                    "  Seller: " + GVOfferstoSell.CurrentRow.Cells(6).Value
            End If
        End If

    End Sub


    Private Sub NumericUpDown1_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmounttoPay.ValueChanged
        If Not GVPurchasing.CurrentRow Is Nothing Then
            If Val(txtAmounttoPay.Text) < 0 Or Val(txtAmounttoPay.Text) > GVPurchasing.CurrentRow.Cells(3).Value Then
                lblPay.Text = "Amount must be less than " + Format(GVPurchasing.CurrentRow.Cells(3).Value, "########.########")
            Else
                Dim UnitPrice As Double = GVPurchasing.CurrentRow.Cells(2).Value
                Dim Coins As Double = Val(txtAmounttoPay.Text) / UnitPrice
                lblPay.Text = "btc to " + GVPurchasing.CurrentRow.Cells(6).Value + ".  You get " + Format(Coins, "########.########") + " " + cboCurrencyPayment.Text + " (" + _
                    Format(UnitPrice, "########.########") + " per " + cboCurrencyPayment.Text + " )"
            End If
        End If

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If lOK("Clear the Sent Log?") Then
            My.Settings.DTLog = GetLogTable()
            My.Settings.Save()
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim lFound As Boolean = False
        For Each Row1 In My.Settings.DTAddressBook.Rows
            If Row1("Address") = Trim(txtSendtoBTCA.Text) Then
                lFound = True
                Exit For
            End If
        Next

        If Not lFound Then
            Dim Address As String = Trim(txtSendtoBTCA.Text)
            My.Settings.DTAddressBook.Rows.Add(Address, Address)
            My.Settings.Save()
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If lOK("Remove " + txtSendtoBTCA.Text + "?") Then
            For Each Row1 In My.Settings.DTAddressBook.Rows
                If Row1("Address") = txtSendtoBTCA.Text Then
                    My.Settings.DTAddress.Rows.Remove(Row1)
                    My.Settings.Save()
                    Exit For
                End If
            Next

        End If
    End Sub


    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        txtMessage.Text = ""
        clearprogbar()
        importaddressfromwallet()
    End Sub

    Private Sub cboCurrencyTrans_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCurrencyTrans.SelectedIndexChanged
        Getz_Transactions()
    End Sub

    Sub Getz_Transactions()
        Dim CurrencyID As Integer = GetCurrencyInt(cboCurrencyTrans.Text)
        If CurrencyID > 0 Then
            ResetProgbar(3)
            updateprogbar()
            Dim AddressID As Integer = (New mymastercoins).GetAddressID(cboAddress.Text)
            Dim Sql As String = ""
            Dim DS As DataSet
            With GVTrans
                .Columns("colIn").DefaultCellStyle.Format = "########.########"
                .Columns("colOut").DefaultCellStyle.Format = "########.########"
                .Columns("colTotalPrice").DefaultCellStyle.Format = "########.########"
                .Columns("colTotalPrice").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .RowHeadersVisible = False
                .AutoGenerateColumns = False
            End With

            updateprogbar()
            Sql = "SELECT *" + _
                    " FROM z_Transactions LEFT OUTER JOIN z_Exodus ON z_Transactions.ExodusID = z_Exodus.ExodusID " & _
                    " where AddressID=" + AddressID.ToString + " and z_Transactions.CurrencyID=" + CurrencyID.ToString + _
                    " order by z_Transactions.dtrans"
            DS = (New AWS.DB.ConnectDB).SQLdataset(Sql)
            txtTransactionBalance.Value = 0
            GVTrans.DataSource = DS.Tables(0).DefaultView
            GVTrans.Visible = True
            updateprogbar()
            ShowAvailable(AddressID, CurrencyID)
            clearprogbar()
        End If
    End Sub

    Private Sub cboCurrencySell_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCurrencySell.SelectedIndexChanged
        Getz_CurrencyforSale()
    End Sub
    Sub Getz_CurrencyforSale()

        Dim CurrencyID As Integer = GetCurrencyInt(cboCurrencySell.Text)
        If CurrencyID > 0 Then
            ResetProgbar(3)
            Dim AddressID As Integer = (New mymastercoins).GetAddressID(cboAddress.Text)
            Dim Sql As String = ""
            Dim DS As DataSet
            Sql = "SELECT z_CurrencyforSale.*,00000000.00000000 as UnitPrice" + _
    " FROM z_CurrencyforSale " + _
    " WHERE  CurrencyID = " + CurrencyID.ToString + " and IsNewOffer=1 " + _
    " and AddressID=" + AddressID.ToString + " and CurrencyID=" + CurrencyID.ToString + _
    " order by dtrans desc"
            Dim AmountforSale As Double = 0
            Dim AmountBTCDesired As Double = 0
            Dim TimeLimit As Integer = 0
            Dim MinTransFee As Double = 0
            updateprogbar()
            DS = (New AWS.DB.ConnectDB).SQLdataset(Sql)
            If DS.Tables(0).DefaultView.Count > 0 Then
                cmdSellOffer.Text = "Update Sell Offer"
                With DS.Tables(0).Rows(0)
                    AmountforSale = .Item("Available")
                    AmountBTCDesired = .Item("BTCDesired")
                    TimeLimit = .Item("TimeLimit")
                    MinTransFee = .Item("TransFee")
                    If Not IsDBNull(.Item("Action")) Then
                        If .Item("Action") = 3 Or AmountforSale = 0 Then
                            'Last Action is Cancel
                            cmdSellOffer.Text = "New Sell Offer"
                        End If
                    Else
                        cmdSellOffer.Text = "New Sell Offer"
                    End If
                End With
                cmdCancelOffer.Enabled = True
            Else
                cmdSellOffer.Text = "New Sell Offer"
                cmdCancelOffer.Enabled = False
            End If
            txtCSOAmountforSale.Text = AmountforSale
            txtCSOAmountBTCDesired.Text = FormatNumber(AmountBTCDesired, 8)
            txtCSOTimeLimit.Text = TimeLimit
            txtCSOMinTransFee.Text = FormatNumber(MinTransFee, 8)
            txtCSOUnitPrice.Text = "0"
            If AmountBTCDesired > 0 And AmountforSale > 0 Then
                txtCSOUnitPrice.Text = FormatNumber((AmountBTCDesired / AmountforSale), 8)
            End If
            updateprogbar()
            ShowAvailable(AddressID, CurrencyID)
            clearprogbar()
        End If
    End Sub

    Private Sub cboCurrencyBuy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCurrencyBuy.SelectedIndexChanged
        Getz_CurrencyforSaleAll()
    End Sub
    Sub Getz_CurrencyforSaleAll()
        Dim CurrencyID As Integer = GetCurrencyInt(cboCurrencyBuy.Text)
        If CurrencyID > 0 Then
            ResetProgbar(3)
            updateprogbar()
            Dim AddressID As Integer = (New mymastercoins).GetAddressID(cboAddress.Text)
            Dim Sql As String = ""
            Dim DS As DataSet
            With GVOfferstoSell
                .Columns("colSCoinstoSell").DefaultCellStyle.Format = "########.########"
                .Columns("colSUnitPrice").DefaultCellStyle.Format = "########.########"
                .Columns("colSBTCDesired").DefaultCellStyle.Format = "########.########"
                .Columns("colSTimeLimit").DefaultCellStyle.Format = "####"
                .Columns("colSTransFee").DefaultCellStyle.Format = "########.########"
                .AutoGenerateColumns = False
                .Columns(0).DataPropertyName = "dTrans"
                .Columns(1).DataPropertyName = "Available"
                .Columns(2).DataPropertyName = "UnitPrice"
                .Columns(3).DataPropertyName = "TotalPrice"
                .Columns(4).DataPropertyName = "TimeLimit"
                .Columns(4).Width = 50
                .Columns(5).DataPropertyName = "TransFee"
                .Columns(5).Width = 75
                .Columns(6).DataPropertyName = "Seller"
                .RowHeadersVisible = False
            End With

            Sql = "SELECT z_CurrencyforSale.*, BTCDesired/AmountforSale as UnitPrice, (BTCDesired/AmountforSale)*Available as TotalPrice,space(100) as Seller" + _
    " FROM z_CurrencyforSale " + _
    " WHERE  CurrencyID = " + CurrencyID.ToString + " and IsNewOffer=1 " + _
    "  and Available>0" + _
    " order by UnitPrice"
            updateprogbar()
            DS = (New AWS.DB.ConnectDB).SQLdataset(Sql)
            If DS.Tables(0).DefaultView.Count > 0 Then
                For Each DT In DS.Tables(0).Rows
                    DT.item("Seller") = (New mymastercoins).GetAddress(DT.item("AddressID"))
                Next
                GVOfferstoSell.DataSource = DS.Tables(0).DefaultView
                GVOfferstoSell.Visible = True
            Else
                GVOfferstoSell.Visible = False
            End If
            clearprogbar()
        End If
    End Sub

    Private Sub cboCurrencyPayment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCurrencyPayment.SelectedIndexChanged
        Getz_PurchasingCurrencyforPayment()
    End Sub
    Sub Getz_PurchasingCurrencyforPayment()
        Dim CurrencyID As Integer = GetCurrencyInt(cboCurrencyPayment.Text)
        If CurrencyID > 0 Then
            ResetProgbar(3)
            updateprogbar()
            Dim AddressID As Integer = (New mymastercoins).GetAddressID(cboAddress.Text)
            Dim Sql As String = ""
            Dim DS As DataSet
            Dim CurrentBlockNo As Integer = (New Bitcoin).GetBlockCount()
            Sql = "SELECT *, space(40) as Seller, TotalBTC-Paid as UnPaid," + CurrentBlockNo.ToString + " as CurrentBlockNo" + _
            " FROM z_PurchasingCurrency " + _
            " WHERE  CurrencyID = " + CurrencyID.ToString + _
            " and MaxBlockNo>=" + CurrentBlockNo.ToString + _
            " and AddressID=" + AddressID.ToString + _
            " order by dtrans desc"


            With GVPurchasing
                .Columns(1).DefaultCellStyle.Format = "########.########"
                .Columns(2).DefaultCellStyle.Format = "########.########"
                .Columns(3).DefaultCellStyle.Format = "########.########"
                .Columns(4).DefaultCellStyle.Format = "##.########"
                .Columns(5).DefaultCellStyle.Format = "####"
                .Columns(7).DefaultCellStyle.Format = "########.########"
                .Columns(8).DefaultCellStyle.Format = "########.########"
                .AutoGenerateColumns = False
                .Columns(0).DataPropertyName = "dTrans"
                .Columns(1).DataPropertyName = "PurchasedAmount"
                .Columns(2).DataPropertyName = "UnitPrice"
                .Columns(3).DataPropertyName = "TotalBTC"
                .Columns(4).DataPropertyName = "TransFee"
                .Columns(5).DataPropertyName = "MaxBlockNo"
                .Columns(6).DataPropertyName = "Seller"
                .Columns(7).DataPropertyName = "Paid"
                .Columns(8).DataPropertyName = "UnPaid"
                .RowHeadersVisible = False
            End With

            updateprogbar()
            DS = (New AWS.DB.ConnectDB).SQLdataset(Sql)
            If DS.Tables(0).DefaultView.Count > 0 Then
                For Each DT In DS.Tables(0).Rows
                    DT.item("Seller") = (New mymastercoins).GetAddress(DT.item("SellerID"))
                Next
                GVPurchasing.DataSource = DS.Tables(0).DefaultView
                GVPurchasing.Visible = True
            Else
                GVPurchasing.Visible = False
            End If
            lblCurrentBlockTime.Text = "Current Block Time: " + (New Bitcoin).GetBlockCount.ToString
        End If
    End Sub

    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Timer1.Enabled = False
        txtMessage.Text = "Uploading Transactions "
        Call (New mymastercoins).GetExodusTransBlockExplorer()
        txtMessage.Text = "Processing Transactions "
        Call (New mymastercoins).ProcessTransactions()

        ResetProgbar(1)
        txtMessage.Text = "Loading " + Trim(cboVerify.Text)
        Dim obj As New JArray
        Dim json As String = (New Bitcoin).getjson(Trim(cboVerify.Text))
        obj = JsonConvert.DeserializeObject(json)
        Dim i As Integer
        Dim s As String = ""
        ResetProgbar(obj.Count)
        For i = 0 To obj.Count - 1
            updateprogbar()
            Dim Address As String = Trim(obj.Item(i).Item("address").ToString)
            Dim Balance As Double = 0
            If obj.Item(i).Item("balance").ToString <> "" Then
                Balance = Math.Round(CDbl(obj.Item(i).Item("balance")), 8)
            End If
            'Exclude Exodus Address from Checking
            If Address <> "1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P" Then
                Dim AddressID As Integer = (New mymastercoins).GetAddressID(Address)
                If AddressID > 0 Then
                    Dim OffertoSell As Double = (New mymastercoins).GetCurrencyForSale(AddressID, 1)
                    Dim BalanceMM As Double = Math.Round((New mymastercoins).GetAddressBalance(AddressID, 1) - OffertoSell, 8)
                    If BalanceMM <> Balance Then
                        s += BalanceMM.ToString + " " + Address + " " + Balance.ToString + vbCrLf
                    End If
                End If
            End If
        Next
        clearprogbar()
        If s <> "" Then
            txtMessage.Text = "Wallet is not verified."
            MsgBox(s)
        Else
            txtMessage.Text = "Wallet is verified and updated. " + Now().ToString
        End If
        Timer1.Enabled = True
    End Sub
End Class


