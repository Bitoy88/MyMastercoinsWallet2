Imports System.Numerics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Microsoft.VisualBasic
Imports System.Data

Public Class mymastercoinsold2
    Dim Ctr As Integer = 0
    Dim DecodeDataString As String = ""
    Function GetandProcessTx(ByVal LastProcess As String, Optional ByVal RefreshMin As Integer = 10) As String
        If LastProcess = "" Then

            Dim SQL2 As String = "SELECT top 1 * from z_Exodus where Isprocessed=1 order by dTrans desc "
            Dim DS2 As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL2)
            If DS2.Tables(0).DefaultView.Count > 0 Then
                LastProcess = DS2.Tables(0).Rows(0).Item("dTrans").ToString
            Else
                LastProcess = Now().ToString
            End If
        End If

        Dim LastProcessed As DateTime = DateTime.Parse(LastProcess)
        'Process again if more than 10 minutes
        If DateDiff(DateInterval.Minute, LastProcessed, Now) > RefreshMin Then
            LastProcess = Now().ToString
            Call (New mymastercoinsold2).GetExodusTransBlockExplorer()
            Call (New mymastercoinsold2).ProcessTransactions()
            '            lblMessage.Text = "Processed " & Application("LastProcessed")
        Else
            '            lblMessage.Text = "Last Processed " & Application("LastProcessed")
        End If
        Return LastProcess
    End Function
    Function DecodeHexString(ByVal s As String, ByVal NBytes As Integer) As String
        Dim HexString As String = DecodeDataString.Substring(Ctr, NBytes)

        Dim n As BigInteger = ("&H" & HexString)
        '        TextBox1.Text += s + n.ToString + " (" + HexString + ")" + vbCrLf
        Ctr += NBytes
        Return n.ToString
    End Function

    Sub ProcessTransactions()
        Dim SQL As String = "SELECT  * from z_Exodus where IsProcessed=0 order by blocknumber,ExodusID"
        Dim MSCBlockStart As Integer = 999999999
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
        Form1.ResetProgbar(DS.Tables(0).DefaultView.Count)
        For Each Row1 In DS.Tables(0).Rows
            Form1.updateprogbar()
            Dim ExodusID As Integer = Row1.Item("ExodusID")
            Dim dTrans As DateTime = Row1.Item("dTrans")
            Dim BlockNumber As Integer = Row1.Item("BlockNumber")
            Dim SenderAddress As String = ""
            Dim MSCData1 As String = ""
            Dim MSCData2 As String = ""
            Dim DataString2 As String = ""
            Dim DataString1 As String = ""
            Dim Receipient As String = ""
            Dim IsMultiSig As Boolean = False
            Dim CurrencyID As String = "0"


            'Use Block Explorer to get Class B multisig  
            Dim txhash As String = Trim(Row1.Item("TxID"))
            Dim json As String = ""
            Dim ClassAReceipient As String = ""
            Dim ClassAData As String = ""
            Dim ThirdAddress As String = ""


            If Not IsDBNull(Row1("berawtx")) Then
                json = Trim(Row1("berawtx"))
                Dim obj As New JObject
                obj = JsonConvert.DeserializeObject(json)

                Dim ccArray As JArray = obj.Item("out")
                Dim scriptPubKey As String = ""
                For Each subitem In ccArray
                    If subitem("scriptPubKey").ToString.Contains("OP_CHECKMULTISIG") Then
                        IsMultiSig = True
                        scriptPubKey = subitem("scriptPubKey").ToString
                        Exit For
                    End If
                Next
                If IsMultiSig Then
                    Dim aTemp As Array = Split(scriptPubKey)
                    If aTemp(2).length = 66 Then
                        MSCData1 = aTemp(2)
                    End If
                    If aTemp(3).length = 66 Then
                        MSCData2 = aTemp(3)
                    End If
                    If aTemp.Length > 4 Then
                        If aTemp(4).length = 66 Then
                            Dim MCData3 As String = aTemp(4)
                        End If
                    End If
                End If

                '****
                If Not IsDBNull(Row1.Item("beReferenceAddress")) Then
                    SenderAddress = Trim(Row1.Item("beReferenceAddress"))
                End If
                GetbeReceipient(Row1.Item("beOut"), ClassAReceipient, ClassAData, ThirdAddress)
                Receipient = ClassAReceipient
            Else
                If Trim(Row1("bcrawtx")) <> "" Then
                    json = Trim(Row1("bcrawtx"))
                    Dim obj As New JObject
                    obj = JsonConvert.DeserializeObject(json)


                    SenderAddress = GetbcSender(obj.Item("inputs").ToString)

                    Dim ccArray As JArray = obj.Item("out")
                    Dim scriptPubKey As String = ""
                    For Each subitem In ccArray
                        If Not IsNothing(subitem("script")) Then
                            scriptPubKey = subitem("script").ToString
                            If Right(scriptPubKey, 4) = "52ae" Then
                                IsMultiSig = True
                                Dim First As Integer = scriptPubKey.IndexOf("2102")
                                If First > 0 Then
                                    MSCData1 = Mid(scriptPubKey, First + 3, 66)
                                    Dim ScriptSig2 As String = Mid(scriptPubKey, First + 3 + 66)
                                    If ScriptSig2.Length > 66 Then
                                        Dim Sec As Integer = ScriptSig2.IndexOf("2102")
                                        If Sec >= 0 Then
                                            Dim temp As String = Mid(ScriptSig2, Sec + 3, 66)
                                            MSCData2 = MSCData1
                                            MSCData1 = temp
                                        End If
                                    End If
                                End If
                                Exit For
                            End If
                            If Right(scriptPubKey, 4) = "53ae" Then
                                IsMultiSig = True
                                Dim First As Integer = scriptPubKey.IndexOf("2102")
                                If First > 0 Then
                                    MSCData1 = Mid(scriptPubKey, First + 3, 66)
                                    Dim ScriptSig2 As String = Mid(scriptPubKey, First + 3 + 66)
                                    If ScriptSig2.Length > 66 Then
                                        Dim Sec As Integer = ScriptSig2.IndexOf("2102")
                                        If Sec >= 0 Then
                                            MSCData2 = Mid(ScriptSig2, Sec + 3, 66)
                                        Else
                                            Sec = ScriptSig2.IndexOf("2202")
                                            If Sec >= 0 Then
                                                MSCData2 = Mid(ScriptSig2, Sec + 3, 66)
                                            End If
                                        End If

                                        Dim ScriptSig3 As String = Mid(ScriptSig2, Sec + 3 + 66)
                                        If ScriptSig3.Length > 66 Then
                                            Dim Third = ScriptSig3.IndexOf("2102")
                                            If Third >= 0 Then
                                                MSCData1 = MSCData2
                                                MSCData2 = Mid(ScriptSig3, Third + 3, 66)
                                            End If
                                        End If

                                    End If
                                End If
                                Exit For
                            End If
                        End If
                    Next

                    'Find Recipient
                    Dim Value As Double = 99999999
                    For Each subitem In ccArray
                        Dim Address As String = Trim(subitem("addr"))
                        If Address = "1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P" And Val(subitem("value")) < Value Then
                            Value = Val(subitem("value"))
                        End If
                    Next
                    Dim PossibleData As String = ""
                    For Each subitem In ccArray
                        Dim Address As String = Trim(subitem("addr"))
                        If Address <> "" And Address <> "1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P" Then
                            If Val(subitem("value")) = Value Then
                                'Get First 3 Address
                                ThirdAddress = ClassAData
                                ClassAData = ClassAReceipient
                                ClassAReceipient = Address
                            Else
                                If Address <> SenderAddress And PossibleData = "" Then
                                    PossibleData = Address
                                End If
                            End If
                        End If
                    Next
                    If ClassAData = "" Then
                        ClassAData = PossibleData
                    End If
                    Receipient = ClassAReceipient

                End If
            End If

            If False Then
                If Receipient = "" And IsMultiSig Then
                    ' For cases where there is only an exodus address and one address in a multisig that doesn't have the same value as the exodus address
                    ' ie 3ac51761bfcbc12fa110b3da5bf540573af0c3fa6732723d709061e169b68784 
                    If ClassAData <> "" Then
                        Receipient = ClassAData
                        ClassAData = ""
                    End If
                End If
            End If


            Dim IsValidExodus As Boolean = False
            Dim Remarks As String = ""
            If SenderAddress <> "" Then
                If Not ClassAReceipient = "" And Not ClassAData = "" Then
                    If Not IsMultiSig Then
                        'Class A transaction
                        Dim TestR As String = (New Bitcoin).DecodeBTCAddresstoHexString(ClassAReceipient)
                        Dim TestD As String = (New Bitcoin).DecodeBTCAddresstoHexString(ClassAData)
                        Dim SeqR As Integer = CInt("&H" + TestR.Substring(2, 2))
                        Dim SeqD As Integer = CInt("&H" + TestD.Substring(2, 2))
                        Dim IsValidClassA As Boolean = False

                        'Peek and Decode
                        If ThirdAddress = "" Then
                            If IsDataCode(ClassAReceipient) Then
                                Dim Temp As String = ClassAReceipient
                                ClassAReceipient = ClassAData
                                ClassAData = Temp
                            End If
                            IsValidClassA = True
                        Else
                            'There are 3 address in the out choose which one is the data and receipient
                            Dim Addr1 As String = ClassAReceipient
                            Dim Addr2 As String = ClassAData
                            Dim Addr3 As String = ThirdAddress
                            Dim Seq1 As Integer = CInt("&H" + (New Bitcoin).DecodeBTCAddresstoHexString(ClassAReceipient).Substring(2, 2))
                            Dim Seq2 As Integer = CInt("&H" + (New Bitcoin).DecodeBTCAddresstoHexString(ClassAData).Substring(2, 2))
                            Dim Seq3 As Integer = CInt("&H" + (New Bitcoin).DecodeBTCAddresstoHexString(ThirdAddress).Substring(2, 2))
                            Dim ReceipentSeq As Integer = 0

                            If IsDataCode(Addr1) Then
                                ClassAData = Addr1
                                ReceipentSeq = Seq1 + 1
                                If ReceipentSeq > 255 Then
                                    ReceipentSeq = 0
                                End If
                                If Seq2 = ReceipentSeq Then
                                    ClassAReceipient = Addr2
                                Else
                                    ClassAReceipient = Addr3
                                End If
                            Else
                                If IsDataCode(Addr2) Then
                                    ClassAData = Addr2
                                    ReceipentSeq = Seq2 + 1
                                    If ReceipentSeq > 255 Then
                                        ReceipentSeq = 0
                                    End If
                                    If Seq1 = ReceipentSeq Then
                                        ClassAReceipient = Addr1
                                    Else
                                        ClassAReceipient = Addr3
                                    End If

                                Else
                                    ClassAData = Addr3
                                    ReceipentSeq = Seq3 + 1
                                    If ReceipentSeq > 255 Then
                                        ReceipentSeq = 0
                                    End If
                                    If Seq2 = ReceipentSeq Then
                                        ClassAReceipient = Addr2
                                    Else
                                        ClassAReceipient = Addr1
                                    End If
                                End If
                            End If
                            IsValidClassA = True
                        End If

                        If False Then
                            If Math.Abs(SeqR - SeqD) = 1 Or Math.Abs(SeqR - SeqD) = 255 Then
                                If (SeqR < SeqD) Or (SeqR = 255 And SeqD = 0) Then
                                    Dim Temp As String = ClassAReceipient
                                    ClassAReceipient = ClassAData
                                    ClassAData = Temp
                                End If
                                IsValidClassA = True
                            Else
                            End If

                        End If

                        If IsValidClassA Then
                            Receipient = ClassAReceipient
                            MSCData1 = (New Bitcoin).DecodeBTCAddresstoHexString(ClassAData).Substring(2)

                            If MSCData1.Length > 0 Then
                                DataString1 = "01" + MSCData1.Substring(2)
                            End If
                        Else
                            DataString1 = ""
                        End If
                    End If
                End If

                If IsMultiSig Then
                    If MSCData1.Length > 2 Then
                        Dim Obsfu As String = MSCData1.Substring(2)
                        Obsfu = Obsfu.Remove(Obsfu.Length - 2)
                        Dim hashString1 As String = (New Bitcoin).GetHash(SenderAddress)
                        DataString1 = (New Bitcoin).ObsXor(Obsfu, hashString1)

                        If MSCData2.Length > 0 Then
                            Dim Obsfu2 As String = MSCData2.Substring(2)
                            Obsfu2 = Obsfu2.Remove(Obsfu2.Length - 2)
                            '                    TextBox1.Text += "Data to Parse: " + MSCData2 + vbCrLf
                            '                   TextBox1.Text += "OBFUSCATED MASTERCOIN PACKET: " + Obsfu + vbCrLf
                            Dim HashString2 As String = (New Bitcoin).GetHash(hashString1.ToUpper)
                            '                  TextBox1.Text += "SHA256 HASH: " + HashString2 + vbCrLf
                            DataString2 = (New Bitcoin).ObsXor(Obsfu2, HashString2)
                            '                 TextBox1.Text += "CLEARTEXT MASTERCOIN PACKET: " + DataString2 + vbCrLf + vbCrLf
                        End If

                    End If
                End If
                Dim IsPayment As Boolean = True
                If DataString1.Length > 0 Then
                    Ctr = 0
                    DecodeDataString = DataString1
                    If DataString2.Length > 2 Then
                        DecodeDataString += DataString2.Substring(2) ' Remove first 2 char in 2nd packet
                    End If
                    '            TextBox1.Text += "FULL CLEAR TEXT: " + DecodeDataString + vbCrLf
                    Dim Sequence As Integer = CInt(DecodeHexString("01: ", 2))
                    If Sequence = 1 Then
                        IsPayment = False
                        '2 Bytes Version 2 Bytes Type
                        Dim Version As String = Val(DecodeHexString("Trans Version: ", 4))
                        Dim TransType As String = DecodeHexString("Trans Type: ", 4)
                        Select Case Val(TransType)
                            Case 0
                                'Simple Send
                                '                        lblTitle.Text = "Simple Send"
                                CurrencyID = DecodeHexString("Currency ID: ", 8)
                                Dim AmounttoSend As String = DecodeHexString("Amount to Send: ", 16)
                                Dim Currency As String = (New Bitcoin).GetCurrency(CurrencyID)

                                'Check if User has enough funds
                                Dim SenderID As Integer = GetAddressID(SenderAddress)
                                Dim AmountToSend2 As Double = Val(AmounttoSend) / 100000000
                                Dim Balance As Double = GetAddressBalance(SenderID, CurrencyID)
                                Dim Reserved As Double = (New mymastercoinsold2).GetCurrencyForSale(SenderID, CurrencyID)
                                '                            Dim Pending As Double = (New mymastercoins).GetSellerCoinsPendingPayment(SenderID, BlockNumber)
                                Dim Available As Double = (Balance - Reserved)

                                If Available >= AmountToSend2 Then
                                    Dim ReceipientID As Integer = AddAddress(Receipient)
                                    If ReceipientID <> 0 Then
                                        If SenderID <> ReceipientID Then
                                            Call AddTransaction(dTrans, SenderID, ExodusID, Receipient, 0, AmountToSend2, CurrencyID)
                                            Call AddTransaction(dTrans, ReceipientID, ExodusID, SenderAddress, AmountToSend2, 0, CurrencyID)
                                        Else
                                            Call AddTransaction(dTrans, SenderID, ExodusID, Receipient, AmountToSend2, AmountToSend2, CurrencyID)
                                        End If
                                        IsValidExodus = True
                                        SQL = "update z_Exodus set MSCSenderAddress='" & SenderAddress & "',MSCReceiverAddress='" & Receipient & "' where ExodusID=" & ExodusID
                                        Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                                        Remarks = "Simple Send: " + AmountToSend2.ToString + " " + Currency
                                    Else
                                        Remarks = "No Recipient. "
                                    End If

                                Else
                                    Remarks = "Sender " + SenderAddress + " doesn't have " + AmountToSend2.ToString + " " + Currency + ". Bal:" + Balance.ToString + " Res:" + Reserved.ToString + " Avail:" + Available.ToString
                                End If
                            Case 20
                                'Selling Coins for Bitcoins
                                CurrencyID = DecodeHexString("Currency ID: ", 8)
                                If CurrencyID = 1 And dTrans <= Convert.ToDateTime("2014/03/15") Then
                                    '*** 
                                    Remarks = "MSC Trading has not yet started. (starts in 3/15/2014)"
                                Else
                                    Dim Currency As String = (New Bitcoin).GetCurrency(CurrencyID)
                                    Dim AmountForSale As String = DecodeHexString("Amount for Sale: ", 16)
                                    AmountForSale = (CDbl(AmountForSale) / 100000000).ToString
                                    Dim BTCDesired As String = DecodeHexString("BTC Desired:  ", 16)
                                    BTCDesired = (CDbl(BTCDesired) / 100000000).ToString
                                    Dim TimeLimit As String = DecodeHexString("Time Limit:  ", 2)
                                    Dim TransFee As String = DecodeHexString("Min Trans Fee:  ", 16)
                                    TransFee = (CDbl(TransFee) / 100000000).ToString
                                    Dim Action As String = DecodeHexString("Action:  ", 2)

                                    Dim AddressID As Integer = GetAddressID(SenderAddress)
                                    SQL = "select ExodusID,Action,Available from z_CurrencyforSale where IsNewOffer=1 and AddressID=" & AddressID
                                    Dim DSPrevious As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
                                    Dim IsValidSellOffer As Boolean = False
                                    Dim PrevAction As Integer = 0
                                    Dim PrevAvailable As Double = 0
                                    If DSPrevious.Tables(0).DefaultView.Count > 0 Then
                                        If Not IsDBNull(DSPrevious.Tables(0).Rows(0).Item("Action")) Then
                                            PrevAction = DSPrevious.Tables(0).Rows(0).Item("Action")
                                        End If
                                        PrevAvailable = DSPrevious.Tables(0).Rows(0).Item("Available")
                                    End If
                                    Dim PendingCoins As Double = (New mymastercoinsold2).GetSellerCoinsPendingPayment(AddressID, BlockNumber, CurrencyID)

                                    Select Case Version
                                        Case 0
                                            If AmountForSale = 0 Then
                                                'Cancel
                                                Action = 3
                                            Else
                                                If PrevAction = 1 Or PrevAction = 2 Then
                                                    Action = 2
                                                Else
                                                    Action = 1
                                                End If
                                            End If
                                            IsValidSellOffer = True
                                        Case 1
                                            Select Case Val(Action)
                                                Case 1
                                                    'New
                                                    If PrevAvailable = 0 Or PrevAction = 3 Then
                                                        If PendingCoins > 0 Then
                                                            Remarks = "Invalid. Seller has " + PendingCoins.ToString + " " + Currency + " pending payment "
                                                        Else
                                                            IsValidSellOffer = True
                                                        End If
                                                    Else
                                                        Remarks = "Can't post a new Sell Offer is there is an existing Sell Offer."
                                                    End If
                                                Case 2
                                                    'Update
                                                    If PrevAction = 1 Or PrevAction = 2 Then
                                                        IsValidSellOffer = True
                                                    Else
                                                        Remarks = "The Previous Sell Offer is a Cancel.  You can only update only if the previous Sell Offer is New or an Update."
                                                    End If
                                                Case 3
                                                    IsValidSellOffer = True
                                                    AmountForSale = "0"
                                            End Select
                                    End Select
                                    If BTCDesired = 0 And Val(Action) <> 0 Then
                                        IsValidSellOffer = False
                                        Remarks = "BTC Desired is 0"
                                    End If
                                    If IsValidSellOffer Then
                                        Dim Balance As Double = GetAddressBalance(AddressID, CurrencyID)
                                        Dim Pending As Double = GetSellerCoinsPendingPayment(AddressID, BlockNumber, CurrencyID)
                                        If AmountForSale > Balance - Pending Then
                                            AmountForSale = Balance - Pending
                                        End If
                                        SQL = "update z_CurrencyforSale set IsNewOffer=0 where IsNewOffer=1 and AddressID=" & AddressID
                                        Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                                        SQL = "insert into z_CurrencyforSale (Action,IsNewOffer,Available,dTrans,AddressID,CurrencyID,AmountForSale,BTCDesired,TimeLimit,TransFee,ExodusID) values ( " + _
                                        Action + ",1," + AmountForSale.ToString + ",'" + dTrans.ToString + "'," + AddressID.ToString + "," + CurrencyID.ToString + "," + AmountForSale.ToString + "," + BTCDesired.ToString + "," + TimeLimit.ToString + "," + TransFee.ToString + "," + ExodusID.ToString + ")"
                                        Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                                        IsValidExodus = True
                                        Select Case Val(Action)
                                            Case 3
                                                Remarks = "Cancel Offer to Sell"
                                            Case 2
                                                Remarks = "Update Offer to Sell: " + AmountForSale.ToString + " " + Currency + " for " + BTCDesired.ToString + " btc"
                                            Case 1
                                                Remarks = "New Offer to Sell: " + AmountForSale.ToString + " " + Currency + " for " + BTCDesired.ToString + " btc"
                                        End Select
                                    End If
                                End If
                            Case 22
                                'Purchasing a Currency Offered for Sale
                                CurrencyID = DecodeHexString("Currency ID: ", 8)
                                Dim Currency As String = (New Bitcoin).GetCurrency(CurrencyID)

                                Dim AmountPurchasing As String = DecodeHexString("Amount Purchasing: ", 16)
                                AmountPurchasing = (Val(AmountPurchasing) / 100000000).ToString

                                Dim SellerID As Integer = AddAddress(Receipient)
                                SQL = "select * from z_CurrencyforSale where AddressID=" & SellerID.ToString & " and CurrencyID= " & CurrencyID & " and IsNewOffer=1 order by dTrans"
                                Dim DSAvailable As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
                                Dim AmountPurchased As Double = CDbl(AmountPurchasing)
                                If DSAvailable.Tables(0).DefaultView.Count > 0 Then
                                    Dim Available = DSAvailable.Tables(0).Rows(0)
                                    'Get  with waiting for payment

                                    Dim CurrencyforSaleID As Integer = Available.Item("CurrencyforSaleID")
                                    Dim Pending As Double = GetCoinsPendingPayment(CurrencyforSaleID, BlockNumber)
                                    Dim CFSAvailableCoins As Double = Available.Item("Available")
                                    Dim AvailableCoins As Double = CFSAvailableCoins - Pending
                                    If AvailableCoins > 0 Then
                                        Dim UnitPrice As Double = Available.Item("BTCDesired") / Available.Item("AmountForSale")
                                        If AmountPurchased > AvailableCoins Then
                                            AmountPurchased = AvailableCoins
                                        End If
                                        Dim BuyerID As Integer = AddAddress(SenderAddress)
                                        Dim TotalBTC As Double = UnitPrice * AmountPurchased
                                        Dim TransFee As Double = Available.Item("TransFee")
                                        Dim PurchasedAmount As Double = AmountPurchased
                                        Dim MaxBlockNo As Double = BlockNumber + Available.Item("TimeLimit")
                                        SQL = "insert into z_PurchasingCurrency (PaidPurchasedAmount,MaxBlockNo,CurrencyforSaleID,PurchasedAmount,dTrans,AddressID,CurrencyID,ExodusID,PurchasingAmount,SellerID,IsPaid,UnitPrice,TotalBTC,TransFee,Paid) values ( " + _
                                            "0," + MaxBlockNo.ToString + "," + CurrencyforSaleID.ToString + "," + PurchasedAmount.ToString + ",'" + dTrans.ToString + "'," + BuyerID.ToString + "," + CurrencyID.ToString + "," + ExodusID.ToString + "," + AmountPurchasing.ToString + "," + SellerID.ToString + ",0," + UnitPrice.ToString + "," + TotalBTC.ToString + "," + TransFee.ToString + ",0)"
                                        Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                                        IsValidExodus = True
                                        Remarks = "Purchase Offer: " + PurchasedAmount.ToString + "  " + Currency
                                    Else
                                        Remarks = "Seller " + Receipient + " has no more coins for sale.  (CFS " + Format(CFSAvailableCoins, "########.########") + " Pending " + Format(Pending, "########.########") + ")"
                                    End If
                                Else
                                    Remarks = "No Offer to Sell from Seller " + Receipient + " or Currency " & Currency
                                End If
                            Case Else
                                'It maybe a payment transaction
                                IsPayment = True
                                Remarks = "Unknown Transaction Number "
                        End Select
                    End If
                End If

                If IsPayment Then
                    'Check if it is a payment
                    Dim BuyerID As Integer = GetAddressID(SenderAddress)
                    SQL = "select * from z_PurchasingCurrency where AddressID=" & BuyerID
                    Dim DSTest As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
                    If DSTest.Tables(0).DefaultView.Count > 0 Then

                        Dim OutArray As JArray
                        'Block Explorer
                        Dim AddressName As String = "address"
                        If Not IsDBNull(Row1("berawtx")) Then
                            OutArray = JsonConvert.DeserializeObject(Row1.Item("beOut"))
                        Else
                            Dim objBC As New JObject
                            objBC = JsonConvert.DeserializeObject(Trim(Row1("bcrawtx")))
                            OutArray = objBC.Item("out")
                            AddressName = "addr"
                        End If

                        For Each subitem In OutArray
                            If Not IsNothing(subitem(AddressName)) Then
                                Dim SubAddress As String = subitem(AddressName).ToString
                                Dim SubValue As String = ""
                                SubValue = subitem("value")
                                If AddressName = "addr" Then
                                    SubValue = (Val(subitem("value")) / 100000000).ToString
                                End If

                                If SubAddress <> "1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P" Then
                                    Dim BTCPaid As Double = Val(SubValue)
                                    Dim SellerAddress = SubAddress
                                    Dim SellerID = GetAddressID(SellerAddress)
                                    If SellerID > 0 Then
                                        SQL = "select * from z_PurchasingCurrency where AddressID=" & BuyerID & " and SellerID=" & SellerID & _
                                            " and Paid<TotalBTC and MaxBlockNo>= " & BlockNumber
                                        Dim DSPC As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
                                        If DSPC.Tables(0).DefaultView.Count > 0 Then
                                            Dim CurrencyID2 As Integer = DSPC.Tables(0).Rows(0).Item("CurrencyID")
                                            Dim CFSID As Integer = DSPC.Tables(0).Rows(0).Item("CurrencyforSaleID")
                                            SQL = "select * from z_CurrencyforSale where CurrencyforSaleID=" + CFSID.ToString
                                            Dim DSCFS As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
                                            If DSCFS.Tables(0).DefaultView.Count > 0 Then
                                                '                                                Dim UnitPrice As Double = Math.Round(DSCFS.Tables(0).Rows(0).Item("BTCDesired") / DSCFS.Tables(0).Rows(0).Item("AmountforSale"), 8)
                                                '                                                Dim AmountToSend2 As Double = Math.Round(BTCPaid / UnitPrice, 8)
                                                Dim AmountToSend2 As Double = Math.Round(BTCPaid * DSCFS.Tables(0).Rows(0).Item("AmountforSale") / DSCFS.Tables(0).Rows(0).Item("BTCDesired"), 8)
                                                If AmountToSend2 > DSPC.Tables(0).Rows(0).Item("PurchasedAmount") - DSPC.Tables(0).Rows(0).Item("PaidPurchasedAmount") Then
                                                    AmountToSend2 = DSPC.Tables(0).Rows(0).Item("PurchasedAmount") - DSPC.Tables(0).Rows(0).Item("PaidPurchasedAmount")
                                                    Remarks = "Overpaid"
                                                    'Check if there is still reserves from seller
                                                    'Dim Reserved As Double = (New mymastercoins).GetCurrencyForSale(SellerID, CurrencyID2)
                                                    'Dim Pending As Double = (New mymastercoins).GetSellerCoinsPendingPayment(SellerID, BlockNumber, CurrencyID2)
                                                    'Dim BuyerPurchaseAmountRemaining = DSPC.Tables(0).Rows(0).Item("PurchasedAmount") - DSPC.Tables(0).Rows(0).Item("PaidPurchasedAmount")
                                                    'If Reserved - Pending + BuyerPurchaseAmountRemaining < AmountToSend2 Then
                                                    'AmountToSend2 = Reserved - Pending + BuyerPurchaseAmountRemaining
                                                    'End If

                                                End If

                                                'Check if Seller has enough balance
                                                Dim Balance As Double = GetAddressBalance(SellerID, CurrencyID2)
                                                If Balance >= AmountToSend2 Then
                                                    SQL = "update z_PurchasingCurrency set PaidPurchasedAmount=PaidPurchasedAmount+" + AmountToSend2.ToString + _
                                                        ",Paid=Paid+" + BTCPaid.ToString + _
                                                        ",IsPaid=1 where PurchasingCurrencyID=" & DSPC.Tables(0).Rows(0).Item("PurchasingCurrencyID")
                                                    Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                                                    Call AddTransaction(dTrans, SellerID, ExodusID, SenderAddress, 0, AmountToSend2, CurrencyID2, BTCPaid.ToString, CFSID)
                                                    Call AddTransaction(dTrans, BuyerID, ExodusID, SellerAddress, AmountToSend2, 0, CurrencyID2, BTCPaid.ToString)
                                                    SQL = "update z_CurrencyforSale set Available=Available- " + AmountToSend2.ToString + " where CurrencyforSaleID=" + CFSID.ToString
                                                    Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                                                    IsValidExodus = True
                                                    Remarks = "Payment: " + BTCPaid.ToString + " btc for " + AmountToSend2.ToString + " " + (New Bitcoin).GetCurrency(CurrencyID2)
                                                    Exit For
                                                Else
                                                    Remarks = "Buyer Payment:  Seller " + SellerAddress + " doesn't have enough coins to send."
                                                End If
                                            Else
                                                Remarks = "Currency for Sale Transaction " + CFSID.ToString + "not found."
                                            End If
                                        Else
                                            Remarks = "Buyer Payment:  Purchase confirmation not found or payment time has expired."
                                        End If
                                    Else
                                        Remarks = "Buyer Payment:  Seller " + SellerAddress + " not found."
                                    End If
                                End If

                            End If
                        Next
                    End If
                End If
            Else
                Remarks = "Invalid Sender (Multisig Input) "
            End If

            If IsValidExodus Then
                SQL = "update z_Exodus set IsValid=1 where ExodusID= " + ExodusID.ToString
            Else
                SQL = "update z_Exodus set IsValid=0 where ExodusID= " + ExodusID.ToString
            End If
            Call (New AWS.DB.ConnectDB).SQLExecute(SQL)

            SQL = "update z_Exodus set Sender='" + Trim(SenderAddress) + "', Receiver='" + Trim(Receipient) + "',CurrencyID=" + CurrencyID + ",IsProcessed=1,Remarks='" + (New AWS.DB.ConnectDB).HTMLEncode(Remarks) + "' where ExodusID= " + ExodusID.ToString
            Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
        Next
        Form1.clearprogbar()

    End Sub

    Function GetbcSender(ByVal bcIn As String) As String

        Dim InArray As JArray
        Dim Prevouthash As String = ""
        Dim n As Integer = 0
        InArray = JsonConvert.DeserializeObject(bcIn)
        Dim Value As Double = 0
        Dim Address As String = ""
        Dim aAddress(10) As String
        Dim aValue(10) As Long
        Dim nCtr As Integer = 0
        Dim IsMultiSigInput = False
        For Each subitem In InArray
            If Not IsNothing(subitem.Item("prev_out").Item("addr2")) Then
                IsMultiSigInput = True
                Exit For
                '                Dim First4 As String = Mid(subitem.Item("script"), 5, 4)0
                'If First4 = "3044" Or First4 = "3045" Or First4 = "3046" Then
                'MultiSig
                '                Dim First6 As String = Left(subitem.Item("script"), 6)
                '               If First6 = "473044" Or First6 = "473045" Or First6 = "473046" Then
                'End If
            End If
            Address = Trim(subitem.Item("prev_out").Item("addr").ToString)
            Value = Val(subitem.Item("prev_out").Item("value"))
            Dim i As Integer
            Dim lFound As Boolean = False
            For i = 0 To nCtr
                If aAddress(i) = Address Then
                    aValue(i) += Value
                    lFound = True
                    Exit For
                End If
            Next
            If Not lFound Then
                nCtr += 1
                aAddress(nCtr) = Address
                aValue(nCtr) = Value
            End If
        Next
        Dim InputAddress As String = ""
        If Not IsMultiSigInput Then
            Dim j As Integer
            Dim HighestVal As Long = 0
            Dim EqualHighest As Boolean = False
            For j = 0 To nCtr
                If aValue(j) > HighestVal Then
                    InputAddress = aAddress(j)
                    HighestVal = aValue(j)
                    EqualHighest = False
                Else
                    If aValue(j) = HighestVal Then
                        EqualHighest = True
                    End If
                End If
            Next
            If EqualHighest Then
                InputAddress = ""
            End If
        End If
        Return InputAddress
    End Function

    Function GetbeSender(ByVal beIn As String) As String

        Dim InArray As JArray
        Dim Prevouthash As String = ""
        Dim n As Integer = 0
        InArray = JsonConvert.DeserializeObject(beIn)
        Dim Value As Double = 0
        Dim Address As String = ""
        Dim aAddress(10) As String
        Dim aValue(10) As Double
        Dim nCtr As Integer = 0

        Dim IsMultiSigInput = False
        For Each subitem In InArray
            If subitem("scriptSig").ToString <> "" Then
                If Left(subitem("scriptSig").ToString, 2) = "00" Then
                    IsMultiSigInput = True
                    Exit For
                End If
            End If
            Prevouthash = subitem("prev_out").Item("hash")
            n = CInt(subitem("prev_out").Item("n"))
            Address = GetReferenceFromPrevHash(Prevouthash, n, Value)
            Dim i As Integer
            Dim lFound As Boolean = False
            For i = 0 To nCtr
                If aAddress(i) = Address Then
                    aValue(i) += Value
                    lFound = True
                    Exit For
                End If
            Next
            If Not lFound Then
                nCtr += 1
                aAddress(nCtr) = Address
                aValue(nCtr) = Value
            End If
        Next
        Dim j As Integer
        Dim InputAddress As String = ""
        If Not IsMultiSigInput Then
            Dim HighestVal As Double = 0
            For j = 0 To nCtr
                If aValue(j) > HighestVal Then
                    InputAddress = aAddress(j)
                    HighestVal = aValue(j)
                End If
            Next
        End If
        Return InputAddress
    End Function
    Sub GetbeReceipient(ByVal beOut As String, ByRef ClassAReceipient As String, ByRef ClassAData As String, ByRef ThirdAddress As String)
        Dim OutArray As JArray
        OutArray = JsonConvert.DeserializeObject(beOut)
        Dim Value As Double = 99999999
        Dim Address As String = ""

        For Each subitem In OutArray
            If IsNothing(subitem("address")) Then
                Dim Hash160 As String = GetStringBetween(subitem("scriptPubKey"), "OP_HASH160", "OP_EQUAL")
                If Hash160.Length > 0 Then
                    Address = (New Bitcoin).Hash160toAddress(Trim(Hash160))
                End If
            Else
                Address = Trim(subitem("address"))
            End If

            If Address = "1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P" And Val(subitem("value")) < Value Then
                Value = Val(subitem("value"))
            End If
        Next


        Dim PossibleData As String = ""
        For Each subitem In OutArray
            If IsNothing(subitem("address")) Then
                Dim Hash160 As String = GetStringBetween(subitem("scriptPubKey"), "OP_HASH160", "OP_EQUAL")
                If Hash160.Length > 0 Then
                    Address = (New Bitcoin).Hash160toAddress(Trim(Hash160))
                End If
            Else
                Address = Trim(subitem("address"))
            End If

            If Address <> "" And Address <> "1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P" Then
                If Val(subitem("value")) = Value Then
                    ThirdAddress = ClassAData
                    ClassAData = ClassAReceipient
                    ClassAReceipient = Address
                Else
                    If PossibleData = "" Then
                        PossibleData = Address
                    End If
                End If
            End If
        Next
        If ClassAData = "" Then
            ClassAData = PossibleData
        End If

    End Sub
    Function GetExodusID(ByVal TxID As String) As Integer
        Dim SQL As String = "SELECT  * from z_Exodus where TxID='" & TxID & "'"
        Dim ExodusID As Integer = 0
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
        If DS.Tables(0).DefaultView.Count > 0 Then
            ExodusID = DS.Tables(0).Rows(0).Item("ExodusID")
        End If
        Return ExodusID
    End Function
    Sub Getreferencefrombin()
        '        Dim s1 = (New Bitcoin).DecodeBTCAddresstoHexString("155DH9FHD1PEjMweTJjuSZceKN92TPTnj9")
        '       Dim temp As String = "2cac6a9954415885b285de8833fc3056eb0e1a23"
        '      Dim s = (New Bitcoin).Hash160toAddress(temp)
        Dim SQL As String = "SELECT  * from z_Exodus where (beReferenceAddress Is Null) or len(beReferenceAddress) <30 "
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
        For Each Row1 In DS.Tables(0).Rows
            Dim ExodusID As Integer = Row1.Item("ExodusID")
            Dim ReferenceAddress As String = GetReference(Row1.Item("beIn"))

            SQL = "Update z_Exodus set beReferenceAddress='" + ReferenceAddress + "' where ExodusID=" & ExodusID
            Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
        Next

        If False Then
            Dim PostURL As String = "C:\Users\Nivla\Downloads\1Exodus.txt"
            Dim json As String = (New Bitcoin).getjson(PostURL)
            Dim obj As New JObject
            obj = JsonConvert.DeserializeObject(json)
            Dim results As List(Of JToken) = obj.Children().ToList
            Dim Ctr As Integer = 0
            For Each item As JProperty In results
                item.CreateReader()
                Dim obj2 As New JObject
                obj2 = JsonConvert.DeserializeObject(item.Value.ToString)
                Dim hash As String = obj2.Item("hash")
                Dim beIn As String = obj2.Item("in").ToString
                Dim beOut As String = obj2.Item("out").ToString
                Dim ReferenceAddress As String = GetReference(beIn)
                SQL = "Update z_Exodus set beReferenceAddress='" + ReferenceAddress + "' where TxID='" & hash & "'"
                Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
            Next

        End If

    End Sub

    '12/6/2013
    Function GetReference(ByVal beIn As String) As String
        Dim ReferenceAddress As String = ""
        Dim InArray As New JArray
        InArray = JsonConvert.DeserializeObject(beIn)
        Dim Prevouthash As String = ""
        Dim n As Integer = 0
        Dim LargestInput As Double = 0
        Dim Value As Double = 0
        For Each subitem In InArray
            Prevouthash = subitem("prev_out").Item("hash")
            n = CInt(subitem("prev_out").Item("n"))
            Dim ReferenceAddress2 As String = GetReferenceFromPrevHash(Prevouthash, n, Value)
            If Value > LargestInput Then
                LargestInput = Value
                ReferenceAddress = ReferenceAddress2
            End If
        Next
        Return ReferenceAddress
    End Function
    '12/6/2013
    Function GetReferenceFromPrevHash(ByVal PrevHash As String, ByVal n As Integer, ByRef Value1 As Double) As String
        Dim Address As String = ""
        Dim json As String = (New Bitcoin).getjson("http://blockexplorer.com/rawtx/" + PrevHash)
        Dim objX = JsonConvert.DeserializeObject(json)
        Dim OutArray As New JArray
        OutArray = JsonConvert.DeserializeObject(objX.item("out").ToString)
        Dim Ctr As Integer = 0
        For Each subitem In OutArray
            If n = Ctr Then
                Dim Hash160 As String = GetStringBetween(subitem("scriptPubKey"), "OP_DUP OP_HASH160", "OP_EQUALVERIFY OP_CHECKSIG")
                If Hash160.Length = 0 Then
                    Dim InArray As New JArray
                    InArray = JsonConvert.DeserializeObject(objX.item("in").ToString)
                    For Each subitemIn In InArray
                        Dim Prevouthash2 As String = subitemIn("prev_out").Item("hash")
                        Dim n2 As Integer = CInt(subitemIn("prev_out").Item("n"))
                        Address = GetReferenceFromPrevHash(Prevouthash2, n2, Value1)
                        Exit For
                    Next
                Else
                    Address = (New Bitcoin).Hash160toAddress(Trim(Hash160))
                    Value1 = CDbl(subitem("value"))
                End If
                Exit For
            End If
            Ctr = Ctr + 1
        Next
        Return Address
    End Function
    Public Function GetStringBetween(ByVal InputText As String, _
                                     ByVal starttext As String, _
                                     ByVal endtext As String) As String

        Dim startPos As Integer
        Dim endPos As Integer
        Dim lenStart As Integer
        Dim RetVal As String = ""
        startPos = InputText.IndexOf(starttext, StringComparison.CurrentCultureIgnoreCase)
        If startPos >= 0 Then
            lenStart = startPos + starttext.Length
            endPos = InputText.IndexOf(endtext, lenStart, StringComparison.CurrentCultureIgnoreCase)
            If endPos >= 0 Then
                RetVal = InputText.Substring(lenStart, endPos - lenStart)
            End If
        End If
        Return RetVal
    End Function
    Function IsDataCode(ByVal Address As String) As Boolean
        Dim PeekR As String = (New Bitcoin).DecodeBTCAddresstoHexString(Address)
        Return PeekR.Substring(4, 14) = "00000000000000" And _
            (PeekR.Substring(18, 2) = "01" Or PeekR.Substring(18, 2) = "02")
    End Function
    Function GetCoinsPendingPayment(ByVal CurrencyforSaleID As Integer, ByVal BlockNumber As Integer) As Double
        Dim Pending As Double = 0
        Dim Sql As String = "SELECT sum(PurchasedAmount) as Pending" + _
" FROM z_PurchasingCurrency " + _
" WHERE TotalBTC<Paid and  CurrencyForSaleID=" + CurrencyforSaleID.ToString + _
" and MaxBlockNo>=" + BlockNumber.ToString
        Dim DSPending As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
        'Deduct from available those waiting for payment
        If DSPending.Tables(0).DefaultView.Count > 0 Then
            If Not IsDBNull(DSPending.Tables(0).Rows(0).Item("Pending")) Then
                Pending = DSPending.Tables(0).Rows(0).Item("Pending")
            End If
        End If
        Return Pending
    End Function

    Function GetAddressID(ByVal Address As String) As Integer
        Address = (New AWS.DB.ConnectDB).HTMLEncode(Address)
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset("select * from z_Address where Address='" & Address & "'")
        Dim ID As Integer = 0
        If DS.Tables(0).DefaultView.Count > 0 Then
            ID = DS.Tables(0).Rows(0).Item("AddressID")
        End If
        Return ID
    End Function

    Function AddAddress(ByVal Address As String) As Integer
        If Address.Length = 34 Or Address.Length = 33 Then
            Dim AddressID As Integer = GetAddressID(Address)
            If AddressID = 0 Then
                Dim Sql As String = "insert into z_Address (Address,MSC,TMSC) values ('" + (New AWS.DB.ConnectDB).HTMLEncode(Address) + "',0,0) "
                Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
                AddressID = GetAddressID(Address)
            End If
            Return AddressID
        Else
            'Invalid address
            Return 0
        End If
    End Function
    Sub AddTransaction(ByVal dTrans As DateTime, ByVal AddressID As Integer, ByVal ExodusID As Integer, ByVal Description As String, ByVal AmountIn As Double, ByVal AmountOut As Double, ByVal CurrencyID As Integer, Optional ByVal TotalPrice As Double = 0, Optional ByVal CFSID As Integer = 0)
        Dim Sql As String = "insert into z_Transactions (CFSID,dTrans,AddressID,ExodusID,Description,AmountIn,AmountOut,CurrencyID,TotalPrice) values (" + _
            CFSID.ToString + ",'" + dTrans.ToString + "'," + AddressID.ToString + "," + ExodusID.ToString + ",'" + (New AWS.DB.ConnectDB).HTMLEncode(Description) + "'," + AmountIn.ToString + "," + AmountOut.ToString + ", " + CurrencyID.ToString + "," + TotalPrice.ToString + " ) "
        Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
        Dim CurrencyCode As String = (New Bitcoin).GetCurrency(CurrencyID)
        Sql = "update z_Address set " & CurrencyCode & "=" & CurrencyCode & "+" & (AmountIn - AmountOut).ToString & " where AddressID=" & AddressID
        Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
    End Sub

    Function GetAddressBalance(ByVal AddressID As Integer, ByVal CurrencyID As Integer) As Double
        Dim Sql As String = "select * from z_Address where AddressID=" & AddressID.ToString
        Dim DS2 As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
        Dim Balance As Double = 0
        If DS2.Tables(0).DefaultView.Count > 0 Then
            Balance = DS2.Tables(0).Rows(0).Item((New Bitcoin).GetCurrency(CurrencyID))
        End If
        Return Balance
    End Function

    Function GetCurrencyForSale(ByVal AddressID As Integer, ByVal CurrencyID As Integer) As Double
        Dim Sql As String = "select * from z_CurrencyforSale where AddressID=" & AddressID.ToString & " and CurrencyID= " & CurrencyID & " and IsNewOffer=1 order by dTrans"
        Dim DSAvailable As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
        Dim Result As Double = 0
        If DSAvailable.Tables(0).DefaultView.Count > 0 Then
            Dim CurrencyforSaleID As Integer = DSAvailable.Tables(0).Rows(0).Item("CurrencyforSaleID")
            '            Dim CoinsPending As Double = (New mymastercoins).GetCoinsPendingPayment(CurrencyforSaleID, CurrentBlockNo)
            Dim Available As Double = DSAvailable.Tables(0).Rows(0).Item("Available")
            '            Result = Available - CoinsPending
            Result = Available
        End If
        Return Result
    End Function

    Function GetSellerCoinsPendingPayment(ByVal SellerID As Integer, ByVal BlockNumber As Integer, ByVal CurrencyID As Integer) As Double
        Dim Pending As Double = 0
        Dim Sql As String = "SELECT sum(PurchasedAmount) as Pending" + _
" FROM z_PurchasingCurrency " + _
" WHERE  SellerID=" + SellerID.ToString + _
" and MaxBlockNo>=" + BlockNumber.ToString + _
" and CurrencyID>=" + CurrencyID.ToString
        Dim DSPending As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
        If DSPending.Tables(0).DefaultView.Count > 0 Then
            If Not IsDBNull(DSPending.Tables(0).Rows(0).Item("Pending")) Then
                Pending = DSPending.Tables(0).Rows(0).Item("Pending")
            End If
        End If
        Return Pending
    End Function

    Function GetAddress(ByVal AddressID As Integer) As String
        Dim Sql As String = "select * from z_Address where AddressID=" & AddressID.ToString
        Dim DS2 As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
        Dim Address As String = ""
        If DS2.Tables(0).DefaultView.Count > 0 Then
            Address = Trim(DS2.Tables(0).Rows(0).Item("Address"))
        End If
        Return Address
    End Function
    Function GetShortAddress(ByVal AddressID As Integer) As String
        Dim Sql As String = "select * from z_Address where AddressID=" & AddressID.ToString
        Dim DS2 As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
        Dim Address As String = ""
        If DS2.Tables(0).DefaultView.Count > 0 Then
            Address = Left(Trim(DS2.Tables(0).Rows(0).Item("Address")), 6)
        End If
        Return Address
    End Function
    Function GetShortTxID(ByVal ExodusID As Integer) As String
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset("select TxID from z_Exodus where ExodusID=" & ExodusID)
        Dim TXID As String = ""
        If DS.Tables(0).DefaultView.Count > 0 Then
            TXID = Left(DS.Tables(0).Rows(0).Item("TXID"), 6)
        End If
        Return TXID
    End Function

    Function IsTxIDinExodus(ByVal hash As String) As Integer
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset("select * from z_Exodus where TxID='" & hash & "'")
        Return DS.Tables(0).DefaultView.Count > 0
    End Function

    Function GetExodusTransBlockExplorer() As String
        Dim SQL As String = "SELECT top 1 block,blocknumber from z_Exodus order by dTrans desc,BlockNumber desc "
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL)
        Dim StartBlock As String = ""
        Dim StartBlockNumber As String = ""
        If DS.Tables(0).DefaultView.Count > 0 Then
            If Not IsDBNull(DS.Tables(0).Rows(0).Item("block")) Then
                StartBlock = Trim(DS.Tables(0).Rows(0).Item("block"))
            End If
            StartBlockNumber = Trim(DS.Tables(0).Rows(0).Item("blocknumber").ToString)
        End If
        Dim temp As String = "http://blockexplorer.com/q/getblockhash/" + Trim(StartBlockNumber)
        If StartBlock = "" Then
            StartBlock = (New Bitcoin).getjson(temp)
        End If
        Dim PostURL As String = "http://blockexplorer.com/q/mytransactions/1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P/" + StartBlock
        Dim Result As String = PostURL
        Dim json As String = (New Bitcoin).getjson(PostURL)
        Dim obj As New JObject
        Dim lError = False

        Try
            obj = JsonConvert.DeserializeObject(json)
        Catch ex As Exception
            lError = True
        End Try
        If Not lError Then
            Dim results As List(Of JToken) = obj.Children().ToList
            Form1.ResetProgbar(results.Count)
            Dim dTrans As String
            For Each item As JProperty In results
                Form1.updateprogbar()
                item.CreateReader()

                Dim obj2 As New JObject
                obj2 = JsonConvert.DeserializeObject(item.Value.ToString)
                Dim hash As String = obj2.Item("hash")
                dTrans = obj2.Item("time")
                Dim blocknumber As Integer = obj2.Item("blocknumber")
                Dim block As String = obj2.Item("block")
                Dim beIn As String = (New AWS.DB.ConnectDB).HTMLEncode(obj2.Item("in").ToString)
                Dim beOut As String = (New AWS.DB.ConnectDB).HTMLEncode(obj2.Item("out").ToString)

                hash = (New AWS.DB.ConnectDB).HTMLEncode(hash)
                If Not (New mymastercoinsold2).IsTxIDinExodus(hash) Then
                    PostURL = "http://blockexplorer.com/rawtx/" + hash
                    json = (New Bitcoin).getjson(PostURL)
                    Dim berawtx As String = (New AWS.DB.ConnectDB).HTMLEncode(json)

                    '                    PostURL = "http://blockchain.info/rawtx/" + hash
                    '                   json = (New Bitcoin).getjson(PostURL)
                    '                    Dim bcrawtx As String = (New AWS.DB.ConnectDB).HTMLEncode(json)
                    'Depracated function
                    Dim bcrawtx As String = ""

                    Dim beReferenceAddress As String = ""
                    beReferenceAddress = GetbeSender(beIn)

                    SQL = "insert into z_Exodus (beReferenceAddress,beIn,beOut,block,TxID,dTrans,blocknumber,IsValid,IsProcessed,berawtx,bcrawtx) values (" + _
                        "'" + beReferenceAddress + "','" + beIn + "','" + beOut + "','" + block + "','" + hash + "','" + dTrans.ToString + "'," + blocknumber.ToString + _
                         ",0,0," + _
                         "'" + berawtx + "'," + _
                         "'" + bcrawtx + "')"

                    Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                    Result += hash + vbCrLf
                Else
                    SQL = "update  z_Exodus set beIn='" + beIn + "',beOut='" + beOut + "',block='" + block + "' where TXID='" + hash + "'"
                    Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
                End If
            Next
        End If
        UpdateExodusDevFunds()
        Form1.clearprogbar()
        Return Result
    End Function
    Function GetExodusTransBlockChain() As String
        Dim Offset As Integer = 0
        Dim lStop = False
        Dim Result As String = ""

        Do While Not lStop
            Dim PostURL As String = "http://blockchain.info/address/1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P?format=json&limit=50&offset=" & Offset
            Dim json As String = (New Bitcoin).getjson(PostURL)
            Dim obj As New JObject
            obj = JsonConvert.DeserializeObject(json)
            Dim Hash As String
            If json.Length > 0 Then
                For i = 0 To obj.Item("txs").Count - 1
                    Hash = obj.Item("txs").Item(i).Item("hash").ToString
                    Dim UTC As Long = obj.Item("txs").Item(i).Item("time").ToString
                    Dim dTrans As DateTime = DateAdd(DateInterval.Second, UTC, New DateTime(1970, 1, 1, 0, 0, 0))

                    Dim blocknumber As Integer = 0
                    If Not obj.Item("txs").Item(i).Item("block_height") Is Nothing Then
                        blocknumber = obj.Item("txs").Item(i).Item("block_height").ToString
                    End If


                    If blocknumber > 0 Then
                        Hash = (New AWS.DB.ConnectDB).HTMLEncode(Hash)
                        If Not IsTxIDinExodus(Hash) Then
                            PostURL = "http://blockchain.info/rawtx/" + Hash + "?scripts=true"
                            json = (New Bitcoin).getjson(PostURL)
                            Dim bcrawtx As String = (New AWS.DB.ConnectDB).HTMLEncode(json)
                            If Not bcrawtx = "" Then
                                '                                Dim dTrans As DateTime = GetDateTimefromBlock(blocknumber)
                                Dim Sql As String = "insert into z_TempExodus (TxID,dTrans,blocknumber,bcrawtx) values (" + _
                                    "'" + Hash + "','" + dTrans.ToString + "'," + blocknumber.ToString + _
                                     ",'" + (New AWS.DB.ConnectDB).HTMLEncode(bcrawtx) + "')"
                                Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
                                Result += i.ToString + dTrans.ToString + " " + Hash + vbCrLf
                            End If
                        Else
                            lStop = True
                        End If
                    End If
                Next
            Else
                lStop = True
                Result = "API call blank"
            End If

            Offset = Offset + 50
        Loop
        Dim SQL2 As String = "select * from z_TempExodus order by TempExodusID desc"
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset(SQL2)
        If DS.Tables(0).DefaultView.Count > 0 Then

            For Each obj2 In DS.Tables(0).Rows
                Dim txid As String = Trim(obj2.Item("txid"))
                Dim dTrans As DateTime = obj2.Item("dTrans")
                Dim blocknumber As Integer = obj2.Item("blocknumber")
                Dim bcrawtx As String = obj2.Item("bcrawtx")
                If Not IsTxIDinExodus(txid) Then
                    SQL2 = "insert into z_Exodus (TxID,dTrans,blocknumber,IsValid,IsProcessed,bcrawtx) values (" + _
                        "'" + txid + "','" + dTrans.ToString + "'," + blocknumber.ToString + _
                         ",0,0," + _
                         "'" + bcrawtx + "')"
                    Call (New AWS.DB.ConnectDB).SQLExecute(SQL2)
                End If
                Call (New AWS.DB.ConnectDB).SQLExecute("delete from z_TempExodus where TempExodusID=" & obj2.Item("TempExodusID"))
            Next
        End If


        UpdateExodusDevFunds()
        Return Result
    End Function

    Function GetBlockTime(ByVal BlockHeight As Integer) As DateTime
        Dim PostURL As String = "https://blockchain.info/block-height/" + BlockHeight.ToString + "?format=json"
        Dim json As String = (New Bitcoin).getjson(PostURL)
        Dim obj As New JObject
        obj = JsonConvert.DeserializeObject(json)
        Dim UnixTime As Integer = obj.Item("blocks").Item(0).Item("time").ToString
        Dim dTrans As DateTime = DateAdd(DateInterval.Second, UnixTime, New DateTime(1970, 1, 1, 0, 0, 0))

        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset("select * from z_Blocks where BlockHeight=" & BlockHeight)
        If DS.Tables(0).DefaultView.Count = 0 Then
            Dim SQL As String = "insert into z_Blocks (blockheight,dTrans) values (" + BlockHeight.ToString + ",'" + dTrans.ToString + "')"
            Call (New AWS.DB.ConnectDB).SQLExecute(SQL)
        End If
        Return dTrans
    End Function

    Function GetDateTimefromBlock(ByVal blockheight As Integer) As DateTime
        Dim dTrans As DateTime
        Dim DS As DataSet = (New AWS.DB.ConnectDB).SQLdataset("select * from z_Blocks where BlockHeight=" & blockheight)
        If DS.Tables(0).DefaultView.Count > 0 Then
            dTrans = DS.Tables(0).Rows(0).Item("dTrans")
        Else
            dTrans = GetBlockTime(blockheight)
        End If
        Return dTrans
    End Function
    Function GetTxID(ByVal ExodusID As Integer) As String
        Dim Sql As String = "select * from z_Exodus where ExodusID=" & ExodusID.ToString
        Dim DS2 As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
        Dim TxID As String = ""
        If DS2.Tables(0).DefaultView.Count > 0 Then
            TxID = Trim(DS2.Tables(0).Rows(0).Item("TxID"))
        End If
        Return TxID
    End Function
    Sub UpdateExodusDevFunds()
        '        Dim Sql As String = "SELECT top 1 dtrans from z_Exodus order by dTrans desc,BlockNumber desc "
        Dim dtrans As DateTime = Now()
        Dim BlockCount As Integer = 0

        Dim PostURL As String = "http://blockchain.info/latestblock"
        Dim json As String = (New Bitcoin).getjson(PostURL)
        Dim obj As New JObject
        obj = JsonConvert.DeserializeObject(json)
        If Not obj Is Nothing Then
            BlockCount = obj.Item("height")
            Dim ExodusAddressID As Integer = AddAddress("1EXoDusjGwvnjZUyKkxZ4UHEf77z6A5S4P")
            Dim SecondsFromSept12013 As Integer = 0
            Dim Time1 As Long = CInt(obj.Item("time"))
            dtrans = DateAdd(DateInterval.Second, Time1, New DateTime(1970, 1, 1, 0, 0, 0))
            SecondsFromSept12013 = Time1 - 1377993874

            Dim Gencoins As Double = (1 - (0.5 ^ (SecondsFromSept12013 / 31556926))) * 56316.23576222
            Gencoins = Math.Round(Gencoins, 8)
            Dim Sql As String = "select * from z_Transactions where AddressID=" + ExodusAddressID.ToString + " and dTrans='2013/09/01' and CurrencyID=1"
            Dim DSX As DataSet = (New AWS.DB.ConnectDB).SQLdataset(Sql)
            If DSX.Tables(0).DefaultView.Count = 0 Then
                Call AddTransaction(#9/1/2013#, ExodusAddressID, 0, "Generated Coins MSC Dev Funds as of " + dtrans.ToString, Gencoins, 0, 1)
            Else
                Sql = "Update z_Transactions set description='Generated Coins MSC Dev Funds as of " + dtrans.ToString + "  Block: " + BlockCount.ToString + "', AmountIn=" + Gencoins.ToString + " where AddressID=" + ExodusAddressID.ToString + " and dTrans='2013/09/01' and CurrencyID=1"
                Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
            End If
            'recalculate Exodus Address Balance
            Sql = "select sum(AmountIn-AmountOut) as Balance from z_Transactions where  AddressID=" + ExodusAddressID.ToString + " and  CurrencyID=1"
            DSX = (New AWS.DB.ConnectDB).SQLdataset(Sql)
            Sql = "update z_Address set MSC=" & DSX.Tables(0).Rows(0).Item("Balance") & " where AddressID=" & ExodusAddressID.ToString
            Call (New AWS.DB.ConnectDB).SQLExecute(Sql)
        End If
    End Sub

End Class
