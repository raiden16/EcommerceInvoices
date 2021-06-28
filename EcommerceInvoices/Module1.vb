Imports System.Net.Mail
Imports AE.Net.Mail.ImapClient
Imports AE.Net.Mail.SearchCondition
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Module Module1

    Public SBOCompany As SAPbobsCOM.Company

    Sub Main()

        Conecction()
        ReadEmail()
        Disconnect()

    End Sub

    Public Function Conecction()

        Try

            SBOCompany = New SAPbobsCOM.Company

            SBOCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            SBOCompany.Server = My.Settings.Server
            SBOCompany.LicenseServer = My.Settings.LicenseServer
            SBOCompany.DbUserName = My.Settings.DbUserName
            SBOCompany.DbPassword = My.Settings.DbPassword

            SBOCompany.CompanyDB = My.Settings.CompanyDB

            SBOCompany.UserName = My.Settings.UserName
            SBOCompany.Password = My.Settings.Password

            SBOCompany.Connect()

        Catch ex As Exception

            Dim stError As String
            stError = "Error al tratar de hacer conexión con SAP B1. " & ex.Message
            Setlog(stError, " ", " ", " ", " ", " ")

        End Try

    End Function


    Public Function ReadEmail()

        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim uids2 As String()
        Dim uidsf As String()
        Dim order, sku, Destinatario, DestinatarioPass As String
        Dim ic As AE.Net.Mail.ImapClient

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        order = " "
        sku = " "

        Try

            stQueryH = "Select * from ""@CORREOTEKNO"" where ""Name""='Ecommerce'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                Destinatario = oRecSetH.Fields.Item("U_Email").Value
                DestinatarioPass = oRecSetH.Fields.Item("U_Password").Value

            End If

            ic = New AE.Net.Mail.ImapClient(My.Settings.IMAP, Destinatario, DestinatarioPass, AE.Net.Mail.AuthMethods.Login, My.Settings.Puerto, True)
            ic.SelectMailbox("INBOX")
            uids2 = ic.Search(Unseen)
            uidsf = ic.Search(From(My.Settings.Remitente))

            For Each uid As String In uids2

                For Each uidf As String In uidsf

                    If uidf = uid Then

                        Dim mensaje As MailMessage = ic.GetMessage(uid)
                        Dim skuidx, skulgt, skucmpt, skureverse, skulike, qtyidx, dateidx, qtylgt, itemidx, itemlgt, itemcmpt, itemreverse, statusidx, priceidx, pricelgt, pricecmpt, pricereverse,
                            taxidx, dryidx, drylgt, drycmpt, dryreverse, taxdryidx, measures, measuresreverse, orderidx, orderlgt, mssgidx, body, header, frommsg, datemsg, solditemsidx,
                            solditemslgt, solditems, pthsidx, profits, steps, limit, dlvrydateidx, dlvrydatelgt, dlvrydatecmpt, statuslgt, statuscmpt, amazonidx, amazonlgt, amazoncmpt, profitsidx, profitslgt,
                            profitscmpt, profitslbl, pathpdf, soldmssg As String
                        Dim measure1, measure2, m2, qty, lineTotal As Double
                        Dim price, totalm2, dry As Decimal

                        frommsg = mensaje.From.ToString
                        datemsg = Now.Date.ToString
                        header = mensaje.Subject.ToString
                        body = mensaje.Body.ToString

                        soldmssg = header.IndexOf("Vendido, ¡envíalo ya!:")

                        If soldmssg > 0 Then

                            orderidx = mensaje.Body.ToString.IndexOf("Número de pedido:")
                            mssgidx = mensaje.Body.ToString.IndexOf("Envía")
                            orderlgt = mssgidx - (orderidx + 17)
                            order = mensaje.Body.ToString.Substring(orderidx + 17, orderlgt).Trim

                            solditemsidx = header.IndexOf("artículos vendidos")

                            If solditemsidx > 0 Then

                                pthsidx = header.IndexOf("(")
                                solditemslgt = solditemsidx - (pthsidx + 1)
                                solditems = header.Substring(pthsidx + 1, solditemslgt).Trim

                            Else

                                solditems = "1"

                            End If

                            skuidx = mensaje.Body.ToString.IndexOf("SKU:")
                            qtyidx = mensaje.Body.ToString.IndexOf("Cantidad:")
                            skulgt = qtyidx - (skuidx + 4)
                            skucmpt = mensaje.Body.ToString.Substring(skuidx + 4, skulgt)
                            skureverse = StrReverse(skucmpt)
                            sku = mensaje.Body.ToString.Substring(skuidx + 4, skulgt - (skureverse.IndexOf("-") + 1)).Trim

                            dateidx = mensaje.Body.ToString.IndexOf("Fecha del pedido:")
                            qtylgt = dateidx - (qtyidx + 9)
                            qty = mensaje.Body.ToString.Substring(qtyidx + 9, qtylgt).Trim

                            skulike = sku.Substring(0, 2)

                            If skulike = "TG" Then

                                itemidx = mensaje.Body.ToString.IndexOf("Artículo:")
                                statusidx = mensaje.Body.ToString.IndexOf("Estado:")
                                itemlgt = statusidx - (itemidx + 9)
                                itemcmpt = mensaje.Body.Substring(itemidx + 9, itemlgt)
                                itemreverse = StrReverse(itemcmpt)

                                measures = StrReverse(itemreverse.Substring(0, itemreverse.IndexOf("|")))
                                measure1 = measures.Substring(0, measures.IndexOf("m"))
                                measuresreverse = StrReverse(measures).Trim
                                measure2 = StrReverse(measuresreverse.Substring(1, measuresreverse.IndexOf("x") - 1))
                                m2 = measure1 * measure2

                                totalm2 = qty * m2

                            ElseIf skulike = "SG" Then

                                itemidx = mensaje.Body.ToString.IndexOf("Artículo:")
                                statusidx = mensaje.Body.ToString.IndexOf("Estado:")
                                itemlgt = statusidx - (itemidx + 9)
                                itemcmpt = mensaje.Body.Substring(itemidx + 9, itemlgt)
                                itemreverse = StrReverse(itemcmpt)

                                measures = StrReverse(itemreverse.Substring(0, itemreverse.IndexOf("|")))
                                measure1 = measures.Substring(0, measures.IndexOf("piezas")).Trim
                                m2 = measure1 / 4

                                totalm2 = qty * m2

                            Else

                                itemidx = mensaje.Body.ToString.IndexOf("Artículo:")
                                statusidx = mensaje.Body.ToString.IndexOf("Estado:")
                                itemlgt = statusidx - (itemidx + 9)
                                itemcmpt = mensaje.Body.Substring(itemidx + 9, itemlgt)
                                totalm2 = qty

                            End If

                            priceidx = mensaje.Body.ToString.IndexOf("Precio:")
                            taxidx = mensaje.Body.ToString.IndexOf("Impuesto:")
                            pricelgt = taxidx - (priceidx + 7)
                            pricecmpt = mensaje.Body.Substring(priceidx + 7, pricelgt).Trim
                            pricereverse = StrReverse(pricecmpt)

                            lineTotal = StrReverse(pricereverse.Substring(0, pricereverse.IndexOf("$") - 1))
                            price = Format(lineTotal / totalm2, "0.00")

                            dlvrydateidx = mensaje.Body.ToString.IndexOf("Fecha límite de envío:")
                            dlvrydatelgt = itemidx - (dlvrydateidx + 22)
                            dlvrydatecmpt = mensaje.Body.Substring(dlvrydateidx + 22, dlvrydatelgt).Trim

                            statuslgt = skuidx - (statusidx + 7)
                            statuscmpt = mensaje.Body.Substring(statusidx + 7, statuslgt).Trim

                            amazonidx = mensaje.Body.ToString.IndexOf("Cargos de Amazon:")
                            profitsidx = mensaje.Body.ToString.IndexOf("Tus ganancias:")
                            amazonlgt = profitsidx - (amazonidx + 17)
                            amazoncmpt = mensaje.Body.Substring(amazonidx + 17, amazonlgt)

                            profits = profitsidx + 14
                            steps = mensaje.Body.ToString.IndexOf("- - - - - - - - - - - - - - - - - - -")

                            If solditems > 1 Then

                                limit = steps - profits
                                dlvrydateidx = mensaje.Body.Substring(profits, limit).ToString.IndexOf("Fecha límite de envío:")
                                profitscmpt = mensaje.Body.Substring(profits, dlvrydateidx).Trim

                            Else

                                profitslgt = steps - profits
                                profitscmpt = mensaje.Body.Substring(profits, profitslgt).Trim

                            End If

                            stQueryH2 = "Insert Into " & My.Settings.CompanyDB & ".TEMP_Ecommerce values('" & sku & "'," & price & "," & totalm2 & ",'1','" & dlvrydatecmpt & "','" & itemcmpt & "','" & statuscmpt & "','" & qty & "'," & lineTotal & ",'" & amazoncmpt & "','" & profitscmpt & "')"
                            oRecSetH2.DoQuery(stQueryH2)

                            dryidx = mensaje.Body.ToString.IndexOf("Costo del envío:")

                            If dryidx > 0 Then

                                taxdryidx = mensaje.Body.ToString.IndexOf("Impuesto sobre el envío:")
                                drylgt = taxdryidx - (dryidx + 16)
                                drycmpt = mensaje.Body.Substring(dryidx + 16, drylgt).Trim
                                dryreverse = StrReverse(drycmpt)
                                dry = StrReverse(dryreverse.Substring(0, dryreverse.IndexOf("$") - 1))

                                stQueryH2 = "Insert Into " & My.Settings.CompanyDB & ".TEMP_Ecommerce values('SERV'," & dry & ",1,'1','" & dlvrydatecmpt & "','" & itemcmpt & "','" & statuscmpt & "','" & qty & "'," & lineTotal & ",'" & amazoncmpt & "','" & profitscmpt & "')"
                                oRecSetH2.DoQuery(stQueryH2)

                            End If

                            If solditems > 1 Then

                                For x As Integer = 0 To solditems - 2

                                    limit = steps - profits

                                    skuidx = mensaje.Body.Substring(profits, limit).IndexOf("SKU:")
                                    qtyidx = mensaje.Body.Substring(profits, limit).IndexOf("Cantidad:")
                                    skulgt = qtyidx - (skuidx + 4)
                                    skucmpt = mensaje.Body.Substring(profits, limit).ToString.Substring(skuidx + 4, skulgt)
                                    skureverse = StrReverse(skucmpt)
                                    sku = mensaje.Body.Substring(profits, limit).ToString.Substring(skuidx + 4, skulgt - (skureverse.IndexOf("-") + 1)).Trim

                                    dateidx = mensaje.Body.Substring(profits, limit).IndexOf("Fecha del pedido:")
                                    qtylgt = dateidx - (qtyidx + 9)
                                    qty = mensaje.Body.Substring(profits, limit).ToString.Substring(qtyidx + 9, qtylgt).Trim

                                    skulike = sku.Substring(0, 2)

                                    If skulike = "TG" Then

                                        itemidx = mensaje.Body.Substring(profits, limit).IndexOf("Artículo:")
                                        statusidx = mensaje.Body.Substring(profits, limit).IndexOf("Estado:")
                                        itemlgt = statusidx - (itemidx + 9)
                                        itemcmpt = mensaje.Body.Substring(profits, limit).Substring(itemidx + 9, itemlgt)
                                        itemreverse = StrReverse(itemcmpt)

                                        measures = StrReverse(itemreverse.Substring(0, itemreverse.IndexOf("|")))
                                        measure1 = measures.Substring(0, measures.IndexOf("m"))
                                        measuresreverse = StrReverse(measures).Trim
                                        measure2 = StrReverse(measuresreverse.Substring(1, measuresreverse.IndexOf("x") - 1))
                                        m2 = measure1 * measure2

                                        totalm2 = qty * m2

                                    ElseIf skulike = "SG" Then

                                        itemidx = mensaje.Body.Substring(profits, limit).IndexOf("Artículo:")
                                        statusidx = mensaje.Body.Substring(profits, limit).IndexOf("Estado:")
                                        itemlgt = statusidx - (itemidx + 9)
                                        itemcmpt = mensaje.Body.Substring(profits, limit).Substring(itemidx + 9, itemlgt)
                                        itemreverse = StrReverse(itemcmpt)

                                        measures = StrReverse(itemreverse.Substring(0, itemreverse.IndexOf("|")))
                                        measure1 = measures.Substring(0, measures.IndexOf("Piezas")).Trim
                                        m2 = measure1 / 4

                                        totalm2 = qty * m2

                                    Else

                                        'requrido tapetes es por piezas
                                        totalm2 = qty

                                    End If

                                    priceidx = mensaje.Body.Substring(profits, limit).IndexOf("Precio:")
                                    taxidx = mensaje.Body.Substring(profits, limit).IndexOf("Impuesto:")
                                    pricelgt = taxidx - (priceidx + 7)
                                    pricecmpt = mensaje.Body.Substring(profits, limit).Substring(priceidx + 7, pricelgt).Trim
                                    pricereverse = StrReverse(pricecmpt)
                                    lineTotal = StrReverse(pricereverse.Substring(0, pricereverse.IndexOf("$") - 1))
                                    price = Format(lineTotal / totalm2, "0.00")

                                    dlvrydateidx = mensaje.Body.Substring(profits, limit).ToString.IndexOf("Fecha límite de envío:")
                                    dlvrydatelgt = itemidx - (dlvrydateidx + 22)
                                    dlvrydatecmpt = mensaje.Body.Substring(profits, limit).Substring(dlvrydateidx + 22, dlvrydatelgt).Trim

                                    statuslgt = skuidx - (statusidx + 7)
                                    statuscmpt = mensaje.Body.Substring(profits, limit).Substring(statusidx + 7, statuslgt).Trim

                                    amazonidx = mensaje.Body.Substring(profits, limit).ToString.IndexOf("Cargos de Amazon:")
                                    profitsidx = mensaje.Body.Substring(profits, limit).ToString.IndexOf("Tus ganancias:")
                                    amazonlgt = profitsidx - (amazonidx + 17)
                                    amazoncmpt = mensaje.Body.Substring(profits, limit).Substring(amazonidx + 17, amazonlgt)

                                    If x < solditems - 2 Then

                                        profitslbl = profits + mensaje.Body.ToString.Substring(profits, limit).IndexOf("Tus ganancias:") + 14
                                        limit = steps - profitslbl
                                        dlvrydateidx = mensaje.Body.Substring(profitslbl, limit).ToString.IndexOf("Fecha límite de envío:")
                                        profitscmpt = mensaje.Body.Substring(profitslbl, dlvrydateidx).Trim

                                    Else

                                        profitslbl = profits + mensaje.Body.ToString.Substring(profits, limit).IndexOf("Tus ganancias:") + 14
                                        profitslgt = steps - profitslbl
                                        profitscmpt = mensaje.Body.Substring(profitslbl, profitslgt).Trim

                                    End If

                                    stQueryH2 = "Insert Into " & My.Settings.CompanyDB & ".TEMP_Ecommerce values('" & sku & "'," & price & "," & totalm2 & ",'" & x + 2 & "','" & dlvrydatecmpt & "','" & itemcmpt & "','" & statuscmpt & "','" & qty & "'," & lineTotal & ",'" & amazoncmpt & "','" & profitscmpt & "')"
                                    oRecSetH2.DoQuery(stQueryH2)

                                    dryidx = mensaje.Body.Substring(profits, limit).IndexOf("Costo del envío:")

                                    If dryidx > 0 Then

                                        taxdryidx = mensaje.Body.Substring(profits, limit).IndexOf("Impuesto sobre el envío:")
                                        drylgt = taxdryidx - (dryidx + 16)
                                        drycmpt = mensaje.Body.Substring(profits, limit).Substring(dryidx + 16, drylgt).Trim
                                        dryreverse = StrReverse(drycmpt)
                                        dry = StrReverse(dryreverse.Substring(0, dryreverse.IndexOf("$") - 1))

                                        stQueryH2 = "Insert Into " & My.Settings.CompanyDB & ".TEMP_Ecommerce values('SERV'," & dry & ",1,'" & x + 2 & "','" & dlvrydatecmpt & "','" & itemcmpt & "','" & statuscmpt & "','" & qty & "'," & lineTotal & ",'" & amazoncmpt & "','" & profitscmpt & "')"
                                        oRecSetH2.DoQuery(stQueryH2)

                                    End If

                                    profits = profits + mensaje.Body.ToString.Substring(profits, limit).IndexOf("Tus ganancias:") + 14

                                Next


                            End If

                            pathpdf = PDF(order, frommsg, datemsg, header, body)

                            ORDR(order, pathpdf)

                            sku = Nothing
                            price = Nothing
                            totalm2 = Nothing
                            dry = Nothing
                            order = Nothing

                            stQueryH3 = "DELETE FROM " & My.Settings.CompanyDB & ".TEMP_Ecommerce"
                            oRecSetH3.DoQuery(stQueryH3)

                        End If

                    End If

                Next

            Next

        Catch ex As Exception

            stQueryH3 = "DELETE FROM " & My.Settings.CompanyDB & ".TEMP_Ecommerce"
            oRecSetH3.DoQuery(stQueryH3)

            Dim stError As String
            stError = "Error al leer el correo electrónico, ReadEmail. " & ex.Message
            Setlog(stError, order, sku, " ", " ", " ")

        End Try

    End Function


    Public Function PDF(ByVal order As String, ByVal frommsg As String, ByVal datemsg As String, ByVal header As String, ByVal body As String)

        Dim stQueryH, stQueryH2 As String
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim oDoc As New iTextSharp.text.Document(PageSize.A4, 0, 0, 0, 0)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter
        Dim cb As PdfContentByte
        Dim fuente As iTextSharp.text.pdf.BaseFont
        Dim NombreArchivo As String = My.Settings.RutaPDF & order & ".pdf"
        Dim DueDate, Dscription, Status, Item, Pieces, CreateDate, EcmmCharges, Profits, Dscriptionprt, Dscriptionlgt, headerlgt, headerprt, EcmmChargeslgt, EcmmChargesprt As String
        Dim LineTotal, TaxLineTotal As Decimal
        Dim skip As Integer = 35
        Dim y As Integer

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            pdfw = PdfWriter.GetInstance(oDoc, New FileStream(NombreArchivo, FileMode.Create, FileAccess.Write, FileShare.None))
            'Apertura del documento.
            oDoc.Open()
            cb = pdfw.DirectContent
            'Agregamos una pagina.
            oDoc.NewPage()
            'Iniciamos el flujo de bytes.
            cb.BeginText()
            'Instanciamos el objeto para la tipo de letra.
            fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
            'Seteamos el tipo de letra y el tamaño.
            cb.SetFontAndSize(fuente, 10)
            'Seteamos el color del texto a escribir.
            cb.SetColorFill(iTextSharp.text.BaseColor.BLACK)
            'Aqui es donde se escribe el texto.
            'Aclaracion: Por alguna razon la coordenada vertical siempre es tomada desde el borde inferior (de ahi que se calcule como “PageSize.A4.Height – 50″)

            '------------Header
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "From: " & frommsg, 25, PageSize.A4.Height - 25, 0)
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Date: " & datemsg, 25, PageSize.A4.Height - 35, 0)
            skip = skip + 20

            '------------Subject
            If header.Length > 112 Then

                headerlgt = header.Length
                headerprt = header.Substring(0, 112)
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, headerprt, 25, PageSize.A4.Height - skip, 0)
                y = 112

                While (y <= headerlgt)

                    skip = skip + 10
                    If y + 112 < headerlgt Then
                        headerprt = header.Substring(y, 112)
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, headerprt, 25, PageSize.A4.Height - skip, 0)
                    Else
                        headerprt = header.Substring(y, headerlgt - y)
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, headerprt, 25, PageSize.A4.Height - skip, 0)
                    End If
                    y = y + 112

                End While

            Else

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, header, 25, PageSize.A4.Height - skip, 0)

            End If

            '------------Body
            skip = skip + 20
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Número de pedido: " & order, 25, PageSize.A4.Height - skip, 0)
            cb.SetFontAndSize(fuente, 8)

            stQueryH = "Select ""SOLDITEM"" from TEMP_Ecommerce group by ""SOLDITEM"""
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                For x As Integer = 0 To oRecSetH.RecordCount - 1

                    stQueryH2 = "Select * from TEMP_Ecommerce where ""SOLDITEM""='" & x + 1 & "'"
                    oRecSetH2.DoQuery(stQueryH2)

                    If oRecSetH2.RecordCount > 0 Then

                        oRecSetH2.MoveFirst()

                        DueDate = oRecSetH2.Fields.Item("DueDate").Value
                        skip = skip + 20
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Fecha Límite de envío: " & DueDate, 25, PageSize.A4.Height - skip, 0)
                        Dscription = oRecSetH2.Fields.Item("Dscription").Value
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Artículo: ", 25, PageSize.A4.Height - skip, 0)

                        If Dscription.Length > 143 Then

                            Dscriptionlgt = Dscription.Length
                            Dscriptionprt = Dscription.Substring(0, 143)
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Dscriptionprt, 53, PageSize.A4.Height - skip, 0)
                            y = 143

                            While (y <= Dscriptionlgt)

                                skip = skip + 10
                                If y + 143 < Dscriptionlgt Then
                                    Dscriptionprt = Dscription.Substring(y, 143)
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Dscriptionprt, 25, PageSize.A4.Height - skip, 0)
                                Else
                                    Dscriptionprt = Dscription.Substring(y, Dscriptionlgt - y)
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Dscriptionprt, 25, PageSize.A4.Height - skip, 0)
                                End If
                                y = y + 143

                            End While

                        Else

                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Dscription, 53, PageSize.A4.Height - skip, 0)

                        End If

                        Status = oRecSetH2.Fields.Item("Status").Value
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Estado: " & Status, 25, PageSize.A4.Height - skip, 0)
                        Item = oRecSetH2.Fields.Item("Item").Value
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SKU: " & Item, 25, PageSize.A4.Height - skip, 0)
                        Pieces = oRecSetH2.Fields.Item("Pieces").Value
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Cantidad: " & Pieces, 25, PageSize.A4.Height - skip, 0)
                        CreateDate = Now.Day & "/" & Now.Month & "/" & Now.Year
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Fecha del pedido: " & CreateDate, 25, PageSize.A4.Height - skip, 0)
                        LineTotal = oRecSetH2.Fields.Item("LineTotal").Value
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Precio: " & LineTotal, 25, PageSize.A4.Height - skip, 0)
                        TaxLineTotal = LineTotal * 0.16
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Impuesto: " & TaxLineTotal, 25, PageSize.A4.Height - skip, 0)

                        '-- Brinco a la segunda linea para el envio e impuesto
                        If oRecSetH2.RecordCount = 2 Then

                            oRecSetH2.MoveNext()
                            LineTotal = oRecSetH2.Fields.Item("Price").Value
                            skip = skip + 10
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Costo del envío: " & LineTotal, 25, PageSize.A4.Height - skip, 0)
                            TaxLineTotal = LineTotal * 0.16
                            skip = skip + 10
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Impuesto sobre el envío: " & TaxLineTotal, 25, PageSize.A4.Height - skip, 0)

                        End If

                        EcmmCharges = oRecSetH2.Fields.Item("AmazonCharges").Value
                        skip = skip + 10
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Cargos de Amazon: ", 25, PageSize.A4.Height - skip, 0)

                        If EcmmCharges.Length > 122 Then

                            EcmmChargeslgt = EcmmCharges.Length
                            EcmmChargesprt = EcmmCharges.Substring(0, 122)
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, EcmmChargesprt, 95, PageSize.A4.Height - skip, 0)
                            y = 122

                            While (y <= EcmmChargeslgt)

                                skip = skip + 10
                                If y + 122 < EcmmChargeslgt Then
                                    EcmmChargesprt = EcmmCharges.Substring(y, 122)
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, EcmmChargesprt, 25, PageSize.A4.Height - skip, 0)
                                Else
                                    EcmmChargesprt = EcmmCharges.Substring(y, EcmmChargeslgt - y)
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, EcmmChargesprt, 25, PageSize.A4.Height - skip, 0)
                                End If
                                y = y + 122

                            End While

                        Else

                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, EcmmCharges, 95, PageSize.A4.Height - skip, 0)

                        End If

                        Profits = oRecSetH2.Fields.Item("Profits").Value
                        skip = skip + 10
                        cb.SetFontAndSize(fuente, 8)
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Tus ganancias: " & Profits, 25, PageSize.A4.Height - skip, 0)

                    End If

                Next

            End If

            'Fin del flujo de bytes.
            cb.EndText()
            'Forzamos vaciamiento del buffer.
            pdfw.Flush()
            'Cerramos el documento.
            oDoc.Close()

            Return NombreArchivo

        Catch ex As Exception

            'Si hubo una excepcion y el archivo existe …
            If File.Exists(NombreArchivo) Then
                'Cerramos el documento si esta abierto.
                'Y asi desbloqueamos el archivo para su eliminacion.
                If oDoc.IsOpen Then oDoc.Close()
                '… lo eliminamos de disco.
                File.Delete(NombreArchivo)
            End If

            Dim stError As String
            stError = "Error al crear el pdf, PDF. " & ex.Message
            Setlog(stError, order, Item, NombreArchivo, " ", " ")

        Finally
            cb = Nothing
            pdfw = Nothing
            oDoc = Nothing
        End Try

    End Function


    Public Function ORDR(ByVal order As String, ByVal pathpdf As String)

        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim oORDR As SAPbobsCOM.Documents
        Dim llError As Long
        Dim lsError As String
        Dim dayw, sku, OrderSAP As String
        Dim addday, price As Decimal
        Dim totalm2 As Double
        Dim deliverydate As Date

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            oORDR = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

            dayw = Now.DayOfWeek
            If dayw = 5 Then

                addday = 3

            ElseIf dayw = 6 Then

                addday = 2

            Else

                addday = 1

            End If

            deliverydate = DateAdd("d", addday, Now.Date)

            oORDR.Series = 9
            oORDR.CardCode = "XAXX010101002"
            oORDR.DocDate = Year(Now.Date).ToString + "-0" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
            oORDR.DocDueDate = Year(deliverydate).ToString + "-" + Month(deliverydate).ToString + "-" + Day(deliverydate).ToString
            oORDR.SalesPersonCode = 30
            oORDR.DocumentsOwner = 60
            oORDR.UserFields.Fields.Item("U_B1SYS_MainUsage").Value = "G01"

            oORDR.TransportationCode = 1 '--Forma de envio depende del articulo

            oORDR.NumAtCard = order
            oORDR.UserFields.Fields.Item("U_Comprobante").Value = pathpdf

            '0 = Items   &     1 = Services
            oORDR.DocType = 0

            stQueryH = "Select * from TEMP_Ecommerce"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                For x As Integer = 0 To oRecSetH.RecordCount - 1

                    sku = oRecSetH.Fields.Item("Item").Value
                    price = oRecSetH.Fields.Item("Price").Value
                    totalm2 = oRecSetH.Fields.Item("Quantity").Value

                    stQueryH2 = "Select ""ItemName"" from OITM where ""ItemCode""='" & sku & "'"
                    oRecSetH2.DoQuery(stQueryH2)

                    If oRecSetH2.RecordCount > 0 Then

                        oRecSetH2.MoveFirst()

                        oORDR.Lines.ItemCode = sku 'sku
                        oORDR.Lines.ItemDescription = oRecSetH2.Fields.Item("ItemName").Value
                        oORDR.Lines.UnitPrice = price 'precio unitario
                        oORDR.Lines.Quantity = totalm2 'cantidad dependiendo del articulo
                        oORDR.Lines.TaxCode = "IVAP16"
                        oORDR.Lines.WarehouseCode = "001"
                        oORDR.Lines.ProjectCode = "001"
                        oORDR.Lines.Currency = "MXN"
                        oORDR.Lines.UserFields.Fields.Item("U_NumPaq").Value = totalm2
                        oORDR.Lines.Add()

                    Else

                        Dim stError As String
                        stError = "Error en sku " & sku & ", no existe. ORDR"
                        Setlog(stError, order, sku, pathpdf, " ", " ")

                    End If

                    oRecSetH.MoveNext()

                Next

            End If

            If oORDR.Add() <> 0 Then

                SBOCompany.GetLastError(llError, lsError)
                Dim stError As String
                stError = "Error al crear orden de venta, ORDR. " & llError & " " & lsError
                Setlog(stError, order, sku, pathpdf, " ", " ")

            Else

                stQueryH3 = "Select ""DocNum"" from ORDR where ""NumAtCard""='" & order & "'"
                oRecSetH3.DoQuery(stQueryH3)

                If oRecSetH3.RecordCount > 1 Then

                    oRecSetH3.MoveFirst()

                    OrderSAP = oRecSetH3.Fields.Item("DocNum").Value
                    SendEmail(OrderSAP, order, pathpdf, sku)

                End If

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error al crear la orden de venta, ORDR. " & ex.Message
            Setlog(stError, order, sku, pathpdf, OrderSAP, " ")

        End Try

    End Function


    Public Function SendEmail(ByVal OrderSAP As String, ByVal order As String, ByVal pathpdf As String, ByVal sku As String)

        'MsgBox("Validacion de Documentos exitosa")
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim oRecSettxb, oRecSettxb2 As SAPbobsCOM.Recordset
        Dim stQuerytxb, stQuerytxb2 As String
        Dim EmailU, Pass, EmailC, EmailCC, Subject, Body, smtpService, Puerto, SegSSL As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Select * from ""@CORREOTEKNO"" where ""Name""='Ecommerce'"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                EmailU = oRecSettxb.Fields.Item("U_Email").Value
                Pass = oRecSettxb.Fields.Item("U_Password").Value

                Subject = oRecSettxb.Fields.Item("U_Subject").Value
                Body = oRecSettxb.Fields.Item("U_Body").Value
                smtpService = oRecSettxb.Fields.Item("U_SMTP").Value
                Puerto = oRecSettxb.Fields.Item("U_Puerto").Value
                SegSSL = oRecSettxb.Fields.Item("U_SeguridadSSL").Value

                'Limpiamos correo destinatario, correo copia y archivos adjuntos
                message.To.Clear()
                message.CC.Clear()
                message.Attachments.Clear()

                'Llenamos encabezado de correo
                message.From = New MailAddress(EmailU)

                EmailC = ArreglarTexto(My.Settings.EmailC, ";", ",")
                message.To.Add(EmailC)

                EmailCC = ArreglarTexto(My.Settings.EmailCC, ";", ",")
                message.CC.Add(EmailCC)

                message.Subject = Subject & " " & OrderSAP

                'Llenamos el cuerpo del correo y prioridad
                message.Body = "Se creo la orden de venta " & OrderSAP & " en SAP B1 basado en la orden de Amazon (" & order & "), se adjunta formato en pdf." & Body
                message.Priority = MailPriority.Normal

                'Adjuntamos archivos pdf
                Dim attpdf As New Net.Mail.Attachment(pathpdf)
                message.Attachments.Add(attpdf)

                'Llenamos datos de smtp
                smtp.Host = smtpService
                smtp.Credentials = New Net.NetworkCredential(EmailU, Pass)
                smtp.Port = Puerto
                smtp.EnableSsl = SegSSL

                'Enviamos Correo
                smtp.Send(message)

                oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb2 = "Update ORDR set ""U_PDF""='" & pathpdf & "' where ""DocNum""=" & OrderSAP
                oRecSettxb2.DoQuery(stQuerytxb2)

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error al enviar el correo electrónico, SendEmail. " & ex.Message
            Setlog(stError, order, sku, pathpdf, OrderSAP, EmailC)

        End Try

    End Function


    Public Function Setlog(ByVal stError As String, ByVal OrderEcomm As String, ByVal Sku As String, ByVal Path As String, ByVal OrderSAP As String, ByVal UserSAP As String)

        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String

        Try

            stError = ArreglarTexto(stError, "'", " ")
            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Insert Into " & My.Settings.CompanyDB & ".LOG_Ecommerce values ('" & OrderEcomm & "','" & Sku & "','" & Path & "','" & OrderSAP & "','" & UserSAP & "','" & stError & "',current_date)"
            oRecSettxb.DoQuery(stQuerytxb)

        Catch ex As Exception

            'MsgBox(stError)

        End Try

    End Function


    Public Function ArreglarTexto(ByVal TextoOriginal As String, ByVal QuitarCaracter As String, ByVal PonerCaracter As String)

        TextoOriginal = TextoOriginal.Replace(QuitarCaracter, PonerCaracter)
        Return TextoOriginal

    End Function


    Public Function Disconnect()

        Try

            SBOCompany.Disconnect()

        Catch ex As Exception

            Dim stError As String
            stError = "Error al tratar de cerrar la conexión con SAP B1. " & ex.Message
            Setlog(stError, " ", " ", " ", " ", " ")

        End Try

    End Function

End Module
