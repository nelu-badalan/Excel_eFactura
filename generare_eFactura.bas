Attribute VB_Name = "generare_eFactura"
'==============================================================================
' generare_eFactura.bas - Generator XML eFactura ANAF (CIUS-RO 1.0.1)
'
' Versiune:    3.0.0
' Licenta:     MIT
' Repository:  https://github.com/<USER>/<REPO>          (completeaza dupa push)
' Articol:     https://conta-si-abilitate.ro/<SLUG>      (completeaza dupa publicare)
'
' Genereaza fisiere XML conforme UBL Invoice 2.1 / CIUS-RO 1.0.1, gata pentru
' incarcare manuala in SPV-ul ANAF (e-Factura).
'
' Acoperire:
'   - B2B (firma catre firma) si B2C (firma catre persoana fizica)
'   - Furnizor platitor sau neplatitor TVA (auto din coloana PlatitorTVA)
'   - Client persoana juridica, PFA cu CNP platitor TVA, sau persoana fizica
'   - Anonimizare CNP la 13 zerouri pentru B2C consumator (conform GDPR)
'   - Storno: tip 380 cu cantitate negativa pe linie (pret ramane pozitiv)
'
' Limitari v1:
'   - Doar moneda RON
'   - Doar tip factura 380
'   - Cote TVA suportate: 0 / 5 / 11 / 21 (categoria S), plus O (neplatitor)
'==============================================================================

Option Explicit

' === Constante ===
Private Const FILE_VERSION   As String = "3.0.0"
Private Const CIUS_RO_ID     As String = "urn:cen.eu:en16931:2017#compliant#urn:efactura.mfinante.ro:CIUS-RO:1.0.1"
Private Const PEPPOL_PROFILE As String = "urn:fdc:peppol.eu:2017:poacc:billing:3.0"
Private Const RECONCILE_TOL  As Double = 0.01
Private Const ANAF_VALIDATOR As String = "https://www.anaf.ro/uploadxml/"
Private Const EXEMPT_DEFAULT As String = "Entitatea nu este inregistrata in scopuri de TVA"

' === Coloane Furnizori (1-based) ===
Private Const F_REGNAME      As Long = 1
Private Const F_COMPID       As Long = 2
Private Const F_STREET       As Long = 3
Private Const F_CITY         As Long = 4
Private Const F_POSTALZONE   As Long = 5
Private Const F_COUNTRY      As Long = 6
Private Const F_ADDSTREET    As Long = 7
Private Const F_SUBENTITY    As Long = 8
Private Const F_PHONE        As Long = 9
Private Const F_EMAIL        As Long = 10
Private Const F_IBAN         As Long = 11
Private Const F_PAYCODE      As Long = 12
Private Const F_REGCODE      As Long = 13
Private Const F_LEGALFORM    As Long = 14
Private Const F_VATPAYER     As Long = 15  ' coloana noua: DA / NU
Private Const F_EXEMPTREASON As Long = 16  ' coloana noua: motiv scutire
Private Const F_BANKNAME     As Long = 17  ' coloana noua: nume banca

' === Coloane Clienti ===
Private Const C_REGNAME    As Long = 1
Private Const C_COMPID     As Long = 2
Private Const C_STREET     As Long = 3
Private Const C_CITY       As Long = 4
Private Const C_POSTALZONE As Long = 5
Private Const C_COUNTRY    As Long = 6
Private Const C_ADDSTREET  As Long = 7
Private Const C_SUBENTITY  As Long = 8
Private Const C_PHONE      As Long = 9
Private Const C_EMAIL      As Long = 10
Private Const C_ID         As Long = 11

' === Coloane Facturi_Antet ===
Private Const FA_SUPPLIERID  As Long = 1
Private Const FA_CUSTOMERID  As Long = 2
Private Const FA_INVOICEID   As Long = 3
Private Const FA_ISSUEDATE   As Long = 4
Private Const FA_DUEDATE     As Long = 5
Private Const FA_TAXPOINTDT  As Long = 6
Private Const FA_TYPECODE    As Long = 7
Private Const FA_DOCCURRENCY As Long = 8
Private Const FA_TAXCURRENCY As Long = 9
Private Const FA_NOTE        As Long = 10
Private Const FA_PAYABLE     As Long = 11
Private Const FA_LINEEXT     As Long = 12
Private Const FA_TAXEXCL     As Long = 13
Private Const FA_TAXINCL     As Long = 14
Private Const FA_ALLOWANCE   As Long = 15
Private Const FA_CHARGE      As Long = 16
Private Const FA_VAT         As Long = 17
Private Const FA_PAYCODE     As Long = 18
Private Const FA_COUNTLINE   As Long = 19
Private Const FA_ID          As Long = 20

' === Coloane Linii_Facturi ===
Private Const L_INVOICE   As Long = 1
Private Const L_LINEID    As Long = 2
Private Const L_QTY       As Long = 3
Private Const L_UNITCODE  As Long = 4
Private Const L_DESC      As Long = 5
Private Const L_NOTE      As Long = 6
Private Const L_PRICE     As Long = 7
Private Const L_CURRENCY  As Long = 8
Private Const L_LINEEXT   As Long = 9
Private Const L_TAXPCT    As Long = 10
Private Const L_TAXCAT    As Long = 11
Private Const L_TAXSCH    As Long = 12
Private Const L_VAT       As Long = 13
Private Const L_LINETOTAL As Long = 14

'==============================================================================
' MAIN - genereaza XML doar pentru factura de pe randul activ din 'Facturi_Antet'
'==============================================================================
Public Sub Genereaza_eFactura_XML_Excel()
Attribute Genereaza_eFactura_XML_Excel.VB_Description = "Genereaza xml"
Attribute Genereaza_eFactura_XML_Excel.VB_ProcData.VB_Invoke_Func = "X\n14"
    On Error GoTo ErrHandler

    Dim wsF As Worksheet, wsC As Worksheet, wsA As Worksheet, wsL As Worksheet
    If Not GetSheets(wsF, wsC, wsA, wsL) Then Exit Sub

    ' Cursorul trebuie sa fie pe foaia 'Facturi_Antet'
    If ActiveSheet.Name <> wsA.Name Then
        MsgBox "Muta cursorul pe foaia 'Facturi_Antet', pe randul facturii pe care vrei sa o generezi, apoi ruleaza din nou.", _
               vbExclamation, "eFactura - v" & FILE_VERSION
        Exit Sub
    End If

    Dim antetRow As Long
    antetRow = ActiveCell.row
    If antetRow < 2 Then
        MsgBox "Selecteaza un rand de factura (randul 1 contine antetul de coloane).", _
               vbExclamation, "eFactura - v" & FILE_VERSION
        Exit Sub
    End If

    Dim invID As String, supId As String, custId As String
    invID = Trim$(GetCellStr(wsA, antetRow, FA_INVOICEID))
    supId = Trim$(GetCellStr(wsA, antetRow, FA_SUPPLIERID))
    custId = Trim$(GetCellStr(wsA, antetRow, FA_CUSTOMERID))

    If invID = "" Then
        MsgBox "Randul " & antetRow & " nu are InvoiceID completat. Selecteaza un rand valid.", _
               vbExclamation, "eFactura - v" & FILE_VERSION
        Exit Sub
    End If

    Dim lastF As Long, lastC As Long, lastL As Long
    lastF = LastUsedRow(wsF)
    lastC = LastUsedRow(wsC)
    lastL = LastUsedRow(wsL)

    Dim furnRow As Long, clieRow As Long
    furnRow = FindRow(wsF, lastF, F_COMPID, supId)
    If furnRow = 0 Then
        MsgBox "Furnizorul cu ID '" & supId & "' nu a fost gasit in foaia 'Furnizori'.", _
               vbCritical, "eFactura - v" & FILE_VERSION
        Exit Sub
    End If
    clieRow = FindRow(wsC, lastC, C_COMPID, custId)
    If clieRow = 0 Then
        MsgBox "Clientul cu ID '" & custId & "' nu a fost gasit in foaia 'Clienti'.", _
               vbCritical, "eFactura - v" & FILE_VERSION
        Exit Sub
    End If

    Dim reconErr As String
    reconErr = ReconcileInvoice(wsA, antetRow, wsL, lastL, invID)
    If reconErr <> "" Then
        MsgBox "Factura " & invID & " - sume incoerente intre antet si linii:" & vbCrLf & reconErr, _
               vbCritical, "eFactura - v" & FILE_VERSION
        Exit Sub
    End If

    Dim outputDir As String
    outputDir = ThisWorkbook.path & Application.PathSeparator & "xml"
    If Dir(outputDir, vbDirectory) = "" Then MkDir outputDir

    Dim xml As String
    xml = BuildInvoiceXml(wsA, antetRow, wsF, furnRow, wsC, clieRow, wsL, lastL, invID)

    Dim fileName As String, supForFn As String
    supForFn = StripROPrefix(supId)
    fileName = outputDir & Application.PathSeparator & supForFn & "_" & SanitizeForFilename(invID) & ".xml"
    WriteUtf8 fileName, xml

    MsgBox "Factura " & invID & " generata:" & vbCrLf & fileName & vbCrLf & vbCrLf & _
           "Valideaza la ANAF (copiaza in browser):" & vbCrLf & ANAF_VALIDATOR, _
           vbInformation, "eFactura - v" & FILE_VERSION
    Exit Sub

ErrHandler:
    MsgBox "Eroare neasteptata (rand " & antetRow & ", cod " & Err.Number & "): " & Err.Description, _
           vbCritical, "eFactura - v" & FILE_VERSION
End Sub

'==============================================================================
' XML BUILDERS
'==============================================================================

Private Function BuildInvoiceXml(wsA As Worksheet, antetRow As Long, _
                                  wsF As Worksheet, furnRow As Long, _
                                  wsC As Worksheet, clieRow As Long, _
                                  wsL As Worksheet, lastL As Long, _
                                  invID As String) As String
    Dim sb As String
    Dim invType As String, issueDate As String, dueDate As String
    Dim docCcy As String, taxCcy As String, note As String, payCode As String
    Dim isVATSupplier As Boolean
    Dim exemptReason As String

    invType = NzStr(GetCellStr(wsA, antetRow, FA_TYPECODE), "380")
    issueDate = FormatIso(wsA.Cells(antetRow, FA_ISSUEDATE).Value)
    dueDate = FormatIso(wsA.Cells(antetRow, FA_DUEDATE).Value)
    If dueDate = "" Then dueDate = issueDate
    docCcy = NzStr(GetCellStr(wsA, antetRow, FA_DOCCURRENCY), "RON")
    taxCcy = NzStr(GetCellStr(wsA, antetRow, FA_TAXCURRENCY), "RON")
    note = GetCellStr(wsA, antetRow, FA_NOTE)
    payCode = GetCellStr(wsA, antetRow, FA_PAYCODE)
    isVATSupplier = IsVATPayerFlag(GetCellStr(wsF, furnRow, F_VATPAYER))
    exemptReason = NzStr(GetCellStr(wsF, furnRow, F_EXEMPTREASON), EXEMPT_DEFAULT)

    sb = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    sb = sb & "<Invoice xmlns=""urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"""
    sb = sb & " xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"""
    sb = sb & " xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"">" & vbCrLf
    sb = sb & "  <cbc:CustomizationID>" & CIUS_RO_ID & "</cbc:CustomizationID>" & vbCrLf
    sb = sb & "  <cbc:ProfileID>" & PEPPOL_PROFILE & "</cbc:ProfileID>" & vbCrLf
    sb = sb & "  <cbc:ID>" & XmlEscape(invID) & "</cbc:ID>" & vbCrLf
    sb = sb & "  <cbc:IssueDate>" & issueDate & "</cbc:IssueDate>" & vbCrLf
    sb = sb & "  <cbc:DueDate>" & dueDate & "</cbc:DueDate>" & vbCrLf
    sb = sb & "  <cbc:InvoiceTypeCode>" & XmlEscape(invType) & "</cbc:InvoiceTypeCode>" & vbCrLf
    If note <> "" Then sb = sb & "  <cbc:Note>" & XmlEscape(note) & "</cbc:Note>" & vbCrLf
    sb = sb & "  <cbc:DocumentCurrencyCode>" & XmlEscape(docCcy) & "</cbc:DocumentCurrencyCode>" & vbCrLf
    sb = sb & "  <cbc:TaxCurrencyCode>" & XmlEscape(taxCcy) & "</cbc:TaxCurrencyCode>" & vbCrLf

    sb = sb & BuildSupplierBlock(wsF, furnRow, isVATSupplier)
    sb = sb & BuildCustomerBlock(wsC, clieRow, isVATSupplier)
    sb = sb & BuildPaymentBlock(wsF, furnRow, payCode)
    sb = sb & BuildTaxTotalBlock(wsA, antetRow, wsL, lastL, invID, isVATSupplier, exemptReason)
    sb = sb & BuildLegalMonetaryBlock(wsA, antetRow)
    sb = sb & BuildAllInvoiceLines(wsL, lastL, invID, isVATSupplier)
    sb = sb & "</Invoice>"

    BuildInvoiceXml = sb
End Function

Private Function BuildSupplierBlock(ws As Worksheet, row As Long, isVAT As Boolean) As String
    Dim sb As String
    Dim regName As String, taxId As String, regCode As String, legalForm As String
    Dim street As String, city As String, country As String, subentity As String

    regName = GetCellStr(ws, row, F_REGNAME)
    taxId = Trim$(GetCellStr(ws, row, F_COMPID))
    street = GetCellStr(ws, row, F_STREET)
    city = GetCellStr(ws, row, F_CITY)
    country = NzStr(GetCellStr(ws, row, F_COUNTRY), "RO")
    subentity = GetCellStr(ws, row, F_SUBENTITY)
    regCode = GetCellStr(ws, row, F_REGCODE)
    legalForm = GetCellStr(ws, row, F_LEGALFORM)

    Dim taxIdNoRO As String
    taxIdNoRO = StripROPrefix(taxId)

    sb = "  <cac:AccountingSupplierParty>" & vbCrLf
    sb = sb & "    <cac:Party>" & vbCrLf
    sb = sb & "      <cac:PartyIdentification><cbc:ID>" & XmlEscape(taxIdNoRO) & "</cbc:ID></cac:PartyIdentification>" & vbCrLf
    sb = sb & BuildAddress(street, city, subentity, country)
    sb = sb & "      <cac:PartyTaxScheme>" & vbCrLf
    If isVAT Then
        sb = sb & "        <cbc:CompanyID>RO" & XmlEscape(taxIdNoRO) & "</cbc:CompanyID>" & vbCrLf
        sb = sb & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
    Else
        sb = sb & "        <cbc:CompanyID>" & XmlEscape(taxIdNoRO) & "</cbc:CompanyID>" & vbCrLf
        sb = sb & "        <cac:TaxScheme/>" & vbCrLf
    End If
    sb = sb & "      </cac:PartyTaxScheme>" & vbCrLf
    sb = sb & "      <cac:PartyLegalEntity>" & vbCrLf
    sb = sb & "        <cbc:RegistrationName>" & XmlEscape(regName) & "</cbc:RegistrationName>" & vbCrLf
    If regCode <> "" Then sb = sb & "        <cbc:CompanyID>" & XmlEscape(regCode) & "</cbc:CompanyID>" & vbCrLf
    If legalForm <> "" Then sb = sb & "        <cbc:CompanyLegalForm>" & XmlEscape(legalForm) & "</cbc:CompanyLegalForm>" & vbCrLf
    sb = sb & "      </cac:PartyLegalEntity>" & vbCrLf
    sb = sb & "    </cac:Party>" & vbCrLf
    sb = sb & "  </cac:AccountingSupplierParty>" & vbCrLf

    BuildSupplierBlock = sb
End Function

Private Function BuildCustomerBlock(ws As Worksheet, row As Long, isVATSupplier As Boolean) As String
    Dim sb As String
    Dim regName As String, custId As String
    Dim street As String, city As String, country As String, subentity As String

    regName = GetCellStr(ws, row, C_REGNAME)
    custId = Trim$(GetCellStr(ws, row, C_COMPID))
    street = GetCellStr(ws, row, C_STREET)
    city = GetCellStr(ws, row, C_CITY)
    country = NzStr(GetCellStr(ws, row, C_COUNTRY), "RO")
    subentity = GetCellStr(ws, row, C_SUBENTITY)

    Dim custIdNoRO As String, hasRO As Boolean, isB2C As Boolean
    hasRO = HasROPrefix(custId)
    custIdNoRO = StripROPrefix(custId)
    ' Client B2C consumator = CNP fara prefix RO -> anonimizare la 13 zerouri
    isB2C = (Not hasRO) And IsCNP(custIdNoRO)

    Dim displayId As String
    If isB2C Then displayId = "0000000000000" Else displayId = custIdNoRO

    ' BR-O-02: cand furnizorul este neplatitor TVA, toate liniile sunt categoria O
    ' si NU avem voie sa emitem identificator VAT pentru client (BT-48), chiar daca
    ' clientul este in realitate platitor TVA (are prefix RO). Deci VAT TaxScheme se
    ' emite doar cand AMBII (furnizor + client) sunt platitori.
    Dim emitVATScheme As Boolean
    emitVATScheme = isVATSupplier And hasRO

    sb = "  <cac:AccountingCustomerParty>" & vbCrLf
    sb = sb & "    <cac:Party>" & vbCrLf
    sb = sb & "      <cac:PartyIdentification><cbc:ID>" & XmlEscape(displayId) & "</cbc:ID></cac:PartyIdentification>" & vbCrLf
    sb = sb & BuildAddress(street, city, subentity, country)
    sb = sb & "      <cac:PartyTaxScheme>" & vbCrLf
    If emitVATScheme Then
        sb = sb & "        <cbc:CompanyID>RO" & XmlEscape(custIdNoRO) & "</cbc:CompanyID>" & vbCrLf
        sb = sb & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
    Else
        ' TaxScheme gol = fara identificator VAT (conform BR-O-02 si pattern-ului ANAF)
        sb = sb & "        <cbc:CompanyID>" & XmlEscape(displayId) & "</cbc:CompanyID>" & vbCrLf
        sb = sb & "        <cac:TaxScheme/>" & vbCrLf
    End If
    sb = sb & "      </cac:PartyTaxScheme>" & vbCrLf
    sb = sb & "      <cac:PartyLegalEntity>" & vbCrLf
    sb = sb & "        <cbc:RegistrationName>" & XmlEscape(regName) & "</cbc:RegistrationName>" & vbCrLf
    If isB2C Then
        sb = sb & "        <cbc:CompanyID>" & displayId & "</cbc:CompanyID>" & vbCrLf
    ElseIf Not emitVATScheme Then
        ' Pentru orice caz care nu este B2B clasic platitor->platitor, includem CompanyID fara RO
        sb = sb & "        <cbc:CompanyID>" & XmlEscape(custIdNoRO) & "</cbc:CompanyID>" & vbCrLf
    End If
    ' Pentru hasRO + isVATSupplier=true (B2B clasic), nu emitem CompanyID in PartyLegalEntity decat daca avem nr. ORC.
    sb = sb & "      </cac:PartyLegalEntity>" & vbCrLf
    sb = sb & "    </cac:Party>" & vbCrLf
    sb = sb & "  </cac:AccountingCustomerParty>" & vbCrLf

    BuildCustomerBlock = sb
End Function

Private Function BuildAddress(street As String, city As String, subentity As String, country As String) As String
    Dim sb As String
    sb = "      <cac:PostalAddress>" & vbCrLf
    sb = sb & "        <cbc:StreetName>" & XmlEscape(street) & "</cbc:StreetName>" & vbCrLf
    sb = sb & "        <cbc:CityName>" & XmlEscape(city) & "</cbc:CityName>" & vbCrLf
    sb = sb & "        <cbc:CountrySubentity>" & XmlEscape(subentity) & "</cbc:CountrySubentity>" & vbCrLf
    sb = sb & "        <cac:Country><cbc:IdentificationCode>" & XmlEscape(country) & "</cbc:IdentificationCode></cac:Country>" & vbCrLf
    sb = sb & "      </cac:PostalAddress>" & vbCrLf
    BuildAddress = sb
End Function

Private Function BuildPaymentBlock(wsF As Worksheet, furnRow As Long, payCode As String) As String
    Dim iban As String, bankName As String
    iban = GetCellStr(wsF, furnRow, F_IBAN)
    bankName = GetCellStr(wsF, furnRow, F_BANKNAME)

    If payCode = "" And iban = "" Then Exit Function

    Dim sb As String
    sb = "  <cac:PaymentMeans>" & vbCrLf
    sb = sb & "    <cbc:PaymentMeansCode>" & XmlEscape(NzStr(payCode, "42")) & "</cbc:PaymentMeansCode>" & vbCrLf
    If iban <> "" Then
        sb = sb & "    <cac:PayeeFinancialAccount>" & vbCrLf
        sb = sb & "      <cbc:ID>" & XmlEscape(iban) & "</cbc:ID>" & vbCrLf
        If bankName <> "" Then sb = sb & "      <cbc:Name>" & XmlEscape(bankName) & "</cbc:Name>" & vbCrLf
        sb = sb & "    </cac:PayeeFinancialAccount>" & vbCrLf
    End If
    sb = sb & "  </cac:PaymentMeans>" & vbCrLf
    BuildPaymentBlock = sb
End Function

Private Function BuildTaxTotalBlock(wsA As Worksheet, antetRow As Long, _
                                     wsL As Worksheet, lastL As Long, _
                                     invID As String, isVATSupplier As Boolean, _
                                     exemptReason As String) As String
    Dim sb As String

    If Not isVATSupplier Then
        ' Neplatitor: o singura subtotal cu categorie O
        Dim taxableSum As Double
        taxableSum = SumLineExt(wsL, lastL, invID)
        sb = "  <cac:TaxTotal>" & vbCrLf
        sb = sb & "    <cbc:TaxAmount currencyID=""RON"">0.00</cbc:TaxAmount>" & vbCrLf
        sb = sb & "    <cac:TaxSubtotal>" & vbCrLf
        sb = sb & "      <cbc:TaxableAmount currencyID=""RON"">" & FormatAmount(taxableSum) & "</cbc:TaxableAmount>" & vbCrLf
        sb = sb & "      <cbc:TaxAmount currencyID=""RON"">0.00</cbc:TaxAmount>" & vbCrLf
        sb = sb & "      <cac:TaxCategory>" & vbCrLf
        sb = sb & "        <cbc:ID>O</cbc:ID>" & vbCrLf
        sb = sb & "        <cbc:TaxExemptionReasonCode>VATEX-EU-O</cbc:TaxExemptionReasonCode>" & vbCrLf
        sb = sb & "        <cbc:TaxExemptionReason>" & XmlEscape(exemptReason) & "</cbc:TaxExemptionReason>" & vbCrLf
        sb = sb & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
        sb = sb & "      </cac:TaxCategory>" & vbCrLf
        sb = sb & "    </cac:TaxSubtotal>" & vbCrLf
        sb = sb & "  </cac:TaxTotal>" & vbCrLf
        BuildTaxTotalBlock = sb
        Exit Function
    End If

    ' Platitor: agregam pe cote distincte
    Dim rates() As Double, exclSums() As Double, vatSums() As Double, nRates As Long
    nRates = 0
    ReDim rates(0 To 9), exclSums(0 To 9), vatSums(0 To 9)

    Dim j As Long
    For j = 2 To lastL
        If Trim$(GetCellStr(wsL, j, L_INVOICE)) = invID Then
            Dim pct As Double, lex As Double, vat As Double
            pct = CDblSafe(wsL.Cells(j, L_TAXPCT).Value)
            lex = CDblSafe(wsL.Cells(j, L_LINEEXT).Value)
            vat = Round(lex * pct / 100#, 2)

            Dim k As Long, found As Boolean
            found = False
            For k = 0 To nRates - 1
                If Abs(rates(k) - pct) < 0.001 Then
                    exclSums(k) = exclSums(k) + lex
                    vatSums(k) = vatSums(k) + vat
                    found = True
                    Exit For
                End If
            Next k
            If Not found Then
                rates(nRates) = pct
                exclSums(nRates) = lex
                vatSums(nRates) = vat
                nRates = nRates + 1
            End If
        End If
    Next j

    Dim totalVat As Double
    For j = 0 To nRates - 1
        totalVat = totalVat + vatSums(j)
    Next j

    sb = "  <cac:TaxTotal>" & vbCrLf
    sb = sb & "    <cbc:TaxAmount currencyID=""RON"">" & FormatAmount(totalVat) & "</cbc:TaxAmount>" & vbCrLf
    For j = 0 To nRates - 1
        sb = sb & "    <cac:TaxSubtotal>" & vbCrLf
        sb = sb & "      <cbc:TaxableAmount currencyID=""RON"">" & FormatAmount(exclSums(j)) & "</cbc:TaxableAmount>" & vbCrLf
        sb = sb & "      <cbc:TaxAmount currencyID=""RON"">" & FormatAmount(vatSums(j)) & "</cbc:TaxAmount>" & vbCrLf
        sb = sb & "      <cac:TaxCategory>" & vbCrLf
        If rates(j) <= 0.001 Then
            sb = sb & "        <cbc:ID>Z</cbc:ID>" & vbCrLf
            sb = sb & "        <cbc:Percent>0.00</cbc:Percent>" & vbCrLf
        Else
            sb = sb & "        <cbc:ID>S</cbc:ID>" & vbCrLf
            sb = sb & "        <cbc:Percent>" & FormatAmount(rates(j)) & "</cbc:Percent>" & vbCrLf
        End If
        sb = sb & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
        sb = sb & "      </cac:TaxCategory>" & vbCrLf
        sb = sb & "    </cac:TaxSubtotal>" & vbCrLf
    Next j
    sb = sb & "  </cac:TaxTotal>" & vbCrLf

    BuildTaxTotalBlock = sb
End Function

Private Function BuildLegalMonetaryBlock(wsA As Worksheet, antetRow As Long) As String
    Dim sb As String
    sb = "  <cac:LegalMonetaryTotal>" & vbCrLf
    sb = sb & "    <cbc:LineExtensionAmount currencyID=""RON"">" & FormatAmount(wsA.Cells(antetRow, FA_LINEEXT).Value) & "</cbc:LineExtensionAmount>" & vbCrLf
    sb = sb & "    <cbc:TaxExclusiveAmount currencyID=""RON"">" & FormatAmount(wsA.Cells(antetRow, FA_TAXEXCL).Value) & "</cbc:TaxExclusiveAmount>" & vbCrLf
    sb = sb & "    <cbc:TaxInclusiveAmount currencyID=""RON"">" & FormatAmount(wsA.Cells(antetRow, FA_TAXINCL).Value) & "</cbc:TaxInclusiveAmount>" & vbCrLf
    sb = sb & "    <cbc:PayableAmount currencyID=""RON"">" & FormatAmount(wsA.Cells(antetRow, FA_PAYABLE).Value) & "</cbc:PayableAmount>" & vbCrLf
    sb = sb & "  </cac:LegalMonetaryTotal>" & vbCrLf
    BuildLegalMonetaryBlock = sb
End Function

Private Function BuildAllInvoiceLines(wsL As Worksheet, lastL As Long, invID As String, isVAT As Boolean) As String
    Dim sb As String, j As Long
    For j = 2 To lastL
        If Trim$(GetCellStr(wsL, j, L_INVOICE)) = invID Then
            sb = sb & BuildInvoiceLine(wsL, j, isVAT)
        End If
    Next j
    BuildAllInvoiceLines = sb
End Function

Private Function BuildInvoiceLine(ws As Worksheet, row As Long, isVAT As Boolean) As String
    Dim sb As String
    Dim lineId As String, qty As Double, unit As String, desc As String
    Dim price As Double, ccy As String, lineExt As Double, pct As Double

    lineId = GetCellStr(ws, row, L_LINEID)
    qty = CDblSafe(ws.Cells(row, L_QTY).Value)
    unit = NzStr(GetCellStr(ws, row, L_UNITCODE), "H87")
    desc = GetCellStr(ws, row, L_DESC)
    price = CDblSafe(ws.Cells(row, L_PRICE).Value)
    ccy = NzStr(GetCellStr(ws, row, L_CURRENCY), "RON")
    lineExt = CDblSafe(ws.Cells(row, L_LINEEXT).Value)
    pct = CDblSafe(ws.Cells(row, L_TAXPCT).Value)

    sb = "  <cac:InvoiceLine>" & vbCrLf
    sb = sb & "    <cbc:ID>" & XmlEscape(lineId) & "</cbc:ID>" & vbCrLf
    sb = sb & "    <cbc:InvoicedQuantity unitCode=""" & XmlEscape(unit) & """>" & FormatQty(qty) & "</cbc:InvoicedQuantity>" & vbCrLf
    sb = sb & "    <cbc:LineExtensionAmount currencyID=""" & XmlEscape(ccy) & """>" & FormatAmount(lineExt) & "</cbc:LineExtensionAmount>" & vbCrLf
    sb = sb & "    <cac:Item>" & vbCrLf
    sb = sb & "      <cbc:Name>" & XmlEscape(desc) & "</cbc:Name>" & vbCrLf
    sb = sb & "      <cac:ClassifiedTaxCategory>" & vbCrLf
    If Not isVAT Then
        sb = sb & "        <cbc:ID>O</cbc:ID>" & vbCrLf
    ElseIf pct <= 0.001 Then
        sb = sb & "        <cbc:ID>Z</cbc:ID>" & vbCrLf
        sb = sb & "        <cbc:Percent>0.00</cbc:Percent>" & vbCrLf
    Else
        sb = sb & "        <cbc:ID>S</cbc:ID>" & vbCrLf
        sb = sb & "        <cbc:Percent>" & FormatAmount(pct) & "</cbc:Percent>" & vbCrLf
    End If
    sb = sb & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
    sb = sb & "      </cac:ClassifiedTaxCategory>" & vbCrLf
    sb = sb & "    </cac:Item>" & vbCrLf
    sb = sb & "    <cac:Price><cbc:PriceAmount currencyID=""" & XmlEscape(ccy) & """>" & FormatPrice(Abs(price)) & "</cbc:PriceAmount></cac:Price>" & vbCrLf
    sb = sb & "  </cac:InvoiceLine>" & vbCrLf

    BuildInvoiceLine = sb
End Function

'==============================================================================
' RECONCILIERE - verifica sumele declarate in antet vs. liniile facturii
'==============================================================================
Private Function ReconcileInvoice(wsA As Worksheet, antetRow As Long, _
                                   wsL As Worksheet, lastL As Long, _
                                   invID As String) As String
    Dim sumLines As Double, sumVat As Double, j As Long
    Dim hasLines As Boolean
    For j = 2 To lastL
        If Trim$(GetCellStr(wsL, j, L_INVOICE)) = invID Then
            sumLines = sumLines + CDblSafe(wsL.Cells(j, L_LINEEXT).Value)
            sumVat = sumVat + Round(CDblSafe(wsL.Cells(j, L_LINEEXT).Value) * CDblSafe(wsL.Cells(j, L_TAXPCT).Value) / 100#, 2)
            hasLines = True
        End If
    Next j

    If Not hasLines Then
        ReconcileInvoice = "nicio linie definita in 'Linii_Facturi' pentru aceasta factura."
        Exit Function
    End If

    Dim hLineExt As Double, hVat As Double, hPayable As Double, hTaxIncl As Double
    hLineExt = CDblSafe(wsA.Cells(antetRow, FA_LINEEXT).Value)
    hVat = CDblSafe(wsA.Cells(antetRow, FA_VAT).Value)
    hPayable = CDblSafe(wsA.Cells(antetRow, FA_PAYABLE).Value)
    hTaxIncl = CDblSafe(wsA.Cells(antetRow, FA_TAXINCL).Value)

    If Abs(hLineExt - sumLines) > RECONCILE_TOL Then
        ReconcileInvoice = "LineExtensionAmount antet (" & FormatAmount(hLineExt) & _
                           ") != suma liniilor (" & FormatAmount(sumLines) & ")."
        Exit Function
    End If
    If Abs(hVat - sumVat) > RECONCILE_TOL Then
        ReconcileInvoice = "VATAmount antet (" & FormatAmount(hVat) & _
                           ") != TVA calculat din linii (" & FormatAmount(sumVat) & ")."
        Exit Function
    End If
    If Abs(hPayable - hTaxIncl) > RECONCILE_TOL Then
        ReconcileInvoice = "PayableAmount (" & FormatAmount(hPayable) & _
                           ") != TaxInclusiveAmount (" & FormatAmount(hTaxIncl) & ")."
        Exit Function
    End If
    ReconcileInvoice = ""
End Function

Private Function SumLineExt(wsL As Worksheet, lastL As Long, invID As String) As Double
    Dim s As Double, j As Long
    For j = 2 To lastL
        If Trim$(GetCellStr(wsL, j, L_INVOICE)) = invID Then
            s = s + CDblSafe(wsL.Cells(j, L_LINEEXT).Value)
        End If
    Next j
    SumLineExt = s
End Function

'==============================================================================
' HELPERS
'==============================================================================
Private Function GetSheets(ByRef wsF As Worksheet, ByRef wsC As Worksheet, _
                            ByRef wsA As Worksheet, ByRef wsL As Worksheet) As Boolean
    On Error Resume Next
    Set wsF = ThisWorkbook.Worksheets("Furnizori")
    Set wsC = ThisWorkbook.Worksheets("Clienti")
    Set wsA = ThisWorkbook.Worksheets("Facturi_Antet")
    Set wsL = ThisWorkbook.Worksheets("Linii_Facturi")
    On Error GoTo 0
    If wsF Is Nothing Or wsC Is Nothing Or wsA Is Nothing Or wsL Is Nothing Then
        MsgBox "Lipsesc una sau mai multe foi din workbook (Furnizori / Clienti / Facturi_Antet / Linii_Facturi).", vbExclamation
        GetSheets = False
    Else
        GetSheets = True
    End If
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
End Function

Private Function GetCellStr(ws As Worksheet, r As Long, c As Long) As String
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If IsEmpty(v) Or IsNull(v) Then
        GetCellStr = ""
    Else
        GetCellStr = CStr(v)
    End If
End Function

Private Function FindRow(ws As Worksheet, lastRow As Long, col As Long, key As String) As Long
    Dim i As Long, k As String
    k = Trim$(key)
    For i = 2 To lastRow
        If Trim$(GetCellStr(ws, i, col)) = k Then
            FindRow = i
            Exit Function
        End If
    Next i
    FindRow = 0
End Function

Private Function NzStr(s As String, fallback As String) As String
    If Trim$(s) = "" Then NzStr = fallback Else NzStr = s
End Function

Private Function CDblSafe(v As Variant) As Double
    On Error Resume Next
    If IsEmpty(v) Or IsNull(v) Or Trim$(CStr(v)) = "" Then
        CDblSafe = 0
    Else
        CDblSafe = CDbl(v)
    End If
    On Error GoTo 0
End Function

Private Function XmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    XmlEscape = s
End Function

Private Function HasROPrefix(ByVal s As String) As Boolean
    HasROPrefix = (UCase$(Left$(Trim$(s), 2)) = "RO")
End Function

Private Function StripROPrefix(ByVal s As String) As String
    s = Trim$(s)
    If HasROPrefix(s) Then StripROPrefix = Mid$(s, 3) Else StripROPrefix = s
End Function

Private Function IsCNP(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) <> 13 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    IsCNP = (InStr("123456789", Left$(s, 1)) > 0)
End Function

Private Function IsVATPayerFlag(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    IsVATPayerFlag = (s = "DA" Or s = "YES" Or s = "Y" Or s = "TRUE" Or s = "1")
End Function

' Standard rounding (half away from zero), evita banker's rounding din VBA Round()
Private Function RoundHalfAway(ByVal v As Double, ByVal d As Integer) As Double
    Dim f As Double
    f = 10 ^ d
    If v >= 0 Then
        RoundHalfAway = Int(v * f + 0.5) / f
    Else
        RoundHalfAway = -Int(-v * f + 0.5) / f
    End If
End Function

Private Function FormatAmount(ByVal v As Variant) As String
    Dim n As Double
    n = RoundHalfAway(CDblSafe(v), 2)
    FormatAmount = Replace(Format$(n, "0.00"), ",", ".")
End Function

Private Function FormatPrice(ByVal v As Variant) As String
    Dim n As Double
    n = RoundHalfAway(CDblSafe(v), 4)
    FormatPrice = Replace(Format$(n, "0.0000"), ",", ".")
End Function

Private Function FormatQty(ByVal v As Variant) As String
    Dim n As Double
    n = RoundHalfAway(CDblSafe(v), 3)
    FormatQty = Replace(Format$(n, "0.000"), ",", ".")
End Function

Private Function FormatIso(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Or Trim$(CStr(v)) = "" Then
        FormatIso = ""
    Else
        FormatIso = Format$(CDate(v), "yyyy-mm-dd")
    End If
End Function

Private Function SanitizeForFilename(ByVal s As String) As String
    s = Replace(s, "/", "_")
    s = Replace(s, "\", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    s = Replace(s, " ", "_")
    SanitizeForFilename = s
End Function

' Scrie UTF-8 fara BOM (ANAF refuza BOM)
Private Sub WriteUtf8(ByVal path As String, ByVal content As String)
    Dim sText As Object, sBin As Object
    Set sText = CreateObject("ADODB.Stream")
    sText.Type = 2 ' adTypeText
    sText.Charset = "utf-8"
    sText.Open
    sText.WriteText content

    Set sBin = CreateObject("ADODB.Stream")
    sBin.Type = 1 ' adTypeBinary
    sBin.Open
    sText.Position = 3 ' sare peste BOM
    sText.CopyTo sBin
    sText.Close

    sBin.SaveToFile path, 2 ' adSaveCreateOverWrite
    sBin.Close
End Sub
