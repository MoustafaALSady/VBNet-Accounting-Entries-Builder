
'شكرا كثيرا لكم .
'أرغب مشاركتكم النتائج .
'هذه الوحدة كاملة 

Imports CC_JO.PaymentMethodHelper

Public Class SaveXTransfer

    Public Shared Sub SaveMovesDataOpeningBalanceTransfor(frm As Form)
        Dim TextID = ControlFinder.GetTextEdit(frm, "TextID")
        Dim DateMovementHistory = ControlFinder.GetDatePicker(frm, "DateMovementHistory")
        Dim TextMovementSymbol = ControlFinder.GetTextEdit(frm, "TextMovementSymbol")

        ModuleGetBalanceAndMaxRecord.GetMaxIDDailyRestrictions()

        AccountingEntries(MovementNumberValue, RegistrationNumberValue,
                      DateMovementHistory.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture),
                      MovementDetails, False, ParseDecimal(BalanceDebit), ParseDecimal(BalanceDebit),
                      CStr(JournalNoValue), "قيد", Symbol, CStr(TextMovementSymbol.EditValue), False)

        Dim b As New AccountingDetailBuilder(CStr(TextMovementSymbol.EditValue), CStr(TextID.EditValue), CStr(RegistrationNumberValue))

        b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod)
        b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(BalanceDebit), MovementDetailsA, FundAccount_Cod)

        b.Build()
    End Sub

    Public Shared Sub SaveMovesDataTransfor(frm As Form)
        Dim TextID = ControlFinder.GetTextEdit(frm, "TextID")
        Dim DateMovementHistory = ControlFinder.GetDatePicker(frm, "DateMovementHistory")
        Dim TextMovementSymbol = ControlFinder.GetTextEdit(frm, "TextMovementSymbol")
        Dim comboPaymentMethod As ComboBox = ControlFinder.GetComboBox(frm, "ComboPaymentMethod")
        Dim selectedMethod As PaymentMethod = CType(PaymentMethodHelper.GetPaymentMethod(comboPaymentMethod.Text), PaymentMethod)

        ModuleGetBalanceAndMaxRecord.GetMaxIDDailyRestrictions()

        AccountingEntries(MovementNumberValue, RegistrationNumberValue,
                      DateMovementHistory.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture),
                      MovementDetails, False, ParseDecimal(BalanceDebit), ParseDecimal(BalanceDebit),
                      CStr(JournalNoValue), "قيد", Symbol, CStr(TextMovementSymbol.EditValue), False)

        If comboPaymentMethod Is Nothing Then Return

        Dim b As New AccountingDetailBuilder(CStr(TextMovementSymbol.EditValue), CStr(TextID.EditValue), CStr(RegistrationNumberValue))

        b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod)

        Select Case selectedMethod
            Case PaymentMethod.Cash
                b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(BalanceDebit), MovementDetailsA, FundAccount_Cod)

            Case PaymentMethod.Check
                b.Credit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(ValueOfCheckCredit), MovementDetailsB, ChecksAccount_Cod)

            Case PaymentMethod.CashAndCheck
                b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(FundValueCredit), MovementDetailsA, FundAccount_Cod)
                b.Credit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(ValueOfCheckCredit), MovementDetailsB, ChecksAccount_Cod)
        End Select

        b.Build()
    End Sub

    ' === دوال مساعدة خاصة بالمبيعات ===
    Private Shared Sub AddSalesDiscountIfAny(b As AccountingDetailBuilder)
        If ParseDecimal(DiscountVal) > 0 Then
            b.Debit(DiscountAccountAE_Name, PurchSalesDiscount_No, ParseDecimal(DiscountVal), MovementDetailsC, DiscountAccount_Cod)
        End If
    End Sub

    Private Shared Sub AddSalesTaxIfAny(b As AccountingDetailBuilder)
        If ParseDecimal(SalesTaxRateVal) > 0 Then
            b.Credit(CalculatingTaxAccount_Name, PurchSalesCalculatingTax_No, ParseDecimal(SalesTaxRateVal), MovementDetailsD, TaxAccount_Cod)
        End If
    End Sub

    Private Shared Sub AddSalesCashPayment(b As AccountingDetailBuilder, selectedMethod As PaymentMethod)
        Select Case selectedMethod
            Case PaymentMethod.Cash
                b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod)

            Case PaymentMethod.Check
                b.Debit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(BalanceDebit), MovementDetails, ChecksAccount_Cod)

            Case PaymentMethod.CashAndCheck
                b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(FundValueDebit), MovementDetails, DebitAccount_Cod)
                b.Debit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(ValueOfCheckDebit), MovementDetailsA, ChecksAccount_Cod)
        End Select

        AddSalesDiscountIfAny(b)
        b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(BalanceMediator), MovementDetailsA, CredAccount_Cod)
        AddSalesTaxIfAny(b)
    End Sub

    Private Shared Sub AddSalesCreditPayment(b As AccountingDetailBuilder)
        b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod)
        AddSalesDiscountIfAny(b)
        b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(BalanceDebit), MovementDetailsA, CredAccount_Cod)
        AddSalesTaxIfAny(b)
    End Sub

    ' === الدالة الرئيسية بعد الرفاكتور النهائي ===
    Public Shared Sub SaveMovesDataSalesTransfor(frm As Form)
        Dim TextID = ControlFinder.GetTextEdit(frm, "TextID")
        Dim DateMovementHistory = ControlFinder.GetDatePicker(frm, "DateMovementHistory")
        Dim TextMovementSymbol = ControlFinder.GetTextEdit(frm, "TextMovementSymbol")
        Dim comboPaymentMethod As ComboBox = ControlFinder.GetComboBox(frm, "ComboPaymentMethod")
        Dim selectedMethod As PaymentMethod = CType(PaymentMethodHelper.GetPaymentMethod(comboPaymentMethod.Text), PaymentMethod)

        ModuleGetBalanceAndMaxRecord.GetMaxIDDailyRestrictions()

        AccountingEntries(MovementNumberValue, RegistrationNumberValue,
                      DateMovementHistory.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture),
                      MovementDetails, False, ParseDecimal(BalanceDebit), ParseDecimal(BalanceDebit),
                      CStr(JournalNoValue), "قيد", Symbol, CStr(TextMovementSymbol.EditValue), False)

        Dim b As New AccountingDetailBuilder(CStr(TextMovementSymbol.EditValue), CStr(TextID.EditValue), CStr(RegistrationNumberValue))

        If IsCash Then
            AddSalesCashPayment(b, selectedMethod)
        Else
            AddSalesCreditPayment(b)
        End If

        b.Build()
    End Sub

    ' === دوال مساعدة خاصة بالمشتريات/القروض ===
    Private Shared Sub AddPurchaseTaxIfAny(b As AccountingDetailBuilder)
        If ParseDecimal(SalesTaxRateVal) > 0 Then
            b.Debit(CalculatingTaxAccount_Name, PurchSalesCalculatingTax_No, ParseDecimal(SalesTaxRateVal), MovementDetailsE, TaxAccount_Cod)
        End If
    End Sub

    Private Shared Sub AddPurchaseDiscountIfAny(b As AccountingDetailBuilder)
        If ParseDecimal(DiscountVal) > 0 Then
            b.Credit(DiscountAccountAE_Name, PurchSalesDiscount_No, ParseDecimal(DiscountVal), MovementDetailsC, DiscountAccount_Cod)
        End If
    End Sub

    Private Shared Sub AddPurchasePayments(b As AccountingDetailBuilder, selectedMethod As PaymentMethod)
        Select Case selectedMethod
            Case PaymentMethod.Cash
                b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(BalanceMediator), MovementDetailsA, FundAccount_Cod)

            Case PaymentMethod.Check
                b.Credit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(ValueOfCheckCredit), MovementDetailsB, ChecksAccount_Cod)

            Case PaymentMethod.CashAndCheck
                b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(FundValueCredit), MovementDetailsA, FundAccount_Cod)
                b.Credit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(ValueOfCheckCredit), MovementDetailsB, ChecksAccount_Cod)

            Case PaymentMethod.AccountsPayable
                b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(ValueAccountsPayableCredit), MovementDetailsD, CredAccount_Cod)

            Case PaymentMethod.CashAndAccountsPayable
                b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(FundValueCredit), MovementDetailsA, FundAccount_Cod)
                b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(ValueAccountsPayableCredit), MovementDetailsD, CredAccount_Cod)

            Case PaymentMethod.CheckAndAccountsPayable
                b.Credit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(ValueOfCheckCredit), MovementDetailsB, ChecksAccount_Cod)
                b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(ValueAccountsPayableCredit), MovementDetailsD, CredAccount_Cod)

            Case PaymentMethod.CashAndCheckAndAccountsPayable
                b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(FundValueCredit), MovementDetailsA, FundAccount_Cod)
                b.Credit(ChecksAccount_Name, ChecksAccount_NO, ParseDecimal(ValueOfCheckCredit), MovementDetailsB, ChecksAccount_Cod)
                b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(ValueAccountsPayableCredit), MovementDetailsD, CredAccount_Cod)
        End Select
    End Sub

    ' === الدالة الرئيسية بعد الرفاكتور النهائي ===
    Public Shared Sub SaveMovesDataLoansPurchasesTransfor(frm As Form)
        Dim TextID = ControlFinder.GetTextEdit(frm, "TextID")
        Dim DateMovementHistory = ControlFinder.GetDatePicker(frm, "DateMovementHistory")
        Dim TextMovementSymbol = ControlFinder.GetTextEdit(frm, "TextMovementSymbol")
        Dim comboPaymentMethod As ComboBox = ControlFinder.GetComboBox(frm, "ComboPaymentMethod")
        Dim selectedMethod As PaymentMethod = CType(PaymentMethodHelper.GetPaymentMethod(comboPaymentMethod.Text), PaymentMethod)

        ModuleGetBalanceAndMaxRecord.GetMaxIDDailyRestrictions()

        AccountingEntries(MovementNumberValue, RegistrationNumberValue,
                      DateMovementHistory.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture),
                      MovementDetails, True, ParseDecimal(BalanceDebit), ParseDecimal(BalanceDebit),
                      CStr(JournalNoValue), "قيد", Symbol, CStr(TextMovementSymbol.EditValue), False)

        Dim b As New AccountingDetailBuilder(CStr(TextMovementSymbol.EditValue), CStr(TextID.EditValue), CStr(RegistrationNumberValue))

        b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod)
        AddPurchaseTaxIfAny(b)
        AddPurchasePayments(b, selectedMethod)
        AddPurchaseDiscountIfAny(b)

        b.Build()
    End Sub

    Public Shared Sub SaveMovesDataLoansTransfor(frm As Form)
        Dim TextID = ControlFinder.GetTextEdit(frm, "TextID")
        Dim DateMovementHistory = ControlFinder.GetDatePicker(frm, "DateMovementHistory")
        Dim TextMovementSymbol = ControlFinder.GetTextEdit(frm, "TextMovementSymbol")

        ModuleGetBalanceAndMaxRecord.GetMaxIDDailyRestrictions()

        AccountingEntries(MovementNumberValue, RegistrationNumberValue,
                      DateMovementHistory.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture),
                      MovementDetails, False, ParseDecimal(BalanceDebit), ParseDecimal(BalanceDebit),
                      CStr(JournalNoValue), "قيد", Symbol, CStr(TextMovementSymbol.EditValue), False)

        Dim b As New AccountingDetailBuilder(CStr(TextMovementSymbol.EditValue), CStr(TextID.EditValue), CStr(RegistrationNumberValue))

        ' المدين الأساسي دائماً موجود
        b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod)

        If Not IsFirstBatch Then
            b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(BalanceDebit), MovementDetailsA, FundAccount_Cod)
        Else
            ' الدفعة الأولى
            b.Credit(CStr(ModuleGeneral.FundName.Value), FundAccount_No, ParseDecimal(FundValueCredit), MovementDetailsA, FundAccount_Cod)
            b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(BalanceMediator), MovementDetails, CredAccount_Cod)
        End If

        b.Build()
    End Sub

    Public Shared Sub SaveMovesCustomerTransfor(frm As Form)
        Dim TextID = ControlFinder.GetTextEdit(frm, "TextID")
        Dim DateMovementHistory = ControlFinder.GetDatePicker(frm, "DateMovementHistory")
        Dim TextMovementSymbol = ControlFinder.GetTextEdit(frm, "TextMovementSymbol")

        ModuleGetBalanceAndMaxRecord.GetMaxIDDailyRestrictions()

        AccountingEntries(MovementNumberValue, RegistrationNumberValue,
                      DateMovementHistory.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture),
                      MovementDetails, True, ParseDecimal(BalanceDebit), ParseDecimal(BalanceDebit),
                      CStr(JournalNoValue), "قيد", Symbol, CStr(TextMovementSymbol.EditValue), False)

        Dim b As New AccountingDetailBuilder(CStr(TextMovementSymbol.EditValue), CStr(TextID.EditValue), CStr(RegistrationNumberValue))

        b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod, isSpecial:=True)
        b.Credit(CredAccount_Name, CredAccount_NO, ParseDecimal(BalanceDebit), MovementDetails, CredAccount_Cod, isSpecial:=True)

        b.Build()
    End Sub

    Private Shared Sub AddSocialSecurityIfAny(b As AccountingDetailBuilder)
        Dim contributionsNo = ParseInt(keyAccounts.GetValue("SocialSecurityContributions_No", MigrateRestrictions.SocialSecurityContributions_No))

        GetNoRecord("ACCOUNTSTREE", "Account_Name", "Account_No", MigrateRestrictions.SocialSecurityContributions_No, 1)
        Dim contributionsName = ID_Nam

        GetNoRecord("ACCOUNTSTREE", "ACC", "Account_No", MigrateRestrictions.SocialSecurityContributions_No, 1)
        Dim contributionsCod = ParseInt(ID_Nam)

        If ParseDecimal(SocialSecuritySubscription) > 0 Then
            b.Debit(contributionsName, contributionsNo, ParseDecimal(SocialSecuritySubscription), MovementDetailsD, contributionsCod)
        End If
    End Sub

    Public Shared Sub SaveMovesEmployeesSalariesTransfor(frm As Form)
        Dim TextID = ControlFinder.GetTextEdit(frm, "TextID")
        Dim DateMovementHistory = ControlFinder.GetDatePicker(frm, "DateMovementHistory")
        Dim TextMovementSymbol = ControlFinder.GetTextEdit(frm, "TextMovementSymbol")

        ModuleGetBalanceAndMaxRecord.GetMaxIDDailyRestrictions()

        AccountingEntries(MovementNumberValue, RegistrationNumberValue,
                      DateMovementHistory.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture),
                      MovementDetails, False, ParseDecimal(BalanceDebit), ParseDecimal(BalanceDebit),
                      CStr(JournalNoValue), "قيد", Symbol, CStr(TextMovementSymbol.EditValue), False)

        Dim b As New AccountingDetailBuilder(CStr(TextMovementSymbol.EditValue), CStr(TextID.EditValue), CStr(RegistrationNumberValue))

        ' رواتب الموظفين
        b.Debit(DebitAccount_Name, DebitAccount_No, ParseDecimal(BalanceDebit), MovementDetails, DebitAccount_Cod)

        ' اشتراكات التأمينات الاجتماعية
        AddSocialSecurityIfAny(b)
        b.Build()
    End Sub


End Class
