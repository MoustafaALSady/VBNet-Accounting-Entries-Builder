
Public Class AccountingDetailBuilder

    Public Structure DetailEntry
        Public AccountName As String
        Public AccountNo As Integer
        Public Debit As Decimal
        Public Credit As Decimal
        Public Details As String
        Public AccountCod As Integer
        Public Symbol As String
        Public TextId As String
        Public IsSpecial As Boolean
        Public RegNumber As String
    End Structure

    Private ReadOnly entries As New List(Of DetailEntry)()
    Private ReadOnly symbol As String
    Private ReadOnly textId As String
    Private ReadOnly regNumber As String

    Public Sub New(_symbol As String, _textId As String, _regNumber As String)
        symbol = _symbol
        textId = _textId
        regNumber = _regNumber
    End Sub

    Public Sub Debit(accountName As String, accountNo As Integer, amount As Decimal, details As String, accountCod As Integer, Optional isSpecial As Boolean = False)
        entries.Add(New DetailEntry With {
            .AccountName = accountName,
            .AccountNo = accountNo,
            .Debit = amount,
            .Credit = 0D,
            .Details = details,
            .AccountCod = accountCod,
            .Symbol = symbol,
            .TextId = textId,
            .IsSpecial = isSpecial,
            .RegNumber = regNumber
        })
    End Sub

    Public Sub Credit(accountName As String, accountNo As Integer, amount As Decimal, details As String, accountCod As Integer, Optional isSpecial As Boolean = False)
        entries.Add(New DetailEntry With {
            .AccountName = accountName,
            .AccountNo = accountNo,
            .Debit = 0D,
            .Credit = amount,
            .Details = details,
            .AccountCod = accountCod,
            .Symbol = symbol,
            .TextId = textId,
            .IsSpecial = isSpecial,
            .RegNumber = regNumber
        })
    End Sub

    Public Sub Build()
        Dim seq As Integer = 1
        For Each e In entries
            DetailsAccountingEntries(seq, e.AccountName, e.AccountNo, e.Debit, e.Credit, e.Details, e.AccountCod, e.Symbol, e.TextId, e.IsSpecial, e.RegNumber)
            seq += 1
        Next
    End Sub

End Class