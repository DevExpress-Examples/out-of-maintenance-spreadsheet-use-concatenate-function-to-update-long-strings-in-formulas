Imports System
Imports System.Collections.Generic
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas

Namespace SpreadsheetApp
    Friend Class ReplaceLongStringsWalker
        Inherits ExpressionVisitor

        Private ReadOnly workbook As IWorkbook
        Private ReadOnly modifiers As New Stack(Of Action(Of IExpression))()

        Public Sub New(ByVal workbook As IWorkbook)
            Me.workbook = workbook
        End Sub

        Private privateIsExpressionModified As Boolean
        Public Property IsExpressionModified() As Boolean
            Get
                Return privateIsExpressionModified
            End Get
            Private Set(ByVal value As Boolean)
                privateIsExpressionModified = value
            End Set
        End Property

        Public Sub Walk(ByVal parsedExpression As IParsedExpression)
            modifiers.Push(Sub(x) parsedExpression.Expression = x)
            Try
                parsedExpression.Expression.Visit(Me)
            Finally
                modifiers.Pop()
            End Try
        End Sub

        Public Overrides Sub Visit(ByVal expression As ConstantExpression)
            If expression.Value.IsText AndAlso expression.Value.TextValue.Length > 255 Then
                Dim modifer = modifiers.Peek()
                modifer(New FunctionExpression(workbook.Functions("CONCATENATE"), SplitLongString(expression.Value.TextValue)))
                IsExpressionModified = True
            End If
        End Sub

        Public Overrides Sub VisitBinary(ByVal expression As BinaryOperatorExpression)
            modifiers.Push(Sub(x) expression.LeftExpression = x)
            Try
                expression.LeftExpression.Visit(Me)
            Finally
                modifiers.Pop()
            End Try
            modifiers.Push(Sub(x) expression.RightExpression = x)
            Try
                expression.RightExpression.Visit(Me)
            Finally
                modifiers.Pop()
            End Try
        End Sub

        Public Overrides Sub VisitFunction(ByVal expression As FunctionExpressionBase)
            For i As Integer = 0 To expression.InnerExpressions.Count - 1
                Dim j = i
                modifiers.Push(Sub(x) expression.InnerExpressions(j) = x)
                Try
                    Dim innerExpression As IExpression = expression.InnerExpressions(i)
                    innerExpression.Visit(Me)
                Finally
                    modifiers.Pop()
                End Try
            Next i
        End Sub

        Private Function SplitLongString(ByVal value As String) As IList(Of IExpression)
            Dim result As New List(Of IExpression)()
            Do While value.Length > 255
                result.Add(New ConstantExpression(value.Substring(0, 255)))
                value = value.Remove(0, 255)
            Loop
            If value.Length > 0 Then
                result.Add(New ConstantExpression(value))
            End If
            Return result
        End Function
    End Class
End Namespace
