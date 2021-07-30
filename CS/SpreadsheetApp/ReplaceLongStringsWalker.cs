using System;
using System.Collections.Generic;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;

namespace SpreadsheetApp
{
    internal class ReplaceLongStringsWalker : ExpressionVisitor
    {
        readonly IWorkbook workbook;
        readonly Stack<Action<IExpression>> modifiers = new Stack<Action<IExpression>>();

        public ReplaceLongStringsWalker(IWorkbook workbook)
        {
            this.workbook = workbook;
        }

        public bool IsExpressionModified { get; private set; }

        public void Walk(IParsedExpression parsedExpression)
        {
            modifiers.Push(x => parsedExpression.Expression = x);
            try
            {
                parsedExpression.Expression.Visit(this);
            }
            finally
            {
                modifiers.Pop();
            }
        }

        public override void Visit(ConstantExpression expression)
        {
            if (expression.Value.IsText && expression.Value.TextValue.Length > 255)
            {
                var modifer = modifiers.Peek();
                modifer(new FunctionExpression(workbook.Functions["CONCATENATE"],
                        SplitLongString(expression.Value.TextValue)));
                IsExpressionModified = true;
            }
        }

        public override void VisitBinary(BinaryOperatorExpression expression)
        {
            modifiers.Push(x => expression.LeftExpression = x);
            try
            {
                expression.LeftExpression.Visit(this);
            }
            finally
            {
                modifiers.Pop();
            }
            modifiers.Push(x => expression.RightExpression = x);
            try
            {
                expression.RightExpression.Visit(this);
            }
            finally
            {
                modifiers.Pop();
            }
        }

        public override void VisitFunction(FunctionExpressionBase expression)
        {
            for (int i = 0; i < expression.InnerExpressions.Count; i++)
            {
                modifiers.Push(x => expression.InnerExpressions[i] = x);
                try
                {
                    IExpression innerExpression = expression.InnerExpressions[i];
                    innerExpression.Visit(this);
                }
                finally
                {
                    modifiers.Pop();
                }
            }
        }

        IList<IExpression> SplitLongString(string value)
        {
            List<IExpression> result = new List<IExpression>();
            while (value.Length > 255)
            {
                result.Add(new ConstantExpression(value.Substring(0, 255)));
                value = value.Remove(0, 255);
            }
            if (value.Length > 0)
                result.Add(new ConstantExpression(value));
            return result;
        }
    }
}