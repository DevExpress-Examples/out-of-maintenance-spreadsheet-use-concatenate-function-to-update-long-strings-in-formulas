using System;
using System.Data;
using System.Linq;

namespace SpreadsheetApp
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
            spreadsheetControl1.Options.Compatibility.TruncateLongStringsInFormulas = false;
            spreadsheetControl1.LoadDocument("FileWithLongStringInFormula.xlsx");
        }

        private void spreadsheetControl1_DocumentLoaded(object sender, EventArgs e)
        {
            var workbook = spreadsheetControl1.Document;
            workbook.BeginUpdate();
            foreach (var sheet in workbook.Worksheets)
            {
                foreach (var cell in sheet.GetExistingCells().Where(x => x.HasFormula))
                {
                    var walker = new ReplaceLongStringsWalker(workbook);
                    var parsedExpression = cell.ParsedExpression;
                    walker.Walk(parsedExpression);
                    if (walker.IsExpressionModified)
                        cell.ParsedExpression = parsedExpression;
                }
            }
            workbook.EndUpdate();
        }
    }
}
