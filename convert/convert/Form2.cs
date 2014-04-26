using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace convert
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 新しいxlsxドキュメントを作成
            SpreadsheetDocument document = SpreadsheetDocument.Create("Test.xlsx", SpreadsheetDocumentType.Workbook, true);

            // ドキュメントのワークブックパートに、ワークブックを設定
            WorkbookPart wbpart = document.AddWorkbookPart();        wbpart.Workbook = new Workbook();

            // ワークブックパートに、ワークシートパートを設定
            WorksheetPart wspart = wbpart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();
            wspart.Worksheet = new Worksheet(sheetData);

            // ワークブックにシートを設定
            Sheets sheets = wbpart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // シートを1つ追加
            Sheet sheet = new Sheet() { Id = wbpart.GetIdOfPart(wspart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);


            // ここまでがお決まりの処理 
            // Sheetではなく、SheetDataにデータを設定していく

            // Cell単独では存在できない模様
            // Rowオブジェクトを作成し、そこにCellデータを追加していく。
            Row row = new Row();

            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellReference = "A1";
            cell.CellValue = new CellValue("A1のセル");
            row.Append(cell); 

            cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellReference = "B1";
            cell.CellValue = new CellValue("B1のセル");
            row.Append(cell);

            // 行が変わるタイミングで、Rowオブジェクトを再設定
            sheetData.Append(row);
            row = new Row();
            cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellReference = "A2";
            cell.CellValue = new CellValue("A2のセル");
            row.Append(cell);

            /* こういう書き方でもOK
             * row.Append(new Cell() {
             *   DataType = CellValues.String,
             *   CellReference = "A1",
             *   CellValue = new CellValue("Hello world!"), 
             * });
             */

            // 最後にRowをSheetDataに追加
            sheetData.Append(row);

            // ファイルを保存
            document.Close();
        }
    }
}
