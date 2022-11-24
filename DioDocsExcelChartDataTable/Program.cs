// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing;

Console.WriteLine("DioDocs for Excelでチャートにデータテーブルを追加");

var workbook = new Workbook();
var worksheet1 = workbook.Worksheets[0];

// データ
var data = new object[,]
{
    {"国・地域", "第1四半期", "第2四半期", "第3四半期", "第4四半期" },
    {"オーストラリア", 16439, 18106, 15193, 14879},
    {"中国", 42659, 14392, 42284, 38270},
    {"日本", 44000, 15039, 27961, 34382},
    {"アメリカ", 23174, 42797, 23637, 26200}
};

worksheet1.Name = "四半期売上レポート";
worksheet1.Range["A1:E5"].Value = data;
worksheet1.Range["A1:E5"].AutoFit();
worksheet1.Range["B2:E5"].NumberFormat = @"¥#,##0";

// チャートを作成
var rect = CellInfo.GetAccurateRangeBoundary(worksheet1.Range["B7:L26"]);
IShape chartCol = worksheet1.Shapes.AddChartInPixel(ChartType.ColumnClustered, rect.Left, rect.Top, rect.Width, rect.Height);
chartCol.Chart.SeriesCollection.Add(worksheet1.Range["A1:E5"]);
chartCol.Chart.ChartTitle.Text = "四半期売上";

// データテーブルを追加
chartCol.Chart.HasDataTable = true;

// Excelファイルに保存
workbook.Save("ChartDataTable.xlsx");

// データテーブルをカスタマイズ
IDataTable dataTable = chartCol.Chart.DataTable;
dataTable.ShowLegendKey = true;
dataTable.HasBorderHorizontal = false;
dataTable.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent6;
dataTable.Font.Color.ObjectThemeColor = ThemeColor.Accent2;
dataTable.Font.Size = 10;
dataTable.Font.Italic = true;

// Excelファイルに保存
workbook.Save("ChartDataTable_Custom.xlsx");
