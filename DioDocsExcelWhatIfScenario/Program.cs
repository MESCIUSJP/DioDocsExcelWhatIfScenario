// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;

Console.WriteLine("ワークシートにWhat-If分析（シナリオ）を追加する");

Workbook workbook = new();
workbook.Open("what-if-scenario.xlsx");
IWorksheet worksheet = workbook.ActiveSheet;

// 値引き率が低いシナリオ
List<object> lessDiscountRatesValues = [0.08, 0.05, 0.05, 0.03, 0.03];
IScenario lessDiscountRates = worksheet.Scenarios.Add(
    "低値引き率",
    worksheet.Range["D2:D6"],
    lessDiscountRatesValues);

// 標準の値引き率のシナリオ
List<object> normalDiscountRatesValues = [0.1, 0.07, 0.07, 0.05, 0.05];
IScenario normalDiscountRates = worksheet.Scenarios.Add(
    "標準値引き率",
    worksheet.Range["D2:D6"],
    normalDiscountRatesValues);

// 割引なしで販売するシナリオ
List<object> sellingWithoutDiscountValues = [0, 0, 0, 0, 0];
IScenario sellingWithoutDiscount = worksheet.Scenarios.Add(
    "割引なし",
    worksheet.Range["D2:D6"],
    sellingWithoutDiscountValues);

// 割引なしで販売するシナリオを表示
worksheet.Scenarios["割引なし"].Show();

// Excelファイルに保存
workbook.Save("result-what-if-scenario.xlsx");