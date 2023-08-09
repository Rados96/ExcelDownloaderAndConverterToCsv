using OfficeOpenXml;

class Program
{
    static readonly HttpClient httpClient = new HttpClient();

    static async Task Main(string[] args)
    {
        await ProcessRigCountData();
    }

    static async Task ProcessRigCountData()
    {

        string url = "https://bakerhughesrigcount.gcs-web.com/static-files/7240366e-61cc-4acb-89bf-86dc1a0dffe8";
        string excelFileName = "Worldwide Rig Count Jun 2023.xlsx";
        string csvFileName = "Worldwide Rig Count Jun 2023.csv";
        int currentYear = DateTime.Now.Year;
        int twoYearsAgo = currentYear - 2;

        httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko; Google Page Speed Insights) Chrome/27.0.1453 Safari/537.36");

        try
        {
            await DownloadFile(url, excelFileName);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(excelFileName)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                if (sheet == null) return;

                int startRow = FindStartRow(sheet, currentYear);
                int endRow = FindEndRow(sheet, twoYearsAgo, startRow);
                if (startRow > 0 &&  endRow > 0)
                {
                    await SaveDataToCSV(sheet, startRow, endRow, csvFileName);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static async Task DownloadFile(string url, string fileName)
    {
        using (HttpResponseMessage response = await httpClient.GetAsync(url))
        {
            response.EnsureSuccessStatusCode();
            using (Stream contentStream = await response.Content.ReadAsStreamAsync())
            using (FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 4096, useAsync: true))
            {
                await contentStream.CopyToAsync(fileStream);
            }
        }
    }

    static int FindStartRow(ExcelWorksheet sheet, int currentYear)
    {
        for (int row = 1; row <= sheet.Dimension.Rows; row++)
        {
            string value = sheet.Cells[row, 2].Text;
            if (value == currentYear.ToString())
            {
                return row;
            }
        }
        return -1;
    }

    static int FindEndRow(ExcelWorksheet sheet, int twoYearsAgo, int startRow)
    {
        for (int row = startRow; row <= sheet.Dimension.Rows; row++)
        {
            string value = sheet.Cells[row, 2].Text;
            if (value == twoYearsAgo.ToString())
            {
                return row - 2;
            }
        }
        return -1;
    }

    static async Task SaveDataToCSV(ExcelWorksheet worksheet, int startRow, int endRow, string csvFileName)
    {
        using (StreamWriter writer = new StreamWriter(csvFileName, false))
        {
            for (int row = startRow; row <= endRow; row++)
            {
                List<string> rowData = new List<string>();
                for (int col = 1; col < worksheet.Dimension.Columns; col++)
                {
                    string cellValue = worksheet.Cells[row, col].Text;
                    if (cellValue == "")
                    {
                        if (col == 2) continue;
                        rowData.Add(",");
                    }
                    else
                    {
                        rowData.Add(cellValue);
                        if (col < worksheet.Dimension.Columns - 1 && worksheet.Cells[row, col + 1].Text != "")
                        {
                            rowData.Add(",");
                        }
                    }
                }
                string rowCsv = string.Join("", rowData);
                await writer.WriteLineAsync(rowCsv);
            }
        }
    }
}
