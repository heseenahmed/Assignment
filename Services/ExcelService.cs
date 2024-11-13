using Assignment.Models;
using OfficeOpenXml;
using System.Globalization;

public class ExcelService
{
    public List<TaxRecord> ImportExcelFile(Stream fileStream)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage(fileStream);
        var worksheet = package.Workbook.Worksheets[0];

        List<TaxRecord> records = new List<TaxRecord>();
        int rowCount = worksheet.Dimension.Rows;

        for (int row = 2; row <= rowCount; row++)
        {
            try
            {
                var record = new TaxRecord
                {
                    InvNo = int.TryParse(worksheet.Cells[row, 1].Text, out int invNo) ? invNo : 0,
                    InvCURNo = worksheet.Cells[row, 2].Text,
                    InvDate = DateTime.TryParse(worksheet.Cells[row, 3].Text, out DateTime invDate) ? invDate : DateTime.MinValue,
                    CustomerCode = int.TryParse(worksheet.Cells[row, 4].Text, out int customerCode) ? customerCode : 0,
                    CustomerName = worksheet.Cells[row, 5].Text,
                    RegCountry = worksheet.Cells[row, 6].Text,
                    TotalValueAfterTaxing = TryParseDecimal(worksheet.Cells[row, 7].Text),
                    TaxingValue = TryParseDecimal(worksheet.Cells[row, 8].Text),
                };

                record.TotalValueBeforeTaxing = record.TotalValueAfterTaxing / (1 + (record.TaxingValue / 100));

                records.Add(record);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing row {row}: {ex.Message}");
            }
        }

        return records;
    }

    private decimal TryParseDecimal(string input)
    {
        if (decimal.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
        {
            return result;
        }
        return 0m;
    }

    public byte[] AddColumnAndCalculateTotals(List<TaxRecord> records)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("ModifiedSheet");

        worksheet.Cells[1, 1].Value = "Inv_No";
        worksheet.Cells[1, 2].Value = "Inv_CURNo";
        worksheet.Cells[1, 3].Value = "Inv_Date";
        worksheet.Cells[1, 4].Value = "Customer Code";
        worksheet.Cells[1, 5].Value = "Customer Name";
        worksheet.Cells[1, 6].Value = "REG_COUNTRY_APREV";
        worksheet.Cells[1, 7].Value = "Total Value After Taxing";
        worksheet.Cells[1, 8].Value = "Taxing Value";
        worksheet.Cells[1, 9].Value = "Total Value Before Taxing";

        for (int i = 0; i < records.Count; i++)
        {
            var record = records[i];

            record.TotalValueBeforeTaxing = record.TotalValueAfterTaxing / (1 + (record.TaxingValue / 100));

            worksheet.Cells[i + 2, 1].Value = record.InvNo;
            worksheet.Cells[i + 2, 2].Value = record.InvCURNo;
            worksheet.Cells[i + 2, 3].Value = record.InvDate.ToString("dd/MM/yyyy"); 
            worksheet.Cells[i + 2, 4].Value = record.CustomerCode;
            worksheet.Cells[i + 2, 5].Value = record.CustomerName;
            worksheet.Cells[i + 2, 6].Value = record.RegCountry;
            worksheet.Cells[i + 2, 7].Value = record.TotalValueAfterTaxing;
            worksheet.Cells[i + 2, 8].Value = record.TaxingValue;
            worksheet.Cells[i + 2, 9].Value = record.TotalValueBeforeTaxing;
        }
        int totalRow = records.Count + 2;
        worksheet.Cells[totalRow, 6].Value = "Total";
        worksheet.Cells[totalRow, 7].Formula = $"SUM(G2:G{records.Count + 1})";

        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

        return package.GetAsByteArray();
    }

}
