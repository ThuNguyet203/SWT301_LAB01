using OfficeOpenXml;
using System;
using System.IO;

namespace QuadraticEquationTester
{
    class Program
    {
        static void Main(string[] args)
        {
            // Thiết lập LicenseContext
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string filePath = "INPUT_LAB_1.xlsx";
            KiemThuPhuongTrinhBacHai(filePath);
        }

        static void KiemThuPhuongTrinhBacHai(string filePath)
        {
            // Kiểm tra nếu file tồn tại
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Không tìm thấy file.");
                return;
            }

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        Console.WriteLine("Workbook không chứa bất kỳ worksheet nào.");
                        return;
                    }

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Kiểm tra nếu dòng là trống thì bỏ qua
                        if (worksheet.Cells[row, 1].Value == null ||
                            worksheet.Cells[row, 2].Value == null ||
                            worksheet.Cells[row, 3].Value == null)
                        {
                            continue;
                        }

                        double a = Convert.ToDouble(worksheet.Cells[row, 1].Value);
                        double b = Convert.ToDouble(worksheet.Cells[row, 2].Value);
                        double c = Convert.ToDouble(worksheet.Cells[row, 3].Value);

                        GiaiPhuongTrinhBacHai(a, b, c);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Input invalid!");
            }
        }

        static void GiaiPhuongTrinhBacHai(double a, double b, double c)
        {
            if (a == 0 || a < 0 || a > 65535 || b < 0 || b > 65535 || c < 0 || c > 65535)
            {
                Console.WriteLine($"Input invalid!");
                return;
            }

            double delta = b * b - 4 * a * c;

            if (delta < 0)
            {
                Console.WriteLine("Phuong Trinh Khong Co Nghiem!");
            }
            else if (delta == 0)
            {
                double X1 = -b / (2 * a);
                Console.WriteLine($"Phong Trinh co nghiem kep: X1 = X2 = {X1}");
            }
            else
            {
                double sqrtDelta = Math.Sqrt(delta);
                double X1 = (-b + sqrtDelta) / (2 * a);
                double X2 = (-b - sqrtDelta) / (2 * a);
                Console.WriteLine($"Phuong Trinh co 2 nghiem: X1 = {X1}, X2 = {X2}");
            }
        }
    }
}
