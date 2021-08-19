using ConsoleApp1;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication3.Service
{
    public class FileServices
    {
        private static FileServices instance;
        static object key = new object();
        public static FileServices Instance
        {
            get
            {
                //lock(key)
                //{
                if (instance == null)
                {
                    instance = new FileServices();
                }
                return instance;
                //}
            }
        }
        public FileServices()
        {

        }
        public byte[] ExportFileExcel(List<Subject> subjects, string[] header)
        {
            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    // đặt tên người tạo file
                    p.Workbook.Properties.Author = "Huy";

                    // đặt tiêu đề cho file
                    p.Workbook.Properties.Title = "Thoi Khoa Bieu Sinh Vien";

                    DateTime startDate = DateTime.ParseExact("16/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime endDate = DateTime.ParseExact("12/12/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    ExcelWorksheet ws = p.Workbook.Worksheets.Add($"ThoiKhoaBieuSV");
                    // fontsize mặc định cho cả sheet
                    ws.Cells.Style.Font.Size = 11;
                    // font family mặc định cho cả sheet
                    ws.Cells.Style.Font.Name = "Calibri";
                    //tự động ngắt dòng
                    ws.Cells.Style.WrapText = true;
                    
                    // lấy ra số lượng cột cần dùng dựa vào số lượng header
                    var countColHeader = header.Count();

                    // merge các column lại từ column 1 đến số column header
                    // gán giá trị cho cell vừa merge là Thống kê thông tni User Kteam
                    ws.Cells[1, 1].Value = "TIME TABLE";
                    ws.Cells[1, 1, 1, countColHeader].Merge = true;
                    ws.Cells[1, 1, 1, countColHeader].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    // in đậm
                    ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                    // căn giữa
                    ws.Cells[1, 1, 1, countColHeader].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[1, 1, 1, countColHeader].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //
                    ws.Cells[1, 1, 1, countColHeader].Style.Font.Size = 22;
                    int colIndex = 1;
                    int rowIndex = 2;
                    for (int i = 1; i <= countColHeader; i++)
                    {
                        if (i == 1)
                        {
                            ws.Column(i).Width = 25;
                        }
                        else ws.Column(i).Width = 30;
                    }
                    //tạo các header từ column header đã tạo từ bên trên
                    foreach (var item in header)
                    {
                        var cell = ws.Cells[rowIndex, colIndex];
                        //set màu thành gray
                        var fill = cell.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(100, 237, 237, 237);

                        //căn chỉnh các border
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);


                        //gán giá trị   
                        cell.Value = item;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cell.Style.Font.Size = 16;
                        colIndex++;
                    }
                    //tạo cột các tuần
                    colIndex = 1;
                    int weekIndex = 1;
                    for (DateTime k = startDate; k <= endDate; k = k.AddDays(7))
                    {
                        rowIndex++;
                        var cell = ws.Cells[rowIndex, colIndex];

                        //set màu
                        var fill = cell.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(100, 237, 237, 237);

                        //căn chỉnh các border
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        //gán giá trị
                        cell.Value = $"Week {weekIndex} " + (char)10 + "(" + String.Format("{0:dd/MM/yyyy}", k) + "-" + String.Format("{0:dd/MM/yyyy}", k.AddDays(6)) + ")";
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        weekIndex++;
                    }
                    // với mỗi item trong danh sách sẽ ghi trên 1 dòng

                    for (int j = 3; j < rowIndex + 1; j++)
                    {
                        List<Subject> listOfWeek = new List<Subject>();
                        foreach (var item in subjects)
                        {
                            var dateAt = DateTime.ParseExact(item.DateAt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                            if ((dateAt >= startDate && dateAt <= startDate.AddDays(6)) && startDate <= endDate)
                            {
                                listOfWeek.Add(item);
                            }
                        }
                        listOfWeek.ForEach(x => subjects.Remove(x));
                        startDate = startDate.AddDays(7);
                        for (int i = 2; i < 9; i++)
                        {
                            string subjectOfDay = "";
                            foreach (var item in listOfWeek)
                            {


                                if (item.weekDays.ToString().Contains(ws.Cells[2, i].Value.ToString()))
                                {
                                    subjectOfDay += $"{item.Name} ({item.Time})" + (char)10;
                                }

                            }
                            //gán giá trị cho từng cell
                            ws.Cells[j, i].Value = subjectOfDay;

                            var border = ws.Cells[j, i].Style.Border;
                            border.Bottom.Style = ExcelBorderStyle.Thick;
                            border.Top.Style =
                            border.Left.Style =
                            border.Right.Style = ExcelBorderStyle.Thin;

                            ws.Cells[j, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[j, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            var fill = ws.Cells[j, i].Style.Fill;
                            fill.PatternType = ExcelFillStyle.Solid;
                            if (j % 2 == 0)
                            {
                                fill.BackgroundColor.SetColor(100, 251, 216, 197);
                            }
                            else fill.BackgroundColor.SetColor(100, 255, 255, 255);

                        }
                        listOfWeek.RemoveRange(0, listOfWeek.Count);
                    }
                    //Lưu file lại
                    byte[] fileByte = p.GetAsByteArray();
                    return fileByte;
                }

            }
            catch (Exception EE)
            {
            }
            return default;
        }
        public List<Subject> ImportFileExcel(string filePath)
        {
            List<Subject> subjects = new List<Subject>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // lấy ra sheet đầu tiên để thao tác
                ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

                // duyệt tuần tự từ dòng thứ 2 đến dòng cuối cùng của file. lưu ý file excel bắt đầu từ số 1 không phải số 0
                for (int i = workSheet.Dimension.Start.Row + 10; i <= workSheet.Dimension.End.Row; i++)
                {
                    try
                    {
                        int j = 1;
                        if (workSheet.Cells[i, j].Value == null)
                        {
                            continue;
                        }
                        var _weekDays = (Subject.WeekDays)Convert.ToInt32(workSheet.Cells[i, j].Value);
                        j += 4;
                        var _Name = workSheet.Cells[i, j].Value != null ? workSheet.Cells[i, j].Value.ToString() : "";
                        j += 4;
                        var _Time = workSheet.Cells[i, j].Value != null ? workSheet.Cells[i, j].Value.ToString() : "";
                        j += 1;
                        var _Class = workSheet.Cells[i, j].Value != null ? workSheet.Cells[i, j].Value.ToString() : "";
                        j += 1;
                        var _DateAt = workSheet.Cells[i, j].Value != null ? workSheet.Cells[i, j].Value.ToString() : "";

                        string[] date = _DateAt.Split('-');
                        DateTime startDate = DateTime.ParseExact(date[0], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        DateTime endDate = DateTime.ParseExact(date[1], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        Subject subject;
                        for (DateTime k = startDate; k <= endDate; k = k.AddDays(1))
                        {
                            if (_weekDays.ToString().Contains(k.DayOfWeek.ToString()))
                            {
                                _DateAt = String.Format("{0:dd/MM/yyyy}", k);
                                subject = new Subject()
                                {
                                    weekDays = _weekDays,
                                    Name = _Name,
                                    Time = _Time,
                                    Class = _Class,
                                    DateAt = _DateAt
                                };
                                subjects.Add(subject);
                            }
                        }

                    }
                    catch (Exception exe)
                    {

                    }
                }
            }
            return subjects;
        }
    }
}
