namespace Button;

using System.Security;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Net;

using Color = System.Drawing.Color;
using File = System.IO.File;
using System.Threading.Tasks;

using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using System.Xml.Linq;
using ClosedXML.Excel;


public class ExcelClass
{

    public class ExcelFile
    {
        public IXLAddress Cell { get; set; }
        public string Element { get; set; }

    }


    public List<ExcelFile> readExcelFile(string Path)
    {

        List<ExcelFile> File = new List<ExcelFile>();


        using var workbook = new XLWorkbook(Path);
        var ws = workbook.Worksheet(1);

        foreach (var row in ws.RowsUsed())
        {
            foreach (var cell in row.CellsUsed())
            {
                Console.WriteLine(cell.Address.ToString() + "  " + cell.GetString());
                File.Add(new ExcelFile
                {

                    Cell = cell.Address,
                    Element = cell.GetString()

                });


            }
        }
        return File;

    }


    public void WriteExcelFile(List<ExcelFile> file, string FileName)
    {
        using var workbook = new XLWorkbook();

        var worksheet = workbook.Worksheets.Add("Data");

        foreach (var item in file)
        {
            worksheet.Cell(item.Cell).Value = item.Element;
        }

        workbook.SaveAs(@"C:\Treiber\VICO_Tool\Files\" + FileName + ".xlsx");
    }

    public void CreateBasicTaglist(List<ExcelFile> file)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");

        IXLAddress address = worksheet.Cell("A1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "IsValid"
        });
        address = worksheet.Cell("B1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Tag"
        });
        address = worksheet.Cell("C1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Address"
        });
        address = worksheet.Cell("D1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Type"
        });
        address = worksheet.Cell("E1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Comment"
        });
        address = worksheet.Cell("F1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Usage"
        });
        address = worksheet.Cell("G1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Cycle"
        });
        address = worksheet.Cell("H1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "WriteAlways"
        });
        address = worksheet.Cell("I1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "HwIndex"
        });
        address = worksheet.Cell("J1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "AcyclicVariable"
        });
        address = worksheet.Cell("K1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "ForcingActive"
        });
        address = worksheet.Cell("L1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "PartOfCommunicationLoop"
        });
        address = worksheet.Cell("M1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Value"
        });
        address = worksheet.Cell("N1").Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "References"
        });
    }

    public void AddtoTaglist(List<ExcelFile> file, int counter, string Tag)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");

        IXLAddress address = worksheet.Cell("A" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "WAHR"
        });
        address = worksheet.Cell("B" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = Tag
        });
        address = worksheet.Cell("C" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = ""
        });
        address = worksheet.Cell("D" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Real"
        });
        address = worksheet.Cell("E" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "AxisValue"
        });
        address = worksheet.Cell("F" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Read"
        });
        address = worksheet.Cell("G" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "Continous"
        });
        address = worksheet.Cell("H" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "FALSCH"
        });
        address = worksheet.Cell("I" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = ""
        });
        address = worksheet.Cell("J" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "FALSCH"
        });
        address = worksheet.Cell("K" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "FALSCH"
        });
        address = worksheet.Cell("L" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "FALSCH"
        });
        address = worksheet.Cell("M" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "0"
        });
        address = worksheet.Cell("N" + counter).Address;
        file.Add(new ExcelFile
        {
            Cell = address,
            Element = "1"
        });
    }

}
