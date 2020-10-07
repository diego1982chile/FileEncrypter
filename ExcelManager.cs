using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Reg17Generator
{
    class ExcelManager
    {
        //Leer excel de wsus summary
        XSSFWorkbook hssfwb;

        //Excel de salida -> REG17
        XSSFWorkbook xssfwb;        

        string inputFile;

        string outputFile;

        public ExcelManager(string inputFile, string outputFile)
        {
            this.inputFile = inputFile;
            this.outputFile = outputFile;

            using (FileStream file = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }                        

        }

        public void FillReg17SpreadSheet(Dictionary<string, Reg17Record> records)
        {
            xssfwb = new XSSFWorkbook();
            XSSFFont myFont = (XSSFFont)xssfwb.CreateFont();
            myFont.FontHeightInPoints = 12;
            myFont.Boldweight = (short) FontBoldWeight.Bold;
            //myFont.FontName = "Tahoma";

            XSSFFont myFont2 = (XSSFFont)xssfwb.CreateFont();
            myFont2.FontHeightInPoints = 18;
            myFont2.Boldweight = (short) FontBoldWeight.Bold;
            //myFont2.FontName = "Tahoma";

            IFont boldFont = xssfwb.CreateFont();
            boldFont.Boldweight = (short)FontBoldWeight.Bold;

            var color = new XSSFColor(new byte[] { 196, 215, 155 });

            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)xssfwb.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            borderedCellStyle.VerticalAlignment = VerticalAlignment.Center;

            XSSFCellStyle borderedCellStyle2 = (XSSFCellStyle)xssfwb.CreateCellStyle();
            borderedCellStyle2.SetFont(myFont2);
            
            borderedCellStyle2.VerticalAlignment = VerticalAlignment.Center;
            borderedCellStyle2.Alignment = HorizontalAlignment.Center;
            borderedCellStyle2.SetFillForegroundColor(color);
            borderedCellStyle2.SetFillBackgroundColor(color);
            borderedCellStyle2.FillBackgroundXSSFColor = color;
            borderedCellStyle2.FillForegroundXSSFColor = color;
            borderedCellStyle2.FillPattern = FillPattern.SolidForeground;

            XSSFCellStyle borderedCellStyle3 = (XSSFCellStyle)xssfwb.CreateCellStyle();
            borderedCellStyle3.SetFont(myFont);

            borderedCellStyle3.VerticalAlignment = VerticalAlignment.Center;
            borderedCellStyle3.Alignment = HorizontalAlignment.Center;
            borderedCellStyle3.SetFillForegroundColor(color);
            borderedCellStyle3.SetFillBackgroundColor(color);
            borderedCellStyle3.FillBackgroundXSSFColor = color;
            borderedCellStyle3.FillForegroundXSSFColor = color;
            borderedCellStyle3.FillPattern = FillPattern.SolidForeground;
            borderedCellStyle3.WrapText = true;

            ISheet Sheet = xssfwb.CreateSheet("Report");

            Sheet.SetColumnWidth(0, 22 * 256);
            Sheet.SetColumnWidth(1, 17 * 256);
            Sheet.SetColumnWidth(2, 25 * 256);
            Sheet.SetColumnWidth(3, 15 * 256);
            Sheet.SetColumnWidth(4, 96 * 256);
            Sheet.SetColumnWidth(5, 35 * 256);

            //Creat The Headers of the excel
            IRow row1 = Sheet.CreateRow(0);

            //styling
            ICellStyle boldStyle = xssfwb.CreateCellStyle();
            boldStyle.SetFont(boldFont);

            //Create The Actual Cells
            row1.CreateCell(0).SetCellValue("Sistema PCI DSS");
            row1.GetCell(0).CellStyle = boldStyle;

            IRow row2 = Sheet.CreateRow(1);

            row2.CreateCell(0).SetCellValue("REG-17");
            row2.GetCell(0).CellStyle = boldStyle;

            IRow row3 = Sheet.CreateRow(3);

            row3.CreateCell(0).SetCellValue("Informe de Aplicación de Parches");
            row3.GetCell(0).CellStyle = borderedCellStyle2;

            var cra = new NPOI.SS.Util.CellRangeAddress(3, 3, 0, 5);

            Sheet.AddMergedRegion(cra);

            row2.CreateCell(0).SetCellValue("REG-17");
            row2.GetCell(0).CellStyle = boldStyle;

            IRow row4 = Sheet.CreateRow(4);

            row4.CreateCell(0).SetCellValue("Servidores");
            row4.GetCell(0).CellStyle = borderedCellStyle;

            row4.CreateCell(1).SetCellValue("SQLPCI");
            row4.GetCell(1).CellStyle = boldStyle;

            row4.CreateCell(2).SetCellValue("Fecha de Aplicación");
            row4.GetCell(2).CellStyle = borderedCellStyle;

            row4.CreateCell(3).SetCellValue(DateTime.Now.ToString());
            row4.GetCell(3).CellStyle = borderedCellStyle;

            IRow row5 = Sheet.CreateRow(5);

            row5.CreateCell(1).SetCellValue("SRVFILE");
            row5.GetCell(1).CellStyle = boldStyle;

            IRow row6 = Sheet.CreateRow(6);

            row6.CreateCell(1).SetCellValue("Site-Transfer");
            row6.GetCell(1).CellStyle = boldStyle;

            IRow row7 = Sheet.CreateRow(7);

            row7.CreateCell(1).SetCellValue("WEBPCI");
            row7.GetCell(1).CellStyle = boldStyle;

            IRow row8 = Sheet.CreateRow(8);

            row8.CreateCell(1).SetCellValue("Proxy Inverso");
            row8.GetCell(1).CellStyle = boldStyle;

            IRow row9 = Sheet.CreateRow(9);

            row9.CreateCell(1).SetCellValue("Proxy Servicios");
            row9.GetCell(1).CellStyle = boldStyle;

            IRow row10 = Sheet.CreateRow(10);

            row10.CreateCell(1).SetCellValue("Proxy Web");
            row10.GetCell(1).CellStyle = boldStyle;

            IRow row11 = Sheet.CreateRow(11);

            row11.CreateCell(1).SetCellValue("Parches");
            row11.GetCell(1).CellStyle = boldStyle;

            IRow row12 = Sheet.CreateRow(12);

            row12.CreateCell(1).SetCellValue("ADMPCI");
            row12.GetCell(1).CellStyle = boldStyle;

            IRow row13 = Sheet.CreateRow(13);

            row13.CreateCell(1).SetCellValue("File Gateway");
            row13.GetCell(1).CellStyle = boldStyle;

            IRow row14 = Sheet.CreateRow(14);

            row14.CreateCell(1).SetCellValue("Volume Gateway");
            row14.GetCell(1).CellStyle = boldStyle;

            IRow row15 = Sheet.CreateRow(15);

            row15.CreateCell(1).SetCellValue("RD Gateway");
            row15.GetCell(1).CellStyle = boldStyle;

            IRow row16 = Sheet.CreateRow(16);

            row16.CreateCell(1).SetCellValue("Data Analysis");
            row16.GetCell(1).CellStyle = boldStyle;

            IRow row18 = Sheet.CreateRow(18);            

            row18.CreateCell(0).SetCellValue("Código Parche");
            row18.GetCell(0).CellStyle = borderedCellStyle3;

            row18.CreateCell(1).SetCellValue("Fecha Publicación");
            row18.GetCell(1).CellStyle = borderedCellStyle3;

            row18.CreateCell(2).SetCellValue("Producto");
            row18.GetCell(2).CellStyle = borderedCellStyle3;

            row18.CreateCell(3).SetCellValue("Clasificación");
            row18.GetCell(3).CellStyle = borderedCellStyle3;

            row18.CreateCell(4).SetCellValue("Mejoras y Correcciones");
            row18.GetCell(4).CellStyle = borderedCellStyle3;

            row18.CreateCell(5).SetCellValue("Opinión Impacto del Parche");
            row18.GetCell(5).CellStyle = borderedCellStyle3;

            XSSFFont myFont4 = (XSSFFont)xssfwb.CreateFont();
            myFont4.FontHeightInPoints = 10;
            //myFont4.Boldweight = (short)FontBoldWeight.Bold;
            //myFont.FontName = "Tahoma";

            XSSFFont myFont5 = (XSSFFont)xssfwb.CreateFont();
            myFont5.FontHeightInPoints = 8.5;
            //myFont5.Boldweight = (short)FontBoldWeight.Bold;
            //myFont2.FontName = "Tahoma";
            
            boldFont.Boldweight = (short)FontBoldWeight.Bold;
            
            XSSFCellStyle borderedCellStyle4 = (XSSFCellStyle) xssfwb.CreateCellStyle();
            borderedCellStyle4.SetFont(myFont4);
            borderedCellStyle4.VerticalAlignment = VerticalAlignment.Center;
            borderedCellStyle4.Alignment = HorizontalAlignment.Center;

            XSSFCellStyle borderedCellStyle5 = (XSSFCellStyle) xssfwb.CreateCellStyle();
            borderedCellStyle5.SetFont(myFont5);

            borderedCellStyle5.VerticalAlignment = VerticalAlignment.Center;
            //borderedCellStyle5.Alignment = HorizontalAlignment.Center;                        
            borderedCellStyle5.WrapText = true;

            XSSFCellStyle borderedCellStyle6 = (XSSFCellStyle)xssfwb.CreateCellStyle();
            borderedCellStyle6.SetFont(myFont4);

            borderedCellStyle6.VerticalAlignment = VerticalAlignment.Center;
            //borderedCellStyle6.Alignment = HorizontalAlignment.Center;
            borderedCellStyle6.WrapText = true;

            int i = 19;

            //List<Reg17Record> records = new List<Reg17Record>();
            foreach (KeyValuePair<string, Reg17Record> entry in records)
            {
                // do something with entry.Value or entry.Key
                Reg17Record record = (Reg17Record)entry.Value;

                IRow rowi = Sheet.CreateRow(i);

                rowi.CreateCell(0).SetCellValue(record.PatchCode);
                rowi.GetCell(0).CellStyle = borderedCellStyle4;

                rowi.CreateCell(1).SetCellValue(record.PublicationDate);
                rowi.GetCell(1).CellStyle = borderedCellStyle4;

                rowi.CreateCell(2).SetCellValue(record.Product);
                rowi.GetCell(2).CellStyle = borderedCellStyle4;

                rowi.CreateCell(3).SetCellValue(record.Classification);
                rowi.GetCell(3).CellStyle = borderedCellStyle4;

                rowi.CreateCell(4).SetCellValue(record.EnhancementsAndCorrections);
                rowi.GetCell(4).CellStyle = borderedCellStyle5;

                rowi.CreateCell(5).SetCellValue(record.ImpactOpinion);
                rowi.GetCell(5).CellStyle = borderedCellStyle6;

                i++;
            }

        }

        public void ExportReg17SpreadSheet()
        {
            // Write Excel to disk 
            using (var fileData = new FileStream(outputFile + ".xlsx", FileMode.Create))
            {
                xssfwb.Write(fileData);
            }
        }

        public Reg17Record ExtractRecord(string sheetName)
        {         

            ISheet sheet = hssfwb.GetSheet(sheetName);

            String patchCode = "";
            String publicationDate = "";
            String product = "";
            String classification = "";
            String enhancementsAndCorrections = "";
            String impactOpinion = "";

            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    try
                    {
                        string cellName = sheet.GetRow(row).GetCell(3).StringCellValue;

                        if (sheetName.Equals("Sheet2"))
                        {
                            if (row == 4)
                            {
                                publicationDate = cellName.Split(' ')[0];

                                if (publicationDate.Split('-').Length != 2)
                                {
                                    string year = DateTime.Now.Year.ToString();
                                    int _month = DateTime.Now.Month - 1;
                                    string month = _month.ToString();

                                    if (_month < 10)
                                    {
                                        month = "0" + month;
                                    }
                                    publicationDate = year + "-" + month;
                                }

                                try
                                {
                                    patchCode = cellName.Split('(')[1].Replace(")", "");

                                    if (!patchCode.StartsWith("KB"))
                                    {
                                        patchCode = cellName.Split('-')[1].Split(' ')[1];
                                    }
                                }                                
                                catch(Exception e)
                                {
                                    patchCode = cellName.Split('-')[1].Split(' ')[1];
                                }

                            }
                        } 
                        else
                        {
                            if (row == 0)
                            {
                                publicationDate = cellName.Split(' ')[0];

                                if (publicationDate.Split('-').Length != 2)
                                {
                                    string year = DateTime.Now.Year.ToString();
                                    int _month = DateTime.Now.Month - 1;
                                    string month = _month.ToString();

                                    if (_month < 10)
                                    {
                                        month = "0" + month;
                                    }
                                    publicationDate = year + "-" + month;
                                }

                                try
                                {
                                    patchCode = cellName.Split('(')[1].Replace(")", "");

                                    if(!patchCode.StartsWith("KB"))
                                    {
                                        patchCode = cellName.Split('-')[1].Split(' ')[1];
                                    }
                                }
                                catch (Exception e)
                                {
                                    patchCode = cellName.Split('-')[1].Split(' ')[1];
                                }
                            }
                        }

                        string cellValue = sheet.GetRow(row).GetCell(9).StringCellValue;

                        if (cellName.Equals("Classification:"))
                        {
                            classification = cellValue;
                        }
                        if (cellName.Equals("Products:"))
                        {
                            product = cellValue;
                        }
                        if (cellName.Equals("More Information:"))
                        {
                            WebSiteReader webSiteReader = new WebSiteReader();
                            enhancementsAndCorrections = webSiteReader.extractSummary(cellValue);
                            impactOpinion = webSiteReader.extractTitle(cellValue);
                            //patchCode = "KB" + cellValue.Split('/')[cellValue.Split('/').Length - 1];
                        }                        
                    }    
                    catch(Exception e)
                    {
                        System.Console.WriteLine(e.Message);
                    }
                }
            }

            return new Reg17Record(patchCode, publicationDate, product, classification, enhancementsAndCorrections, impactOpinion);
        }

        public void ExtractRecords(Dictionary<string, Reg17Record> records, string[] sheets)
        {
            //List<Reg17Record> records = new List<Reg17Record>();            
            for (int sheet = 1; sheet < hssfwb.NumberOfSheets - 1; sheet++)
            //for (int sheet = 1; sheet < hssfwb.NumberOfSheets - 25; sheet++)
            {
                //records.Add(ExtractRecord(hssfwb.GetSheetName(sheet)));
                if(sheets.Contains(hssfwb.GetSheetName(sheet)) || sheets.Length == 0)
                {
                    System.Console.WriteLine("sheet = " + sheet);
                    records[hssfwb.GetSheetName(sheet)] = ExtractRecord(hssfwb.GetSheetName(sheet));
                    Thread.Sleep(120000);
                }                
            }                         
        }

        public string[] getBadRecords(Dictionary<string, Reg17Record> records)
        {
            List<string> badRecords = new List<string>();

            //List<Reg17Record> records = new List<Reg17Record>();
            foreach (KeyValuePair<string, Reg17Record> entry in records)
            {
                // do something with entry.Value or entry.Key
                Reg17Record record = (Reg17Record) entry.Value;

                if (record.EnhancementsAndCorrections.Equals(""))
                {
                    badRecords.Add(entry.Key);
                }
            }

            return badRecords.ToArray();
        }

    }
}
