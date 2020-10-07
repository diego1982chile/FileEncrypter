using HtmlAgilityPack;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Reg17Generator
{
    class Program
    {        

        static void Main(string[] args)
        {
            //Leer excel de wsus summary
            ExcelManager excelManager = new ExcelManager(@"Updates_Report_for_PARCHES_Septiembre_2020.xlsx", @"REG-17_Aplicación_de_Parches_Sep_2020");            

            //excelManager.ExportReg17SpreadSheet();

            Dictionary<string, Reg17Record> records = new Dictionary<string, Reg17Record>();
            excelManager.ExtractRecords(records, new string[0]);

            while(excelManager.getBadRecords(records).Length > 0)
            {
                excelManager.ExtractRecords(records, excelManager.getBadRecords(records));
            }

            //Iterar Por cada parche (hoja de excel)
            excelManager.FillReg17SpreadSheet(records);

            excelManager.ExportReg17SpreadSheet();
        }

        

        
    }
}
