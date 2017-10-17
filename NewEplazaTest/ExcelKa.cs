using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using System.Drawing.Imaging;
using System.IO;
using OpenQA.Selenium;
using System.Drawing;
using System.Net;
using System.Net.Mail;
using OpenQA.Selenium.Support.UI;

namespace NewEplazaTest
{
    public class ExcelKa
    {

       /* public void DisplayInExcel(IEnumerable<Order> orders)
        {
            var excelApp = new Excel.Application();
            Object missing = Type.Missing;
            excelApp.Visible = true;
            excelApp.Workbooks.Add(missing);
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            int[] Rep = new int[] { 4, 6, 8, 10, 12, 14, 2 };
            int[] nRep = new int[] { 3, 5, 7, 9, 11, 13, 15, 1 };

            //обЪединение ячеек
            for (int i = 0; i < 6; i++)
            {
                Excel.Range oRange1;
                oRange1 = workSheet.Range[workSheet.Cells[1, Rep[i]], workSheet.Cells[1, nRep[i]]];
                oRange1.Merge(Type.Missing);
            }

            string[] Arr = new string[] { "D", "F", "H", "J", "L", "N" };
            string[] Arr2 = new string[] { "C", "E", "G", "I", "K", "M", "O", "A", "B" };
            string[] Arr3 = new string[] { "Cтоимость товара(ов)", "Стоимость доставки", "Стоимость заказа", "Заказ в админке", "САП", "Uniteller", "Ссылка на заказ", "Номер заказа", "Наименование товара" };
            string[] Arr4 = new string[] { "должно быть", " факт " };
            for (int i = 0; i < 9; i++)
            {
                workSheet.Cells[1, Arr2[i]] = Arr3[i];
            }
            for (int i = 0; i < 6; i++)
            {
                workSheet.Cells[2, Arr2[i]] = Arr4[0];
            };
            for (int i = 0; i < 6; i++)
            {
                workSheet.Cells[2, Arr[i]] = Arr4[1];
            };
            var row = 2;
            foreach (var ord in orders)
            {
                row++;
                workSheet.Cells[row, "A"] = ord.OrderId;
                workSheet.Cells[row, "B"] = ord.ProdName;
                workSheet.Cells[row, "D"] = ord.ProductValue;
                workSheet.Cells[row, "F"] = ord.DeliveryValue;
                workSheet.Cells[row, "H"] = ord.OrderValue;
                workSheet.Cells[row, "J"] = ord.AdminOrderValue;
                workSheet.Cells[row, "L"] = ord.Sap;
                workSheet.Cells[row, "N"] = ord.Uniteller;
                workSheet.Cells[row, "C"] = ord.nProductValue;
                workSheet.Cells[row, "E"] = ord.nDeliveryValue;
                workSheet.Cells[row, "G"] = ord.nOrderValue;
                workSheet.Cells[row, "I"] = ord.nAdminOrderValue;
                workSheet.Cells[row, "K"] = ord.nSap;
                workSheet.Cells[row, "M"] = ord.nUniteller;
                workSheet.Cells[row, "O"] = ord.UrlAdress;
                workSheet.Cells[row, "P"] = " ";
            }

            //цвет текста          
            for (int i = 0; i < 6; i++)
            {
                if (workSheet.Cells[3, Arr2[i]].FormulaLocal == workSheet.Cells[3, Arr[i]].FormulaLocal)
                {
                    Excel.Range rng2 = workSheet.get_Range(Arr[i] + "3");
                    rng2.Font.Color = ColorTranslator.ToOle(Color.Green);
                }
                else
                {
                    Excel.Range rng2 = workSheet.get_Range(Arr[i] + "3");
                    rng2.Font.Color = ColorTranslator.ToOle(Color.Red);
                }
            }

            //редактирование ячеек 
            workSheet.Range["A1", "O2"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
            workSheet.Range["A3", "O3"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //границы ячеек 
            for (int i = 1; i < 4; i++)
            {
                Excel.Range rt = workSheet.get_Range("A" + i, "O" + i);
                rt.Borders.ColorIndex = 0;
                rt.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                rt.Borders.Weight = Excel.XlBorderWeight.xlThin;
            }

            //перенос текста
            workSheet.Cells[3, "B"].WrapText = true;

            //сохранение отчета
            excelApp.DisplayAlerts = false;
            workSheet.SaveAs(string.Format(@"{0}\Test.xlsx", Environment.CurrentDirectory));
            excelApp.Quit();

            //отправка отчета
            SendMail mail = new SendMail();
            mail.Mail();
        }*/
    }
}

