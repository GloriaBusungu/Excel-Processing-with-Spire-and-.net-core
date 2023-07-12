using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProcessing.WritingExcel
{
    public class WriteExcelUsingSpire
    {
        //
        public static string CommentWriting(string TargetFolder, string Workbook, string ExcelSheetName,  string CommentMessage,int RowNo, int ColNo)
        {

            try
            {
                if (!Directory.Exists(TargetFolder))
                {
                    Directory.CreateDirectory(TargetFolder);
                }
                string outputPath = TargetFolder + Workbook;

                CultureInfo currentculture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

                var ExcelBook = new Workbook();

                /*Load the file from its location */
                ExcelBook.LoadFromFile(TargetFolder + "\\" + Workbook);
                Thread.CurrentThread.CurrentCulture = currentculture;

                //Sheet name to write comments
                Worksheet ExcelSheet = ExcelBook.Worksheets[ExcelSheetName];
                
              
                ExcelSheet.Range[RowNo, ColNo].Style.Color = Color.OrangeRed;
                ExcelSheet.Range[RowNo, ColNo].Style.Font.IsBold = true;
                ExcelSheet.Range[RowNo, ColNo].Style.Font.Color = Color.Black;
                ExcelSheet.Range[RowNo, ColNo].Style.Borders.Color = Color.DarkBlue;

                // Check if there is  comment in the selected cell, 
                if (ExcelSheet.Range[RowNo, ColNo].HasComment)
                {
                    // If there is comment in the selected cell, delete it and write the new comment
                    ExcelSheet.Range[RowNo, ColNo].Comment.Remove();
                    // Write the comment  after deleting the old comment
                    ExcelSheet.Range[RowNo, ColNo].Comment.RichText.Text = CommentMessage;
                    ExcelSheet.Range[RowNo, ColNo].Comment.Width = 600;
                    ExcelSheet.Range[RowNo, ColNo].Comment.Height = 200;
                }
                else
                {
                    ExcelSheet.Range[RowNo, ColNo].Comment.RichText.Text = CommentMessage;
                    ExcelSheet.Range[RowNo, ColNo].Comment.Width = 600;
                    ExcelSheet.Range[RowNo, ColNo].Comment.Height = 200;
                }

                ExcelBook.SaveToFile(outputPath);
                return "Success";
            }
          
            catch (Exception ex)
            {
                return "Failed" + ex.Message + ex.StackTrace;
            }

        }

    }
}
