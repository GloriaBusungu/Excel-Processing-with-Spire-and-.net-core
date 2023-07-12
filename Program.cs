// See https://aka.ms/new-console-template for more information
using Spire.Xls;
using System.Drawing;

Console.WriteLine("Hello, World!");
/*Writing Excel Comments using Spire XLS*/

//1.Download Spire.XLS from Nugets
//2.Set Workbook Name, Excelsheetname, Target folder, RowNumber and Column Number as wel as Comment Message
string TargetFolder = "D:\\MyProjects\\GitHubProjects\\ExcelFiles\\";
string Workbook = "ExcelComments.xlsx";
string ExcelSheetName = "COMMENTSHEET";
string CommentMessage = "This is my first Excel Comments using spire.xls and .net core";
int RowNum = 2;
int ColNum = 2;
//Call Comment writing method
var WriteExcelCommentWithSpire = ExcelProcessing.WritingExcel.WriteExcelUsingSpire.CommentWriting(TargetFolder, Workbook, ExcelSheetName, CommentMessage, RowNum, ColNum);
//Check if it has succeeded
if (WriteExcelCommentWithSpire=="Success")
{
    Console.WriteLine("Hurray i just wrote my first comment in excel using spire. XLS and .net core 7");
}
else
{
    Console.WriteLine("My method failed let me troubleshoot");
}