using PaymentExampleApp.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel =Microsoft.Office.Interop.Excel;
using Word =Microsoft.Office.Interop.Word;

namespace PaymentExampleApp.View.Pages
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class MainPage : Page
    {
        Core db = new Core();
        Excel.Application application;
        public MainPage()
        {
            InitializeComponent();
        }

        private void ReportButtonClick(object sender, RoutedEventArgs e)
        {
           
            //Запускаем приложение
             application = new Excel.Application();
            application.Visible = true;
            //Создание файла
            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);
            //Формируем массив
            var allUsers = db.context.Users.OrderBy(p => p.LastName).ToList();
            //Количество листов в книге
            application.SheetsInNewWorkbook = allUsers.Count();
            
            for (int i = 0; i < allUsers.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                worksheet.Name = allUsers[i].LastName;
                startRowIndex++;
                //вывод заголовков
                worksheet.Cells[1][startRowIndex] = "Дата платежа";
                worksheet.Cells[2][startRowIndex] = "Название";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                worksheet.Cells[4][startRowIndex] = "Количество";
                worksheet.Cells[5][startRowIndex] = "Сумма";

                var usersCategories = allUsers[i].Pay.OrderBy(p => p.Date_payment).GroupBy(p => p.category).OrderBy(p => p.Key.Name_category);

            foreach (var groupCategory in usersCategories)
            {
                Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                headerRange.Merge();
                headerRange.Value = groupCategory.Key.Name_category;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Italic = true;

                    startRowIndex++;

                foreach (var payments in groupCategory)
                {
                        worksheet.Cells[1][startRowIndex] = payments.Date_payment.Value.ToString("dd.MM.yyyy HH:mm");
                        worksheet.Cells[2][startRowIndex] = payments.Name;
                        worksheet.Cells[3][startRowIndex] = payments.Price;
                        worksheet.Cells[4][startRowIndex] = payments.Users;
                        worksheet.Cells[5][startRowIndex].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                }

            }
            
            }

        }

        private void PaymentButton_Click(object sender, RoutedEventArgs e)
        {
            var allUsers = db.context.Users.ToList();
            var allCategories = db.context.category.ToList();

            var application = new Word.Application();

            Word.Document document = application.Documents.Add();

            foreach (var user in allUsers)
            {
                Word.Paragraph userParagrapth = document.Paragraphs.Add();
                Word.Range userRange = userParagrapth.Range;
                userRange.Text = user.LastName+" "+ user.FirstName+" "+ user.ThirdName;
                userParagrapth.set_Style("Заголовок 2");
                userRange.InsertParagraphAfter();

                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count()+1, 3);
                paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Word.Range cellRange;

                cellRange = paymentsTable.Cell(1, 1).Range;
                cellRange.Text = "Иконка";
                cellRange = paymentsTable.Cell(1, 2).Range;
                cellRange.Text = "Категория";
                cellRange = paymentsTable.Cell(1, 3).Range;
                cellRange.Text = "Сумма Расходов";

                paymentsTable.Rows[1].Range.Bold = 1;
                paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                for (int i = 0; i < allCategories.Count(); i ++)
                {
                    var currentCategory = allCategories[i];

                    cellRange = paymentsTable.Cell(i + 2, 1).Range;
                    Word.InlineShape imageShape = cellRange.InlineShapes.AddPicture(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\Assets\\Icons\\" + currentCategory.Icons);
                imageShape.Width = imageShape.Height = 40;
                       cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = paymentsTable.Cell(i + 2, 2).Range;
                     cellRange.Text = currentCategory.Name_category;

                    cellRange = paymentsTable.Cell(i + 2, 3).Range;
                     cellRange.Text = user.Pay.ToList().Where(p => p.category == currentCategory).Sum(p => p.Count * p.Price).Value.ToString("N2") + " руб. ";

                Pay maxPayments = user.Pay.OrderByDescending(p => p.Price * p.Count).FirstOrDefault();
                if (maxPayments != null)
                {
                    Word.Paragraph maxPaymentsParagraph = document.Paragraphs.Add();
                    Word.Range maxPaymentsRange = maxPaymentsParagraph.Range;
                        maxPaymentsRange.Text = $"Самый дорогостоящий платеж - {maxPayments.Name} за {(maxPayments.Price * maxPayments.Count).Value.ToString("N2")}" +
                        $"руб. от {maxPayments.Date_payment.Value.ToString("dd.MM.yyyy HH:mm")}";
                    maxPaymentsParagraph.set_Style("Intense Quote");
                    maxPaymentsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    maxPaymentsRange.InsertParagraphAfter();
                }

                Pay minPayments = user.Pay.OrderBy(p => p.Price * p.Count).FirstOrDefault();
                if (minPayments != null)
                {
                    Word.Paragraph minPaymentsParagraph = document.Paragraphs.Add();
                    Word.Range minPaymentsRange = minPaymentsParagraph.Range;
                    minPaymentsRange.Text = $"Самый дешевый платеж - {minPayments.Name} за {(minPayments.Price * minPayments.Count).Value.ToString("N2")}" + $"руб. от {minPayments.Date_payment.Value.ToString("dd.MM.yyyy HH:mm")}";
                    minPaymentsParagraph.set_Style("Intense Quote");
                    minPaymentsRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                }

                if (user != allUsers.LastOrDefault())
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

            }



            document.SaveAs2(@":\Test.docx");
            document.SaveAs2(@":\Test.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }
            
        }
    }
}
