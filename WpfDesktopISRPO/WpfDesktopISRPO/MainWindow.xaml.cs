using OfficeOpenXml;
using OfficeOpenXml.Style;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
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
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace WpfDesktopISRPO
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        List<Sale> res = new List<Sale>();

        private void btnPost_Click(object sender, RoutedEventArgs e)
        {
            if (!DateTime.TryParse(tbDateStart.Text, out _) || !DateTime.TryParse(tbDateEnd.Text, out _)) return;
            cbDiagram.SelectedItem = null;
            lbSale.Items.Clear();

            var httpWebRequest = (HttpWebRequest)WebRequest.Create($"https://localhost:7100/api/Sale?dateStart={DateTime.Parse(tbDateStart.Text).ToString("MM.dd.yyyy")}&dateEnd={DateTime.Parse(tbDateEnd.Text).ToString("MM.dd.yyyy")}");

            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";//Можно GET
            httpWebRequest.ContentLength = 0;
            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                //ответ от сервера
                var result = streamReader.ReadToEnd();

                //using (StreamWriter writer = new StreamWriter("sad.json"))
                //{
                //    writer.WriteLine(result);
                //}
                //Сериализация
                res = JsonConvert.DeserializeObject<List<Sale>>(result);
                //    MessageBox.Show($"{res.Count}");
            }
            foreach (Sale sale in res)
            {

                lbSale.Items.Add(new Grid
                {
                    Children =
                    {
                        new ComboBoxItem{Content = sale,Visibility=Visibility.Hidden},
                        new Label{ Content=sale.client.lastName+" "+sale.client.firstName[0]+". "+sale.client.patronymic[0]+".\t"+sale.dateSale.ToShortDateString()},
                        new DataGrid
                        {
                            ItemsSource = sale.telephones,
                            Margin = new Thickness(30,30,0,0),
                            AutoGenerateColumns = false,
                            IsReadOnly=true,
                            MinWidth=630,
                            CanUserAddRows=false,
                            CanUserReorderColumns=false,
                            CanUserResizeColumns=false,
                            CanUserResizeRows=false,
                            CanUserDeleteRows=false,
                            Columns =
                            {
                                new DataGridTextColumn
                                {
                                    Header = "Артикул",
                                    Width = new DataGridLength(1,DataGridLengthUnitType.Auto),
                                    Binding = new Binding("articul")
                                },
                                new DataGridTextColumn
                                {
                                    Header = "Производитель",
                                    Width = new DataGridLength(1,DataGridLengthUnitType.Auto),
                                    Binding = new Binding("manufacturer")
                                },
                                new DataGridTextColumn
                                {
                                    Header = "Наименование",
                                    Width = new DataGridLength(1,DataGridLengthUnitType.Auto),
                                    Binding = new Binding("nameTelephone")
                                },
                                new DataGridTextColumn
                                {
                                    Header = "Количество",
                                    Width = new DataGridLength(1,DataGridLengthUnitType.Auto),
                                    Binding = new Binding("count")
                                },
                                new DataGridTextColumn
                                {
                                    Header = "Цена",
                                    Width = new DataGridLength(1,DataGridLengthUnitType.Auto),
                                    Binding = new Binding("cost")
                                }
                            }
                        }
                    }
                });
            }
        }

        private List<Telephone> GetTelephones(List<Sale> sales)
        {
            List<Telephone> telephones = new List<Telephone>();
            foreach (Sale s in sales)
            {
                telephones.AddRange(s.telephones);
            }
            return telephones;
        }

        //private int TelephonesCount()
        //{

        //}

        private void cbDiagram_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //  diagram = null;
            LineDiagram.Plot.Clear();
            LineDiagram.Refresh();
            PieDiagram.plt.Clear();
            PieDiagram.Refresh();
            LineDiagram.Visibility = Visibility.Hidden;
            PieDiagram.Visibility = Visibility.Hidden;
            switch (cbDiagram.SelectedIndex)
            {
                case 0:
                    PieDiagram.Visibility = Visibility.Visible;
                    List<Telephone> tele = GetTelephones(res);
                    List<ManufacturersCount> manufacturers = tele.GroupBy(t => t.manufacturer).Select(t => new ManufacturersCount
                    {
                        title = t.Key,
                        Quantity = t.Sum(s => s.count)
                    }).ToList();
                    double[] valuesPie = manufacturers.Select(m => (double)m.Quantity).ToArray();
                    string[] labelsPie = manufacturers.Select(m => m.title).ToArray();
                    var pie = PieDiagram.plt.AddPie(valuesPie);
                    pie.SliceLabels = labelsPie;
                    pie.ShowPercentages = true;
                    pie.ShowValues = false;
                    pie.ShowLabels = true;
                    PieDiagram.plt.Legend();
                    PieDiagram.Refresh();
                    break;
                case 1:
                    if (res.GroupBy(r => r.dateSale).Select(s => s.Key).Count() < 2) return;

                    LineDiagram.Visibility = Visibility.Visible;
                    List<double> valuesLine = new List<double>();
                    List<DateTime> dates = new List<DateTime>();
                    foreach (DateTime date in res.GroupBy(r => r.dateSale).Select(s => s.Key))
                    {
                        List<Telephone> telephone = GetTelephones(res.Where(r => r.dateSale == date).ToList());
                        valuesLine.Add(telephone.Select(t => (double)t.cost * t.count).Sum());
                        dates.Add(date);
                    }
                    double[] xs = dates.Select(x => x.ToOADate()).ToArray();
                    LineDiagram.Plot.AddScatter(xs, valuesLine.ToArray());
                    LineDiagram.Plot.XAxis.DateTimeFormat(true);
                    List<double> yPositions = new List<double>();
                    List<string> yLabels = new List<string>();
                    for (int i = 0; i <= Math.Round(valuesLine.Max()); i += (int)Math.Floor(Math.Round(valuesLine.Max()) - Math.Round(valuesLine.Min())) / 17)
                    {
                        yPositions.Add(i);
                        yLabels.Add(i.ToString());
                    }
                    yPositions.Add(Math.Floor(valuesLine.Max()));
                    yLabels.Add(Math.Floor(valuesLine.Max()).ToString());
                    yPositions[1] = Math.Floor(valuesLine.Min());
                    yLabels[1] = Math.Floor(valuesLine.Min()).ToString();
                    yPositions[2] = Math.Floor((yPositions[1] + yPositions[3]) / 2);
                    yLabels[2] = Math.Floor((yPositions[1] + yPositions[3]) / 2).ToString();
                    yPositions[18] = Math.Floor((yPositions[16] + yPositions[18]) / 2);
                    yLabels[18] = Math.Floor((yPositions[16] + yPositions[18]) / 2).ToString();
                    //foreach (double money in valuesLine.OrderBy(v=>v))
                    //{
                    //    if (!yPositions.Contains(money))
                    //    {
                    //        yPositions.Add(money);
                    //        yLabels.Add(money.ToString());
                    //    }
                    //}

                    LineDiagram.Plot.YAxis.ManualTickPositions(yPositions.ToArray(), yLabels.ToArray());
                    //diagram.Plot.
                    LineDiagram.Refresh();




                    break;
            }
        }

        private void btnChequeExcel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (lbSale.SelectedItem == null) return;
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Cheque.xlsx"))
            {
                try
                {
                    FileStream fs = File.Open(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Cheque.xlsx", FileMode.Open);
                    fs.Close();
                }
                catch
                {
                    MessageBox.Show("Файл Cheque.xlsx запущен на компьютере. Пожалуйста выключите его",
                        "Файл недоступен",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }
            }
            Grid item = (Grid)lbSale.SelectedItem;
            Sale sale = (item.Children[0] as ComboBoxItem).Content as Sale;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage package = new ExcelPackage();

            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Чек");
            sheet.Columns[1].Width = 4.43;
            sheet.Columns[2].Width = 3.29;
            sheet.Columns[3].Width = 10.57;
            sheet.Columns[4].Width = 1.71;
            sheet.Columns[5].Width = 0.17;
            sheet.Columns[6].Width = 17.57;
            sheet.Columns[7].Width = 16.14;
            sheet.Columns[8].Width = 4.43;
            sheet.Columns[9].Width = 8.71;
            sheet.Columns[10].Width = 0.08;
            sheet.Columns[11].Width = 5.86;
            sheet.Columns[12].Width = 2.86;
            sheet.Columns[13].Width = 4;
            sheet.Columns[14].Width = 6;
            sheet.Columns[15].Width = 0.08;
            sheet.Rows[1].Height = 15.25;
            sheet.Rows[2].Height = 9.25;
            sheet.Rows[3].Height = 3.75;
            sheet.Rows[4].Height = 23;
            sheet.Cells.Style.Font.Size = 9;
            sheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            sheet.Cells[1, 1, 1, 15].Merge = true;
            sheet.Cells[1, 1, 1, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[2, 1, 2, 15].Merge = true;
            sheet.Cells[2, 1, 2, 15].Value = "наименование организации, ИНН";
            sheet.Cells[2, 1, 2, 15].Style.Font.Size = 6;

            sheet.Cells[3, 1, 3, 15].Merge = true;
            sheet.Cells[4, 1, 4, 15].Merge = true;
            sheet.Cells[4, 1, 4, 15].Style.Font.Bold = true;
            sheet.Cells[4, 1, 4, 15].Style.Font.Size = 12;

            int Cheque = 1;
            if (File.Exists("int_i.txt"))
                using (StreamReader reader = new StreamReader("int_i.txt"))
                {
                    Cheque = int.Parse(reader.ReadToEnd());
                }
            sheet.Cells[4, 1, 4, 15].Value = $"Товарный чек № {Cheque} от {sale.dateSale.ToShortDateString()} г.";
            using (StreamWriter writer = new StreamWriter("int_i.txt", false))
            {
                writer.WriteLine(Cheque + 1);
            }
            sheet.Cells[5, 1].Value = "№ п/п";
            sheet.Cells[5, 1].Style.WrapText = true;

            sheet.Cells[5, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[5, 2, 5, 7].Merge = true;
            sheet.Cells[5, 2, 5, 7].Value = "Наименование, характеристика товара";
            sheet.Cells[5, 2, 5, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

            sheet.Cells[5, 2, 5, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[5, 8].Value = "Ед.";
            sheet.Cells[5, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[5, 9].Value = "Кол-во";
            sheet.Cells[5, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[5, 10, 5, 12].Merge = true;
            sheet.Cells[5, 13, 5, 14].Merge = true;
            sheet.Cells[5, 10, 5, 12].Value = "Цена";
            sheet.Cells[5, 10, 5, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[5, 13, 5, 14].Value = "Сумма";
            sheet.Cells[5, 13, 5, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            decimal sum = 0;
            int i = 1;
            foreach (Telephone telephone in sale.telephones)
            {
                sheet.Rows[i + 5].Height = 19;
                sheet.Cells[i + 5, 1].Value = i;
                sheet.Cells[i + 5, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[i + 5, 2, i + 5, 7].Merge = true;
                sheet.Cells[i + 5, 2, i + 5, 7].Value = telephone.nameTelephone + ", " + telephone.articul;
                sheet.Cells[i + 5, 2, i + 5, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[i + 5, 8].Value = "шт";
                sheet.Cells[i + 5, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[i + 5, 9].Value = telephone.count;
                sheet.Cells[i + 5, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[i + 5, 10, i + 5, 12].Merge = true;
                sheet.Cells[i + 5, 13, i + 5, 14].Merge = true;
                sheet.Cells[i + 5, 10, i + 5, 12].Value = telephone.cost;
                sheet.Cells[i + 5, 10, i + 5, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sum += telephone.cost * telephone.count;
                sheet.Cells[i + 5, 13, i + 5, 14].Value = telephone.cost * telephone.count;
                sheet.Cells[i + 5, 13, i + 5, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                i++;
            }
            i += 5;
            sheet.Rows[i].Height = 19;
            sheet.Cells[i, 1, i, 12].Merge = true;
            sheet.Cells[i, 1, i, 12].Value = "Всего";
            sheet.Cells[i, 1, i, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            sheet.Cells[i, 13, i, 14].Merge = true;

            sheet.Cells[i, 13, i, 14].Value = sum;
            sheet.Cells[i, 13, i, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            i++;
            sheet.Rows[i].Height = 3.75;
            sheet.Cells[i, 1, i, 15].Merge = true;
            i++;
            sheet.Rows[i].Height = 12.25;
            sheet.Cells[i, 1, i, 4].Merge = true;
            sheet.Cells[i, 1, i, 4].Value = "Всего отпущено на сумму:";
            sheet.Cells[i, 5, i, 14].Merge = true;
            sheet.Cells[i, 5, i, 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            i++;
            sheet.Rows[i].Height = 7;
            sheet.Cells[i, 1, i, 15].Merge = true;
            i++;
            sheet.Rows[i].Height = 0.75;
            sheet.Cells[i, 1, i, 15].Merge = true;
            i++;
            sheet.Rows[i].Height = 11.5;
            sheet.Cells[i, 1, i, 10].Merge = true;
            sheet.Cells[i, 1, i, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[i, 11].Value = "руб.";
            sheet.Cells[i, 12, i, 13].Merge = true;
            sheet.Cells[i, 12, i, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[i, 14].Value = "коп.";
            i++;
            sheet.Rows[i].Height = 0.75;
            sheet.Cells[i, 1, i, 15].Merge = true;
            i++;
            sheet.Rows[i].Height = 13.75;
            sheet.Cells[i, 1, i, 15].Merge = true;
            i++;
            sheet.Rows[i].Height = 11.5;
            sheet.Cells[i, 1, i, 2].Merge = true;
            sheet.Cells[i, 1, i, 2].Value = "Продавец";
            sheet.Cells[i, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[i, 4, i, 5].Merge = true;
            sheet.Cells[i, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[i, 7, i, 15].Merge = true;
            i++;
            sheet.Rows[i].Height = 11.5;
            sheet.Cells[i, 1, i, 2].Merge = true;
            sheet.Cells[i, 3].Value = "подпись";
            sheet.Cells[i, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            sheet.Cells[i, 4, i, 5].Merge = true;
            sheet.Cells[i, 6].Value = "ф.и.о.";
            sheet.Cells[i, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            sheet.Cells[i, 7, i, 15].Merge = true;
            sheet.Cells[i, 6].Style.Font.Size = 6;
            sheet.Cells[i, 3].Style.Font.Size = 6;
            sheet.Cells[1, 1, i, 15].Style.Border.BorderAround(ExcelBorderStyle.Medium, System.Drawing.Color.Blue);
            sheet.PrinterSettings.FitToPage = true;
            File.WriteAllBytes(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Cheque.xlsx", package.GetAsByteArray());
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Cheque.xlsx");





        }

        private void btnReportExcel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (lbSale.Items.Count < 1) return;
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Report.xlsx"))
            {
                try
                {
                    FileStream fs = File.Open(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Report.xlsx", FileMode.Open);
                    fs.Close();
                }
                catch
                {
                    MessageBox.Show("Файл Report.xlsx запущен на компьютере. Пожалуйста выключите его",
                        "Файл недоступен",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }
            }//01.01.2021
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage();

            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Отчет");
            sheet.Columns[1].Width = 9.5;
            sheet.Columns[2].Width = 14;
            sheet.Columns[3].Width = 9.5;
            sheet.Columns[4].Width = 8.5;
            sheet.Columns[5].Width = 10;
            sheet.Cells.Style.Font.Size = 10;

            sheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            int i = 1;
            sheet.Cells[i, 1, i, 5].Merge = true;
            sheet.Cells[i, 1, i, 5].Value = $"Отчет по продажам за период от {res.GroupBy(r => r.dateSale).Select(s => s.Key).Min().ToShortDateString()} до {res.GroupBy(r => r.dateSale).Select(s => s.Key).Max().ToShortDateString()}";
            i += 2;
            sheet.Cells[i, 1].Value = "Дата продажи";
            sheet.Cells[i, 1].Style.WrapText = true;
            sheet.Cells[i, 2].Value = "Клиент";
            sheet.Cells[i, 3].Value = "Количество";
            sheet.Cells[i, 4].Value = "Цена";
            sheet.Cells[i, 5].Value = "Сумма";
            sheet.Cells[i, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[i, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[i, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[i, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            i++;
            decimal sum = 0;
            foreach (Sale sales in res)
            {
                foreach (Telephone telephone in sales.telephones)
                {
                    sheet.Cells[i, 1].Value = sales.dateSale.ToShortDateString();
                    sheet.Cells[i, 2].Value = sales.client.lastName + " " + sales.client.firstName[0] + ". " + sales.client.patronymic[0] + ".";
                    sheet.Cells[i, 3].Value = telephone.count;
                    sheet.Cells[i, 4].Value = telephone.cost;
                    sheet.Cells[i, 5].Value = telephone.count * telephone.cost;
                    sum += telephone.count * telephone.cost;
                    sheet.Cells[i, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    i++;
                }
            }
            sheet.Cells[i, 4].Value = "Сумма";
            sheet.Cells[i, 5].Value = sum;
            File.WriteAllBytes(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Report.xlsx", package.GetAsByteArray());
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Report.xlsx");

        }

        private void btnReportWord_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // 01.01.2021
            if (lbSale.Items.Count < 1) return;
            int i = 0;
            foreach (Sale sales in res)
                foreach (Telephone telephone in sales.telephones)
                    i++;
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            var pText = doc.Paragraphs.Add();
            pText.Format.SpaceAfter = 10f;
            pText.Range.Text = $"Отчет по продажам за период от {res.GroupBy(r => r.dateSale).Select(s => s.Key).Min().ToShortDateString()} до {res.GroupBy(r => r.dateSale).Select(s => s.Key).Max().ToShortDateString()}";
            pText.Range.InsertParagraphAfter();

            // Insert table
            var pTable = doc.Paragraphs.Add();
            pTable.Format.SpaceAfter = 5;
            pTable.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // добавляем таблицу 10х3
            Word.Table tbl = app.ActiveDocument.Tables.Add(pTable.Range, i, 5, Word.WdDefaultTableBehavior.wdWord9TableBehavior);
            // делаем внутренние и внешние границы таблицы видимыми

            tbl.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            tbl.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            i = 1;
            tbl.Cell(i, 1).Range.Text = "Дата продажи";
            tbl.Cell(i, 2).Range.Text = "Клиент";
            tbl.Cell(i, 3).Range.Text = "Количество";
            tbl.Cell(i, 4).Range.Text = "Цена";
            tbl.Cell(i, 5).Range.Text = "Сумма";
            i++;
            decimal sum = 0;
            foreach (Sale sales in res)
            {
                foreach (Telephone telephone in sales.telephones)
                {
                    tbl.Cell(i, 1).Range.Text = sales.dateSale.ToShortDateString();
                    tbl.Cell(i, 2).Range.Text = sales.client.lastName + " " + sales.client.firstName[0] + ". " + sales.client.patronymic[0] + ".";
                    tbl.Cell(i, 3).Range.Text = $"{telephone.count}";
                    tbl.Cell(i, 4).Range.Text = $"{telephone.cost}";
                    sum += telephone.count * telephone.cost;
                    tbl.Cell(i, 5).Range.Text = $"{telephone.count * telephone.cost}";
                    i++;
                }
            }
            tbl.Range.Font.Size = 12;
            tbl.Columns.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
            doc.Paragraphs.Add();
            pText.Format.SpaceAfter = 10f;
            pText.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            pText.Range.Text = $"Итого {sum}";
            pText.Range.InsertParagraphAfter();
            app.Visible = true;
        }

        private void btnChequeWord_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (lbSale.SelectedItem == null) return;
            Grid item = (Grid)lbSale.SelectedItem;
            Sale sale = (item.Children[0] as ComboBoxItem).Content as Sale;
            int i = 0;
            foreach (Telephone telephone in sale.telephones)
                i++;
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            // Insert table
            var pHeader = doc.Paragraphs.Add();

            pHeader.Format.SpaceAfter = 5;
            Word.Table table = app.ActiveDocument.Tables.Add(pHeader.Range, 2, 3, Word.WdDefaultTableBehavior.wdWord9TableBehavior);
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
            table.Range.Font.Size = 9;
            table.Cell(1, 1).Range.Text = "Организаиция";
            int Cheque = 1;
            if (File.Exists("int_i.txt"))
                using (StreamReader reader = new StreamReader("int_i.txt"))
                {
                    Cheque = int.Parse(reader.ReadToEnd());
                }

            table.Cell(1, 2).Range.Text = $"ТОВАРНЫЙ ЧЕК\n№ {Cheque}";

            using (StreamWriter writer = new StreamWriter("int_i.txt", false))
            {
                writer.WriteLine(Cheque + 1);
            }
            table.Cell(1, 2).Range.Font.Bold = 2;
            table.Cell(1, 2).Range.Font.Size = 12;


            table.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Cell(1, 3).Range.Text = $"от {sale.dateSale.ToShortDateString()}г";
            table.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            table.Rows[2].Cells.Merge();
            table.Rows[1].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            table.Rows[2].Cells[1].Range.Text = "Продавец _______________________ Адрес _______________________________ ОГРН _________________________________________";
            table.Columns.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
            doc.Paragraphs.Add().Range.Font.Size = 1;
            // Insert table
            var pTable = doc.Paragraphs.Add();
            pTable.Format.SpaceAfter = 5;
            pTable.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // добавляем таблицу 10х3
            Word.Table tbl = app.ActiveDocument.Tables.Add(pTable.Range, i, 6, Word.WdDefaultTableBehavior.wdWord9TableBehavior);
            // делаем внутренние и внешние границы таблицы видимыми

            tbl.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            tbl.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            i = 1;

            tbl.Cell(i, 1).Range.Text = "Наименование товара";
            tbl.Cell(i, 2).Range.Text = "Артикул";
            tbl.Cell(i, 3).Range.Text = "Ед. изм.";
            tbl.Cell(i, 4).Range.Text = "Кол-во";
            tbl.Cell(i, 5).Range.Text = "Цена";
            tbl.Cell(i, 6).Range.Text = "Сумма";
            i++;
            decimal sum = 0;
            foreach (Telephone telephone in sale.telephones)
            {
                tbl.Cell(i, 1).Range.Text = telephone.nameTelephone;
                tbl.Cell(i, 2).Range.Text = telephone.articul.ToString();
                tbl.Cell(i, 3).Range.Text = "шт";
                tbl.Cell(i, 4).Range.Text = $"{telephone.count}";
                tbl.Cell(i, 5).Range.Text = $"{telephone.cost}";
                sum += telephone.count * telephone.cost;
                tbl.Cell(i, 6).Range.Text = $"{telephone.count * telephone.cost}";
                i++;
            }
            tbl.Range.Font.Size = 12;
            tbl.Rows[1].Range.Font.Italic = 1;
            tbl.Columns.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
            var pText = doc.Paragraphs.Add();
            pText.Format.SpaceAfter = 10f;
            pText.Range.Text = $"Итого {sum}\t\t\t\t\t\tПодпись ____________________";
            pText.Range.InsertParagraphAfter();

            app.Visible = true;
        }
    }
    public class ManufacturersCount
    {
        public string title { get; set; }
        public int Quantity { get; set; }

    }
    public class Client
    {
        public string lastName { get; set; }
        public string firstName { get; set; }
        public string patronymic { get; set; }
    }
    public class Telephone
    {
        public int articul { get; set; }
        public string nameTelephone { get; set; }
        public string category { get; set; }
        public decimal cost { get; set; }
        public int count { get; set; }
        public string manufacturer { get; set; }
    }
    public class Sale
    {
        public DateTime dateSale { get; set; }
        public Client client { get; set; }
        public List<Telephone> telephones { get; set; }

    }
}
