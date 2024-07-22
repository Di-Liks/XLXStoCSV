using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ExcelDataReader;
using System.IO;
using System.Text.RegularExpressions;
using System.Numerics;
using System.Globalization;

namespace XLXStoCSV
{
    public partial class Form1 : Form
    {
        private string fileName = string.Empty;

        private DataTableCollection tableCollection = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;

                    Text = fileName;

                    OpenExcelFile(fileName);
                }
                else
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenExcelFile(string path)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = false
                }
            });

            tableCollection = db.Tables;

            toolStripComboBox1.Items.Clear();
            toolStripComboBox1.Items.Add("Всё");

            foreach (DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (toolStripComboBox1.SelectedItem.ToString() == "Всё")
            {
                DataTable combinedTable = new DataTable();

                foreach (DataTable table in tableCollection)
                {
                    if (table.Rows.Count < 4) continue;

                    DataRow headerRow = table.Rows[2];

                    if (combinedTable.Columns.Count == 0)
                    {
                        foreach (var item in headerRow.ItemArray)
                        {
                            combinedTable.Columns.Add(item.ToString());
                        }
                    }

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        if (i > 3)
                        {
                            DataRow newRow = combinedTable.NewRow();
                            newRow.ItemArray = table.Rows[i].ItemArray;
                            combinedTable.Rows.Add(newRow);
                        }
                    }
                }

                dataGridView.DataSource = combinedTable;
            }
            else
            {
                DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

                DataTable filteredTable = new DataTable();

                DataRow headerRow = table.Rows[2];
                foreach (var item in headerRow.ItemArray)
                {
                    filteredTable.Columns.Add(item.ToString());
                }

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (i > 3)
                    {
                        DataRow newRow = filteredTable.NewRow();
                        newRow.ItemArray = table.Rows[i].ItemArray;
                        filteredTable.Rows.Add(newRow);
                    }
                }

                dataGridView.DataSource = filteredTable;
            }
        }

        private void сформитоватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView.DataSource == null) return;

            DataTable originalTable = (DataTable)dataGridView.DataSource;
            DataTable newTable = new DataTable();

            newTable.Columns.Add("Населенный пункт", typeof(string));
            newTable.Columns.Add("Адрес/Район", typeof(string));
            newTable.Columns.Add("Даты", typeof(string));
            newTable.Columns.Add("Кол-во отключений", typeof(int));
            newTable.Columns.Add("Общее время отключения", typeof(TimeSpan));

            var combinedData = new Dictionary<string, (int count, TimeSpan totalDuration, List<string> datesList)>();

            foreach (DataRow row in originalTable.Rows)
            {
                if (string.IsNullOrWhiteSpace(row[0].ToString()) && string.IsNullOrWhiteSpace(row[5].ToString())) continue;
                if (string.IsNullOrWhiteSpace(row[4].ToString())) continue;
                if (string.IsNullOrWhiteSpace(row[6].ToString())) continue;
                if (string.IsNullOrWhiteSpace(row[7].ToString())) continue;

                string dateString1 = row[6].ToString();
                string dateString2 = row[7].ToString();

                if (DateTime.TryParseExact(dateString1, "dd.MM.2023 HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime date1) &&
                    DateTime.TryParseExact(dateString2, "dd.MM.2023 HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime date2))
                {
                    TimeSpan dateDifferenceHours = date2 - date1;

                    string townn = row[0].ToString();
                    string town = townn.Trim();
                    string strit = row[5].ToString();
                    string strite = strit.Trim();
                    if (strite != "")
                        town = $"{town}, {row[5]}";
                    string datestr = $"{row[6]} - {row[7]}";
                    string address = row[4].ToString();


                    List<string> adressList = new List<string>();


                    if (address.StartsWith("нас –")) continue;
                    address = address.Replace("\t", "");
                    if (address == "\"\"") continue;
                    string[] addres_split = address.Split(',').Where(addres => !addres.StartsWith(".")).ToArray();
                    foreach (var addres in addres_split)
                    {
                        string trimmedAddres = addres.Trim();
                        if (trimmedAddres.StartsWith(".")) continue;

                        if (string.IsNullOrWhiteSpace(trimmedAddres)) continue;

                        trimmedAddres = trimmedAddres.TrimEnd('.');

                        string type = GetPlaceType(trimmedAddres);
                        string combinedKey;

                        if (trimmedAddres.StartsWith("Переулка") || trimmedAddres.StartsWith("переулка"))
                        {
                            if (trimmedAddres.StartsWith("Переулка"))
                                trimmedAddres = trimmedAddres.Replace("Переулка ", "");
                            else
                                trimmedAddres = trimmedAddres.Replace("переулка ", "");
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }
                        else if (trimmedAddres.StartsWith("Проезда") || trimmedAddres.StartsWith("проезда"))
                        {
                            if (trimmedAddres.StartsWith("Проезда"))
                                trimmedAddres = trimmedAddres.Replace("Проезда ", "");
                            else
                                trimmedAddres = trimmedAddres.Replace("проезда ", "");
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }
                        else if (trimmedAddres.StartsWith("Второго переулка"))
                        {
                            trimmedAddres = trimmedAddres.Replace("Второго переулка ", "");
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }
                        else if (Regex.IsMatch(trimmedAddres, @"\bпервого\b|\bвторого\b|\bтретьего\b|\bчетвертого\b|\bпятого\b|\bшестого\b|\bседьмого\b|\bвосьмого\b|\bдевятого\b|\bдесятого\b"))
                        {
                            if (Regex.IsMatch(trimmedAddres, @"\bпервого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Первого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("первого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bвторого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Второго проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("второго проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bтретьего\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Третьего проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("третьего проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bчетвертого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Четвертого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("четвертого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bпятого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Пятого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("пятого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bшестого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Шестого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("шестого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bседьмого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Седьмого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("седьмого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bвосьмого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Восьмого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("восьмого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bдевятого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Девятого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("девятого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bдесятого\b\s+проезда"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Десятого проезда ", "");
                                trimmedAddres = trimmedAddres.Replace("десятого проезда ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bпервого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Первого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("первого переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bвторого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Второго переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("второго переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bтретьего\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Третьего переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("третьего переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bчетвертого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Четвертого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("четвертого переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bпятого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Пятого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("пятого переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bшестого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Шестого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("шестого переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bседьмого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Седьмого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("седьмого переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bвосьмого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Восьмого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("восьмого переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bдевятого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Девятого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("девятого переулка ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\bдесятого\b\s+переулка"))
                            {
                                trimmedAddres = trimmedAddres.Replace("Десятого переулка ", "");
                                trimmedAddres = trimmedAddres.Replace("десятого переулка ", "");
                            }
                            else if(trimmedAddres.EndsWith("первого")|| trimmedAddres.EndsWith("второго")|| trimmedAddres.EndsWith("третьего") || trimmedAddres.EndsWith("четвертого") || trimmedAddres.EndsWith("пятого"))
                            {
                                adressList.Add(trimmedAddres);
                                continue;
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"проездов"))
                            {
                                string[] number = trimmedAddres.Split(' ');
                                number[0] = number[0].Remove(number[0].Length - 3);
                                var numAddres = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(number[0]);
                                int n=adressList.Count;
                                string keyword = "проездов";
                                int index = trimmedAddres.IndexOf(keyword);
                                string resultAddres = trimmedAddres.Substring(index+9);
                                for (int i = 0; i < n; i++)
                                {
                                    if (adressList[i] == "первого")
                                    {
                                        trimmedAddres = "Первый проезд " + resultAddres;
                                        combinedKey = $"{town}-{type}{trimmedAddres}";
                                        if (combinedData.ContainsKey(combinedKey))
                                        {
                                            combinedData[combinedKey] = (
                                                combinedData[combinedKey].count + 1,
                                                combinedData[combinedKey].totalDuration + dateDifferenceHours,
                                                combinedData[combinedKey].datesList.Append(datestr).ToList()
                                            );
                                        }
                                        else
                                        {
                                            combinedData[combinedKey] = (1, dateDifferenceHours, new List<string> { datestr });
                                        }
                                    }
                                    if (adressList[i] == "второго")
                                    {
                                        trimmedAddres = "Второй проезд " + resultAddres;
                                        combinedKey = $"{town}-{type}{trimmedAddres}";
                                        if (combinedData.ContainsKey(combinedKey))
                                        {
                                            combinedData[combinedKey] = (
                                                combinedData[combinedKey].count + 1,
                                                combinedData[combinedKey].totalDuration + dateDifferenceHours,
                                                combinedData[combinedKey].datesList.Append(datestr).ToList()
                                            );
                                        }
                                        else
                                        {
                                            combinedData[combinedKey] = (1, dateDifferenceHours, new List<string> { datestr });
                                        }
                                    }
                                    if (adressList[i] == "третьего")
                                    {
                                        trimmedAddres = "Третий проезд " + resultAddres;
                                        combinedKey = $"{town}-{type}{trimmedAddres}";
                                        if (combinedData.ContainsKey(combinedKey))
                                        {
                                            combinedData[combinedKey] = (
                                                combinedData[combinedKey].count + 1,
                                                combinedData[combinedKey].totalDuration + dateDifferenceHours,
                                                combinedData[combinedKey].datesList.Append(datestr).ToList()
                                            );
                                        }
                                        else
                                        {
                                            combinedData[combinedKey] = (1, dateDifferenceHours, new List<string> { datestr });
                                        }
                                    }
                                    if (adressList[i] == "четвертого")
                                    {
                                        trimmedAddres = "Четверный проезд " + resultAddres;
                                        combinedKey = $"{town}-{type}{trimmedAddres}";
                                        if (combinedData.ContainsKey(combinedKey))
                                        {
                                            combinedData[combinedKey] = (
                                                combinedData[combinedKey].count + 1,
                                                combinedData[combinedKey].totalDuration + dateDifferenceHours,
                                                combinedData[combinedKey].datesList.Append(datestr).ToList()
                                            );
                                        }
                                        else
                                        {
                                            combinedData[combinedKey] = (1, dateDifferenceHours, new List<string> { datestr });
                                        }
                                    }
                                    if (adressList[i] == "пятого")
                                    {
                                        trimmedAddres = "Пятый проезд " + resultAddres;
                                        combinedKey = $"{town}-{type}{trimmedAddres}";
                                        if (combinedData.ContainsKey(combinedKey))
                                        {
                                            combinedData[combinedKey] = (
                                                combinedData[combinedKey].count + 1,
                                                combinedData[combinedKey].totalDuration + dateDifferenceHours,
                                                combinedData[combinedKey].datesList.Append(datestr).ToList()
                                            );
                                        }
                                        else
                                        {
                                            combinedData[combinedKey] = (1, dateDifferenceHours, new List<string> { datestr });
                                        }
                                    }
                                }
                                trimmedAddres = numAddres + "ый проезд " + resultAddres;
                                combinedKey = $"{town}-{type}{trimmedAddres}";
                                //combinedKey = $"{town}-{type}{trimmedAddres}";
                            }
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }
                        else
                        {
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }

                        if (combinedData.ContainsKey(combinedKey))
                        {
                            combinedData[combinedKey] = (
                                combinedData[combinedKey].count + 1,
                                combinedData[combinedKey].totalDuration + dateDifferenceHours,
                                combinedData[combinedKey].datesList.Append(datestr).ToList()
                            );
                        }
                        else
                        {
                            combinedData[combinedKey] = (1, dateDifferenceHours, new List<string> { datestr });
                        }
                    }
                }
            }

            foreach (var item in combinedData)
            {
                string[] parts = item.Key.Split('-');
                string town = parts[0];
                string address = parts[1];

                string dates = string.Join(", ", item.Value.datesList);

                newTable.Rows.Add(town, address, dates, item.Value.count, item.Value.totalDuration);
            }

            dataGridView.DataSource = newTable;
        }

        static string GetPlaceType(string place)
        {
            place = place.Trim();
            if (place.StartsWith("СНТ"))
            {
                return "";
            }
            else if (place.StartsWith("Садоводческое товарищество") || place.StartsWith("садоводческое товарищество") || place.EndsWith("Садоводческого товарищества") || place.EndsWith("садоводческого товарищества") || place.EndsWith("садоводческого товарищества") || place.EndsWith("садоводческbt товарищества"))
            {
                return "";
            }
            else if ((place.StartsWith("\"") && place.EndsWith("\"")))
            {
                return "Садоводческое товарищество ";
            }
            else if (place.StartsWith("переулок") || place.StartsWith("Переулок"))
            {
                return "";
            }
            else if (place.StartsWith("Переулка") || place.StartsWith("переулка"))
            {
                return "Переулок ";
            }
            else if (place.StartsWith("проезд") || place.StartsWith("Проезд"))
            {
                return "";
            }
            else if (place.StartsWith("Проезда") || place.StartsWith("проезда"))
            {
                return "Проезд ";
            }
            else if (Regex.IsMatch(place, @"\bпервого\b|\bвторого\b|\bтретьего\b|\bчетвертого\b|\bпятого\b|\bшестого\b|\bседьмого\b|\bвосьмого\b|\bдевятого\b|\bдесятого\b"))
            {
                if (Regex.IsMatch(place, @"\bпервого\b\s+проезда"))
                    return "Первый проезд ";
                else if (Regex.IsMatch(place, @"\bвторого\b\s+проезда"))
                    return "Второй проезд ";
                else if (Regex.IsMatch(place, @"\bтретьего\b\s+проезда"))
                    return "Третий проезд ";
                else if (Regex.IsMatch(place, @"\bчетвертого\b\s+проезда"))
                    return "Четверный проезд ";
                else if (Regex.IsMatch(place, @"\bпятого\b\s+проезда"))
                    return "Пятый проезд ";
                else if (Regex.IsMatch(place, @"\bшестого\b\s+проезда"))
                    return "Шестой проезд ";
                else if (Regex.IsMatch(place, @"\bседьмого\b\s+проезда"))
                    return "Седьмой проезд ";
                else if (Regex.IsMatch(place, @"\bвосьмого\b\s+проезда"))
                    return "Восьмой проезд ";
                else if (Regex.IsMatch(place, @"\bдевятого\b\s+проезда"))
                    return "Девятый проезд ";
                else if (Regex.IsMatch(place, @"\bдесятого\b\s+проезда"))
                    return "Десятый проезд ";
                else if (Regex.IsMatch(place, @"\bпервого\b\s+переулка"))
                    return "Первый переулок ";
                else if (Regex.IsMatch(place, @"\bвторого\b\s+переулка"))
                    return "Второй переулок ";
                else if (Regex.IsMatch(place, @"\bтретьего\b\s+переулка"))
                    return "Третий переулок ";
                else if (Regex.IsMatch(place, @"\bчетвертого\b\s+переулка"))
                    return "Четверный переулок ";
                else if (Regex.IsMatch(place, @"\bпятого\b\s+переулка"))
                    return "Пятый переулок ";
                else if (Regex.IsMatch(place, @"\bшестого\b\s+переулка"))
                    return "Шестой переулок ";
                else if (Regex.IsMatch(place, @"\bседьмого\b\s+переулка"))
                    return "Седьмой переулок ";
                else if (Regex.IsMatch(place, @"\bвосьмого\b\s+переулка"))
                    return "Восьмой переулок ";
                else if (Regex.IsMatch(place, @"\bдевятого\b\s+переулка"))
                    return "Девятый переулок ";
                else if (Regex.IsMatch(place, @"\bдесятого\b\s+переулка"))
                    return "Десятый переулок ";
                return "";
            }
            else if (place.StartsWith("Второго переулка"))
                return "Второй переулок ";
            else if (place.EndsWith("село") || place.EndsWith("Село"))
            {
                return "";
            }
            else if (place.EndsWith("перекат") || place.EndsWith("Перекат"))
            {
                return "";
            }
            else if (place.EndsWith("отделение") || place.EndsWith("Отделение"))
            {
                return "";
            }
            else if (place.EndsWith("набережная") || place.EndsWith("Набережная"))
            {
                return "";
            }
            else if (place.EndsWith("тупик") || place.EndsWith("Тупик"))
            {
                return "";
            }
            else if (place.EndsWith("Участок") || place.EndsWith("участок"))
            {
                return "";
            }
            else if (place.EndsWith("Шоссе") || place.EndsWith("шоссе"))
            {
                return "";
            }
            else if (place.EndsWith("дамба") || place.EndsWith("Дамба"))
            {
                return "";
            }
            else if (place.EndsWith("кольцо") || place.EndsWith("Кольцо"))
            {
                return "";
            }
            else
            {
                return "Улица ";
            }
        }


        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView.DataSource == null) return;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                ExportDataGridViewToCSV(filePath);
            }
        }

        private void ExportDataGridViewToCSV(string filePath)
        {
            var sb = new StringBuilder();

            var headers = dataGridView.Columns.Cast<DataGridViewColumn>();
            sb.AppendLine(string.Join(",", headers.Select(column => "\"" + column.HeaderText + "\"")));

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                var cells = row.Cells.Cast<DataGridViewCell>();
                sb.AppendLine(string.Join(",", cells.Select(cell => "\"" + cell.Value?.ToString().Replace("\"", "\"\"") + "\"")));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
    }
}
