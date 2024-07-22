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

        private void �������ToolStripMenuItem_Click(object sender, EventArgs e)
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
                    throw new Exception("���� �� ������!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "������!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            toolStripComboBox1.Items.Add("��");

            foreach (DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (toolStripComboBox1.SelectedItem.ToString() == "��")
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

        private void ������������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView.DataSource == null) return;

            DataTable originalTable = (DataTable)dataGridView.DataSource;
            DataTable newTable = new DataTable();

            newTable.Columns.Add("���������� �����", typeof(string));
            newTable.Columns.Add("�����/�����", typeof(string));
            newTable.Columns.Add("����", typeof(string));
            newTable.Columns.Add("���-�� ����������", typeof(int));
            newTable.Columns.Add("����� ����� ����������", typeof(TimeSpan));

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


                    if (address.StartsWith("��� �")) continue;
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

                        if (trimmedAddres.StartsWith("��������") || trimmedAddres.StartsWith("��������"))
                        {
                            if (trimmedAddres.StartsWith("��������"))
                                trimmedAddres = trimmedAddres.Replace("�������� ", "");
                            else
                                trimmedAddres = trimmedAddres.Replace("�������� ", "");
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }
                        else if (trimmedAddres.StartsWith("�������") || trimmedAddres.StartsWith("�������"))
                        {
                            if (trimmedAddres.StartsWith("�������"))
                                trimmedAddres = trimmedAddres.Replace("������� ", "");
                            else
                                trimmedAddres = trimmedAddres.Replace("������� ", "");
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }
                        else if (trimmedAddres.StartsWith("������� ��������"))
                        {
                            trimmedAddres = trimmedAddres.Replace("������� �������� ", "");
                            combinedKey = $"{town}-{type}{trimmedAddres}";
                        }
                        else if (Regex.IsMatch(trimmedAddres, @"\b�������\b|\b�������\b|\b��������\b|\b����������\b|\b������\b|\b�������\b|\b��������\b|\b��������\b|\b��������\b|\b��������\b"))
                        {
                            if (Regex.IsMatch(trimmedAddres, @"\b�������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b�������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b����������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("���������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("���������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������ ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������ ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b�������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+�������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� ������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b�������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b�������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b����������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("���������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("���������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������ �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������ �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b�������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"\b��������\b\s+��������"))
                            {
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                                trimmedAddres = trimmedAddres.Replace("�������� �������� ", "");
                            }
                            else if(trimmedAddres.EndsWith("�������")|| trimmedAddres.EndsWith("�������")|| trimmedAddres.EndsWith("��������") || trimmedAddres.EndsWith("����������") || trimmedAddres.EndsWith("������"))
                            {
                                adressList.Add(trimmedAddres);
                                continue;
                            }
                            else if (Regex.IsMatch(trimmedAddres, @"��������"))
                            {
                                string[] number = trimmedAddres.Split(' ');
                                number[0] = number[0].Remove(number[0].Length - 3);
                                var numAddres = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(number[0]);
                                int n=adressList.Count;
                                string keyword = "��������";
                                int index = trimmedAddres.IndexOf(keyword);
                                string resultAddres = trimmedAddres.Substring(index+9);
                                for (int i = 0; i < n; i++)
                                {
                                    if (adressList[i] == "�������")
                                    {
                                        trimmedAddres = "������ ������ " + resultAddres;
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
                                    if (adressList[i] == "�������")
                                    {
                                        trimmedAddres = "������ ������ " + resultAddres;
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
                                    if (adressList[i] == "��������")
                                    {
                                        trimmedAddres = "������ ������ " + resultAddres;
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
                                    if (adressList[i] == "����������")
                                    {
                                        trimmedAddres = "��������� ������ " + resultAddres;
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
                                    if (adressList[i] == "������")
                                    {
                                        trimmedAddres = "����� ������ " + resultAddres;
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
                                trimmedAddres = numAddres + "�� ������ " + resultAddres;
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
            if (place.StartsWith("���"))
            {
                return "";
            }
            else if (place.StartsWith("������������� ������������") || place.StartsWith("������������� ������������") || place.EndsWith("�������������� ������������") || place.EndsWith("�������������� ������������") || place.EndsWith("�������������� ������������") || place.EndsWith("�����������bt ������������"))
            {
                return "";
            }
            else if ((place.StartsWith("\"") && place.EndsWith("\"")))
            {
                return "������������� ������������ ";
            }
            else if (place.StartsWith("��������") || place.StartsWith("��������"))
            {
                return "";
            }
            else if (place.StartsWith("��������") || place.StartsWith("��������"))
            {
                return "�������� ";
            }
            else if (place.StartsWith("������") || place.StartsWith("������"))
            {
                return "";
            }
            else if (place.StartsWith("�������") || place.StartsWith("�������"))
            {
                return "������ ";
            }
            else if (Regex.IsMatch(place, @"\b�������\b|\b�������\b|\b��������\b|\b����������\b|\b������\b|\b�������\b|\b��������\b|\b��������\b|\b��������\b|\b��������\b"))
            {
                if (Regex.IsMatch(place, @"\b�������\b\s+�������"))
                    return "������ ������ ";
                else if (Regex.IsMatch(place, @"\b�������\b\s+�������"))
                    return "������ ������ ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+�������"))
                    return "������ ������ ";
                else if (Regex.IsMatch(place, @"\b����������\b\s+�������"))
                    return "��������� ������ ";
                else if (Regex.IsMatch(place, @"\b������\b\s+�������"))
                    return "����� ������ ";
                else if (Regex.IsMatch(place, @"\b�������\b\s+�������"))
                    return "������ ������ ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+�������"))
                    return "������� ������ ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+�������"))
                    return "������� ������ ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+�������"))
                    return "������� ������ ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+�������"))
                    return "������� ������ ";
                else if (Regex.IsMatch(place, @"\b�������\b\s+��������"))
                    return "������ �������� ";
                else if (Regex.IsMatch(place, @"\b�������\b\s+��������"))
                    return "������ �������� ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+��������"))
                    return "������ �������� ";
                else if (Regex.IsMatch(place, @"\b����������\b\s+��������"))
                    return "��������� �������� ";
                else if (Regex.IsMatch(place, @"\b������\b\s+��������"))
                    return "����� �������� ";
                else if (Regex.IsMatch(place, @"\b�������\b\s+��������"))
                    return "������ �������� ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+��������"))
                    return "������� �������� ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+��������"))
                    return "������� �������� ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+��������"))
                    return "������� �������� ";
                else if (Regex.IsMatch(place, @"\b��������\b\s+��������"))
                    return "������� �������� ";
                return "";
            }
            else if (place.StartsWith("������� ��������"))
                return "������ �������� ";
            else if (place.EndsWith("����") || place.EndsWith("����"))
            {
                return "";
            }
            else if (place.EndsWith("�������") || place.EndsWith("�������"))
            {
                return "";
            }
            else if (place.EndsWith("���������") || place.EndsWith("���������"))
            {
                return "";
            }
            else if (place.EndsWith("����������") || place.EndsWith("����������"))
            {
                return "";
            }
            else if (place.EndsWith("�����") || place.EndsWith("�����"))
            {
                return "";
            }
            else if (place.EndsWith("�������") || place.EndsWith("�������"))
            {
                return "";
            }
            else if (place.EndsWith("�����") || place.EndsWith("�����"))
            {
                return "";
            }
            else if (place.EndsWith("�����") || place.EndsWith("�����"))
            {
                return "";
            }
            else if (place.EndsWith("������") || place.EndsWith("������"))
            {
                return "";
            }
            else
            {
                return "����� ";
            }
        }


        private void ���������ToolStripMenuItem_Click(object sender, EventArgs e)
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
