using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;

namespace WindowsFormsApp1
{
    public partial class CalibrationManagerForm : Form
    {
        private List<CalibrationInfo> _infos;

        public CalibrationManagerForm()
        {
            InitializeComponent();

            InitGrid();

            this.Load += CalibrationManagerForm_Load;
            this.LoadButton.Click += LoadButton_Click;
            this.SaveButton.Click += SaveButton_Click;
            this.CalibrationGrid.CellValidating += CalibrationGrid_CellValidating;
        }

        private void CalibrationManagerForm_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void InitGrid()
        {
            CalibrationGrid.AutoGenerateColumns = false;
            CalibrationGrid.AllowUserToAddRows = false;
            CalibrationGrid.AllowUserToDeleteRows = false;
            CalibrationGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            CalibrationGrid.MultiSelect = false;

            CalibrationGrid.Columns.Clear();

            CalibrationGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Id",
                HeaderText = "Id",
                DataPropertyName = "Id",
                ReadOnly = true,
                Width = 50
            });

            CalibrationGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "InstrumentName",
                HeaderText = "儀器名稱",
                DataPropertyName = "InstrumentName",
                ReadOnly = true,
                Width = 280
            });

            CalibrationGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "CalibrationDate",
                HeaderText = "校正日期",
                DataPropertyName = "CalibrationDate",
                Width = 120,
                DefaultCellStyle = { Format = "yyyy.MM.dd" }
            });

            CalibrationGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ExpireDate",
                HeaderText = "有效日期",
                DataPropertyName = "ExpireDate",
                Width = 120,
                DefaultCellStyle = { Format = "yyyy.MM.dd" }
            });
        }

        private void LoadData()
        {
            try
            {
                _infos = InstrumentRepository.GetAll();

                CalibrationGrid.DataSource = null;
                CalibrationGrid.DataSource = _infos;
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取儀器資料失敗：" + ex.Message);
            }
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            try
            {
                CalibrationGrid.EndEdit();

                if (_infos == null || _infos.Count == 0)
                {
                    MessageBox.Show("沒有可儲存的資料");
                    return;
                }

                foreach (CalibrationInfo info in _infos)
                {
                    DateTime calDate;
                    DateTime expDate;

                    if (!DateTime.TryParse(info.CalibrationDate.ToString(), out calDate))
                    {
                        MessageBox.Show(info.InstrumentName + " 的校正日期格式錯誤");
                        return;
                    }

                    if (!DateTime.TryParse(info.ExpireDate.ToString(), out expDate))
                    {
                        MessageBox.Show(info.InstrumentName + " 的有效日期格式錯誤");
                        return;
                    }

                    if (expDate < calDate)
                    {
                        MessageBox.Show(info.InstrumentName + " 的有效日期不能早於校正日期");
                        return;
                    }

                    info.CalibrationDate = calDate;
                    info.ExpireDate = expDate;
                }

                foreach (CalibrationInfo info in _infos)
                {
                    InstrumentRepository.Update(info);
                }

                MessageBox.Show("儀器校正日期已更新");
                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存失敗：" + ex.Message);
            }
        }
        private void CalibrationGrid_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string colName = CalibrationGrid.Columns[e.ColumnIndex].Name;

            if (colName != "CalibrationDate" && colName != "ExpireDate")
                return;

            string text = e.FormattedValue == null ? "" : e.FormattedValue.ToString().Trim();

            DateTime dt;

            string[] formats = new string[]
            {
                "yyyy.MM.dd",
                "yyyy/MM/dd",
                "yyyy-MM-dd",
                "yyyy/M/d",
                "yyyy/M/dd",
                "yyyy/MM/d"
            };

            if (!DateTime.TryParseExact(
                text,
                formats,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out dt))
            {
                MessageBox.Show("日期格式錯誤，請輸入例如：2026.05.03");
                e.Cancel = true;
            }
        }
    }
}