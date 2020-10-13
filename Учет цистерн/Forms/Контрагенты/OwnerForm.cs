using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Учет_цистерн
{
    public partial class OwnerForm : Form
    {
        string role;

        public OwnerForm(string role)
        {
            InitializeComponent();
            this.role = role;
        }

        int SelectItemRow;

        //загрузка данных в DataGridView
        private void OwnerForm_Load(object sender, EventArgs e)
        {
            try
            {
                if (role == "1")
                {
                    btnOwnerAdd.Enabled = true;
                    btnOwnerUpdate.Enabled = true;
                    btnOwnerDelete.Enabled = true;
                    btnOwnerReffresh.Enabled = true;
                }
                else
                {
                    if (role == "2")
                    {
                        btnOwnerAdd.Enabled = true;
                        btnOwnerUpdate.Enabled = true;
                        btnOwnerDelete.Enabled = false;
                        btnOwnerReffresh.Enabled = true;
                    }
                    else
                    {
                        btnOwnerAdd.Enabled = true;
                        btnOwnerUpdate.Enabled = true;
                        btnOwnerDelete.Enabled = true;
                        btnOwnerReffresh.Enabled = true;
                    }
                }

                string Reffresh = "SELECT ID,Name,FullName FROM [dbo].[d__Owner]";
                DataTable dataTable = new DataTable();
                dataTable = DbConnection.DBConnect(Reffresh);
                gridControl1.DataSource = dataTable;
                gridView1.Columns[0].Visible = false;
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Добавление
        private void btnOwnerAdd_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(OwnerAddForm))
                    {
                        form.Activate();
                        return;
                    }
                }
                OwnerAddForm OwnerAddForm = new OwnerAddForm();
                OwnerAddForm.Owner = this;
                OwnerAddForm.ShowDialog();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Изменение
        private void btnOwnerUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(OwnerUpdtForm))
                    {
                        form.Activate();
                        return;
                    }
                }
                OwnerUpdtForm OwnerUpdtForm = new OwnerUpdtForm();
                OwnerUpdtForm.SelectID = SelectItemRow;
                OwnerUpdtForm.textBox1.Text = gridView1.GetFocusedDataRow()[1].ToString();
                OwnerUpdtForm.textBox2.Text = gridView1.GetFocusedDataRow()[2].ToString();
                //OwnerUpdtForm.textBox3.Text = dataGVOwner.CurrentRow.Cells[3].Value.ToString();
                OwnerUpdtForm.Owner = this;
                OwnerUpdtForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Для редактирования записи, необходимо указать строку! " + ex.Message,"",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
        }
        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                string Id = gridView1.GetFocusedDataRow()[0].ToString();
                SelectItemRow = Convert.ToInt32(Id);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Обновление
        public void btnOwnerReffresh_Click(object sender, EventArgs e)
        {
            try
            {
                string Reffresh = "SELECT ID,Name,FullName FROM [Batys].[dbo].[d__Owner]";
                DataTable dataTable = new DataTable();
                dataTable = DbConnection.DBConnect(Reffresh);
                gridControl1.DataSource = dataTable;
                gridView1.Columns[0].Visible = false;
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Удаление
        private void btnOwnerDelete_Click(object sender, EventArgs e)
        {
            if (gridView1.SelectedRowsCount > 0)
            {
                if (MessageBox.Show("Вы действительно хотите удалить эту запись?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        string Delete = "delete from d__Owner where ID = " + SelectItemRow;
                        DbConnection.DBConnect(Delete);
                        MessageBox.Show("Запись удалена!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btnOwnerReffresh_Click(null, null);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show(exp.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Для удаления записи, необходимо выбрать строку полностью!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void gridView1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            GridView View = sender as GridView;

            if (View.IsRowSelected(e.RowHandle))
            {
                e.Appearance.ForeColor = Color.DarkBlue;
                e.Appearance.BackColor = Color.LightBlue;
                //e.HighPriority = true;
            }
        }
    }
}
