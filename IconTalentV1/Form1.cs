using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace IconTalentV1
{
    public partial class wfMain : Form
    {
        private int numTab { get; set; }
        public wfMain()
        {
            InitializeComponent();
        }

        private TabPage CreateTabPage()
        {
            #region Creating left panel
            //Creating a export to excel button to tab pag
            Button btnExportToExcel = new Button();
            btnExportToExcel.Name = "btnExportToExcel";
            btnExportToExcel.Location = new System.Drawing.Point(4, 270);
            btnExportToExcel.Size = new System.Drawing.Size(145, 23);
            btnExportToExcel.TabIndex = 1;
            btnExportToExcel.Text = "Exportar p Excel";
            btnExportToExcel.UseVisualStyleBackColor = true;
            btnExportToExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnExportToExcel.Click += new System.EventHandler(btnExportToExcel_Click);
            btnExportToExcel.Visible = false;
            
            //Creating a close button to tab pag
            Button btnClose = new Button();
            btnClose.Name = "btnList";
            btnClose.Location = new System.Drawing.Point(4, 301);
            btnClose.Size = new System.Drawing.Size(145, 23);
            btnClose.TabIndex = 2;
            btnClose.Text = "Fechar Lista";
            btnClose.UseVisualStyleBackColor = true;
            btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnClose.Click += new System.EventHandler(btnClosePO_Click);

            //Creating a Tool Strip Menu Item to MenuStrip
            ToolStripMenuItem tsmMain = new ToolStripMenuItem();
            tsmMain.Name = "tsmMain" + numTab.ToString();
            tsmMain.Size = new System.Drawing.Size(59, 20);
            tsmMain.Text = "Produtos";

            //Creating five sub Tool Strip Menu Item to Tool Strip Menu Item tsm
            for (int i = 1; i <= 5; i++)
            {
                ToolStripMenuItem tsmSubLevel1 = new ToolStripMenuItem();
                tsmSubLevel1.Name = "tsm" + numTab.ToString() + i.ToString();
                for (int j = 1; j <= 5; j++)
                {
                    //Creating a sub Tool Strip Menu Item to Tool Strip Menu Item tsm
                    ToolStripMenuItem tsmSubLevel2 = new ToolStripMenuItem();
                    tsmSubLevel2.Name = "tsm" + numTab.ToString() + i.ToString() + j.ToString();
                    tsmSubLevel2.Size = new System.Drawing.Size(98, 22);
                    tsmSubLevel2.Text = "Produtos" + i.ToString() + j.ToString();
                    tsmSubLevel2.Click += new System.EventHandler(tsmSubLevel2_Click);
                    tsmSubLevel1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { tsmSubLevel2 });
                }
                tsmSubLevel1.Size = new System.Drawing.Size(114, 22);
                tsmSubLevel1.Text = "Produtos" + i.ToString();
                tsmMain.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { tsmSubLevel1 });
            }

            //Creating a MenuStrip to tab pag
            MenuStrip ms = new MenuStrip();
            ms.Name = "msMain" + numTab.ToString();
            ms.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {tsmMain});
            ms.Location = new System.Drawing.Point(0, 0);
            ms.Size = new System.Drawing.Size(152, 24);
            ms.TabIndex = 0;
            ms.ShowItemToolTips = true;

            //Creating panel to tab page
            Panel pnlLeft = new Panel();
            pnlLeft.Name = "pnlLeft" + numTab.ToString();
            pnlLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            pnlLeft.Controls.Add(btnExportToExcel);
            pnlLeft.Controls.Add(btnClose);
            pnlLeft.Controls.Add(ms);
            pnlLeft.Location = new System.Drawing.Point(8, 6);
            pnlLeft.Size = new System.Drawing.Size(154, 330);
            pnlLeft.TabIndex = 2;

            #endregion //Creating left panel

            #region Creating rigth panel

            //Creating Data Grid View Text Box Column
            DataGridViewTextBoxColumn Group = new DataGridViewTextBoxColumn();
            Group.HeaderText = "Group";
            Group.Name = "Group";
            Group.Width = 50;
            DataGridViewTextBoxColumn ID = new DataGridViewTextBoxColumn();
            ID.HeaderText = "ID";
            ID.Name = "ID";
            ID.Width = 50;
            DataGridViewTextBoxColumn Descrption = new DataGridViewTextBoxColumn();
            Descrption.HeaderText = "Descrption";
            Descrption.Name = "Descrption";
            Descrption.Width = 150;
            DataGridViewTextBoxColumn Value = new DataGridViewTextBoxColumn();
            Value.HeaderText = "Value";
            Value.Name = "Value";
            Value.Width = 50;
            DataGridViewTextBoxColumn Comments = new DataGridViewTextBoxColumn();
            Comments.HeaderText = "Comments";
            Comments.Name = "Comments";
            Comments.Width = 200;
            

            //creating a Data Grid View to rigth panel
            DataGridView dgvMain = new DataGridView();
            dgvMain.Name = "dgvMain" + numTab.ToString();
            dgvMain.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            dgvMain.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            dgvMain.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvMain.Location = new System.Drawing.Point(3, 3);
            dgvMain.Size = new System.Drawing.Size(612, 321);
            dgvMain.TabIndex = 0;
            dgvMain.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
                Group,
                ID,
                Descrption,
                Value,
                Comments}
            );

            //Creating a rigth panel to tab page
            Panel pnlRight = new Panel();
            pnlRight.Name = "pnlRight" + numTab.ToString();
            pnlRight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            pnlRight.Controls.Add(dgvMain);
            pnlRight.Location = new System.Drawing.Point(168, 7);
            pnlRight.Size = new System.Drawing.Size(612, 329);
            pnlRight.TabIndex = 1;

            #endregion //Creating rigth panel

            //Creating a tab page
            TabPage tpNew = new TabPage();
            tpNew.Text = "Lista " + numTab.ToString();
            tpNew.Controls.Add(pnlLeft);
            tpNew.Controls.Add(pnlRight);

            return tpNew;
        }

        private void btnNewDoc_Click(object sender, EventArgs e)
        {
            numTab++;
            tcMain.TabPages.Add(CreateTabPage());
            tcMain.SelectTab(tcMain.TabPages.Count - 1);
        }

        private void btnClosePO_Click(object sender, EventArgs e)
        {
            TabPage objTabPage = new TabPage();
            objTabPage = tcMain.SelectedTab;
            tcMain.TabPages.Remove(objTabPage);
        }

        private void tsmSubLevel2_Click(object sender, EventArgs e)
        {
            string stProducts = sender.ToString();
            TabPage tbSelected = tcMain.SelectedTab;

            //Creating list to data grid view
            List<ProductModel> lstProducts = new List<ProductModel>();
            Random randomNum = new Random();
            for (int i = 1; i <= 10; i++)
            {
                lstProducts.Add(new ProductModel() 
                { 
                    Group = stProducts, 
                    ID = i , 
                    Descrption = stProducts + i.ToString(), 
                    Value = Math.Round(randomNum.NextDouble(),2), 
                    Comments = i%2 == 0 ? "Product available" : "Product unavailable"
                });
            }

            //Changing dtagrid view            
            DataGridView dgvToChange = new DataGridView();
            Button btnToChange = new Button();
            foreach (Control ctMain in tbSelected.Controls)
            {
                foreach (Control ctSubLevel1 in ctMain.Controls)
                {
                    Type type = ctSubLevel1.GetType();
                    if (ctSubLevel1.GetType().Name == "DataGridView")
                    {
                        dgvToChange = (DataGridView)ctSubLevel1;
                        dgvToChange.Columns.Remove("Group");
                        dgvToChange.Columns.Remove("ID");
                        dgvToChange.Columns.Remove("Descrption");
                        dgvToChange.Columns.Remove("Value");
                        dgvToChange.Columns.Remove("Comments");
                        dgvToChange.DataSource = lstProducts;
                    }
                    if (ctSubLevel1.GetType().Name == "Button")
                    {
                        btnToChange = (Button)ctSubLevel1;
                        btnToChange.Visible = true;
                    }
                }
            }            
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Ocorreu uma exceção ao limpar / liberar o objeto " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            TabPage tbSelected = tcMain.SelectedTab;
            DataGridView dgvToCopy = new DataGridView();
            foreach (Control ctMain in tbSelected.Controls)
            {
                foreach (Control ctSubLevel1 in ctMain.Controls)
                {
                    Type type = ctSubLevel1.GetType();
                    if (ctSubLevel1.GetType().Name == "DataGridView")
                    {
                        dgvToCopy = (DataGridView)ctSubLevel1;
                    }
                }
            }

            dgvToCopy.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dgvToCopy.MultiSelect = true;
            dgvToCopy.SelectAll();
            DataObject dataObj = dgvToCopy.GetClipboardContent();

            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = tcMain.SelectedTab.Text+".xls";
            
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //Paste clipboard results to worksheet range
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                // Delete blank column A and select cell A1
                Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                delRng.Delete(Type.Missing);
                xlWorkSheet.get_Range("A1").Select();

                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                dgvToCopy.ClearSelection();

            }
            else
            {
                Clipboard.Clear();
                dgvToCopy.ClearSelection();
            }
        }
    }
}
