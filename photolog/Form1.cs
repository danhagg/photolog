using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Data;
using System.Xml.Linq;
using System.Windows.Forms.DataVisualization.Charting;
using System.Text;
using System.Linq;
using System.Collections.Generic;

namespace photolog
{

    public partial class Form1 : Form
    {

        const int exifOrientationID = 0x112; //274

        public Form1()
        {
            InitializeComponent();

            this.FormClosing += Form1_FormClosing;

            //Create right click menu using contextmenustrip for right click move top bottom
            ContextMenuStrip s = new ContextMenuStrip();     
            ToolStripMenuItem top = new ToolStripMenuItem();
            ToolStripMenuItem bottom = new ToolStripMenuItem();
            top.Text = @"Send to TOP";
            bottom.Text = @"Send to BOTTOM";
            top.Click += top_Click;
            bottom.Click += bottom_Click;
            s.Items.Add(top);
            s.Items.Add(bottom);
            this.ContextMenuStrip = s;

            this.dataGridView1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseDown);
            dataGridView1.RowHeaderMouseClick += new DataGridViewCellMouseEventHandler(OnRowHeaderMouseClick);

            this.AllowDrop = true;
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
            this.dataGridView1.DragOver += new DragEventHandler(dataGridView1_DragOver);
            this.dataGridView1.DragDrop += new DragEventHandler(dataGridView1_DragDrop);
            this.dataGridView1.DragEnter += new DragEventHandler(dataGridView1_DragEnter);
            this.dataGridView1.AllowDrop = true;
            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            this.dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.RaisedHorizontal;
        }


        // Set Form listView and datGridView properties on load
        private void Form1_Load(object sender, EventArgs e)
        {
            this.KeyUp += new KeyEventHandler(KeyEvent);
            this.KeyPreview = true;

            // Form Size and Scale
            this.Load += new EventHandler(Form1_Load);
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.Sizable;

            // dataGridView1
            dataGridView1.RowTemplate.Height = 64;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowDrop = true;
            dataGridView1.MultiSelect = true;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[1].DefaultCellStyle.Font = new System.Drawing.Font("Verdana", 10F);
            dataGridView1.Columns[2].DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 10F);
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Blue;
            dataGridView1.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);

            // pictureBox1
            pictureBox1.SizeMode = PictureBoxSizeMode.CenterImage;

            // Tool tips
            ToolTip toolTip1 = new ToolTip();
            toolTip1.SetToolTip(button2, "Make a Word document");

            // Chart
            chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisX.ScrollBar.Enabled = true;

            // photolog version           
            label5.Text = @"PhotoLog v1.4";
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show(@"You  are exiting Photolog. You may lose unsaved  work. Are you sure you want to quit?", @"Photolog warning", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
            }
            else
            {
                e.Cancel = true;
            }
        }


        // MENU - Save
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var path = textBox6.Text;

            if (path == "")
            {
                var dS = new DataSet();
                var dT2 = GetDataTableFromDGV1(dataGridView1);
                dS.Tables.Add(dT2);
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = @"photolog file (*.photolog)|*.photolog",
                    DefaultExt = "photolog",
                    AddExtension = true,
                    Title = @"Save work as .photolog file"
                };
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    dS.WriteXml(saveFileDialog.FileName);
                    textBox6.Text = saveFileDialog.FileName;
                }
            }
            else
            {
                var dS = new DataSet();
                var dT2 = GetDataTableFromDGV1(dataGridView1);
                dS.Tables.Add(dT2);
                dS.WriteXml(textBox6.Text);
                AutoClosingMessageBox.Show(textBox6.Text, "Saved", 3000);
            }
        }

        // Append
        private void addMultipleTempFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resume(sender, e);
        }

        // MENU - Resume
        private void resumeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            resume(sender, e);
        }


        // MENU - Change My parent folder
        private void changeParentFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //const int exifOrientationID = 0x112; //274
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = @"photolog files (*.photolog; *.XML)| *.photolog; *.XML",
                FilterIndex = 1,
                Title = @"Select .photolog file. The .photolog file must be in same folder as your images for this to work."
            };

            // check user selects pass
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xmlFileName = ofd.FileName;
                string xmlFileDirectory = Path.GetDirectoryName(xmlFileName);

                dataGridView1.Rows.Clear();

                // Read and load info from XML 
                XDocument doc = XDocument.Load(xmlFileName);
                if (doc.Root.Elements().Any())
                {
                    try
                    {
                        foreach (var dm2 in doc.Descendants("Table1"))
                        {
                            string fileName = dm2.Element("fileName")?.Value;
                            string temp = dm2.Element("dataGridView1Path")?.Value.ToString();
                            string fileExt = Path.GetFileName(temp);
                            string fileNameFull = xmlFileDirectory + '\\' + fileExt;
                            Image img = Image.FromFile(fileNameFull);

                            if (!img.PropertyIdList.Contains(exifOrientationID))
                            {
                                var capt = dm2.Element("dataGridView1Caption")?.Value;
                                var pth = fileNameFull;
                                dataGridView1.Rows.Add(fileName, img, capt, pth);
                            }
                            else
                            {

                                RotImage(img);

                                var capt = dm2.Element("dataGridView1Caption")?.Value;
                                var pth = fileNameFull;
                                dataGridView1.Rows.Add(fileName, img, capt, pth);
                            }                             

                        }
                        //capLength();
                        fileSize();
                        fileSizeTotal();
                        updatePictureBox();
                        textBox6.Text = xmlFileName;
                        BarExample();
                    }
                    // Error
                    catch (Exception exc)
                    {
                        StringBuilder myStringBuilder = new StringBuilder("You tried to load images from the following saved project: \n\n");
                        myStringBuilder.Append(xmlFileName + "\n\n");
                        myStringBuilder.Append("The saved project tried to load the following file: \n\n");
                        myStringBuilder.Append(exc.Message + "\n\n");
                        myStringBuilder.Append("But this file does not exist in this folder location. Have you perhaps moved it? Or has it been renamed? \n\n");
                        MessageBox.Show(myStringBuilder.ToString());
                    }
                }
                else
                {
                    StringBuilder myStringBuilder = new StringBuilder("You tried to load images from the following saved project: \n\n");
                    myStringBuilder.Append(xmlFileName + "\n\n");
                    myStringBuilder.Append("But this file is empty. \n\n");
                    MessageBox.Show(myStringBuilder.ToString());
                }
            }

        }



        // Overlay file name on top of image      
        private void dataGridView1_CellPainting(object sender,
                                DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;                  // no image in the header
            if (e.ColumnIndex == dataGridView1.Columns["imageColumn"].Index)

            {
                e.PaintBackground(e.ClipBounds, false);  // no highlighting
                e.PaintContent(e.ClipBounds);

                int y = e.CellBounds.Bottom - 17;         // your  font height
                string mystring = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                var result = mystring.Substring(mystring.Length - Math.Min(4, mystring.Length));

                System.Drawing.Rectangle rect = new System.Drawing.Rectangle(e.CellBounds.Location.X + 1, e.CellBounds.Location.Y + 49, 35, 14);
                e.Graphics.FillRectangle(Brushes.White, rect);
                e.Graphics.DrawString(result, e.CellStyle.Font,
                Brushes.Crimson, e.CellBounds.Left, y);

                e.Handled = true;                        // done with the image column 
            }
        }


        // Paint the row headers
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }



        // MENU - Drag & Drop Individual Images
        void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            System.Drawing.Point clientPoint = dataGridView1.PointToClient(new System.Drawing.Point(e.X, e.Y));
            // Get the row index of the item the mouse is below. 
            int rowIndexOfItemUnderMouseToDrop =
                dataGridView1.HitTest(clientPoint.X, clientPoint.Y).RowIndex;

            // Make an array of all files being dragged in
            string[] fileNames = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (fileNames != null && fileNames.Length != 0)
            {
                // Make an array of all files in the dataGridView
                var array0 = dataGridView1.Rows.Cast<DataGridViewRow>()
                                             .Select(x => x.Cells[3].Value.ToString().Trim()).ToArray();

                // Which files "collide"
                var intersection = array0.Intersect(fileNames);

                string[] remainderTransfer = fileNames.Where(x => !intersection.Contains(x)).ToArray();

                // If collisions exist
                if (intersection.Count() > 0)
                {
                    StringBuilder sb = new StringBuilder("The following file(s) are already present in the list and were NOT inserted: \n\n");
                    foreach (string id in intersection)
                    {
                        sb.Append(id + "\n");

                    }
                    sb.Append("\nHowever, \n\n" + "The following file(s) WERE inserted: \n\n");
                    foreach (string id in remainderTransfer)
                    {
                        sb.Append(id + "\n");
                    }
                    MessageBox.Show(sb.ToString());
                    // drop is into empty row
                    if (rowIndexOfItemUnderMouseToDrop == -1)
                    {
                        int totalRows = dataGridView1.Rows.Count;
                        for (int i = 0; i < remainderTransfer.Length; i++)
                        {
                            string fFull = Path.GetFullPath(remainderTransfer[i]);
                            string fileNam = Path.GetFileNameWithoutExtension(remainderTransfer[i]);
                            Image img = Image.FromFile(fFull);
                            if (!img.PropertyIdList.Contains(exifOrientationID))
                            {
                                Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                dataGridView1.Rows.Insert(totalRows + i, row);
                            }
                            else
                            {
                                RotImage(img);
                                Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                dataGridView1.Rows.Insert(totalRows + i, row);
                            }
                            
                        }
                        dataGridView1.CurrentCell = this.dataGridView1[1, totalRows];
                        //capLength();
                        fileSize();
                        fileSizeTotal();
                        updatePictureBox();
                        BarExample();

                    }
                    // drop into existing row
                    else
                    {
                        for (int i = 0; i < remainderTransfer.Length; i++)
                        {
                            string fFull = Path.GetFullPath(remainderTransfer[i]);
                            string fileNam = Path.GetFileNameWithoutExtension(remainderTransfer[i]);
                            Image img = Image.FromFile(fFull);
                            if (!img.PropertyIdList.Contains(exifOrientationID))
                            {
                                Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop + i, row);
                            }
                            else
                            {

                                RotImage(img);
                                Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop + i, row);
                            }
                        }
                        // Move highlighted + slected to top of that index
                        //dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Selected = true;
                        dataGridView1.CurrentCell = this.dataGridView1[1, rowIndexOfItemUnderMouseToDrop];
                        //capLength();
                        fileSize();
                        fileSizeTotal();
                        updatePictureBox();
                        BarExample();

                    }
                }
                else
                {
                    // For loading into an empty dataGridView1
                    if (dataGridView1.Rows.Count == 0)
                    {
                        for (int i = 0; i < fileNames.Length; i++)
                        {
                            string fFull = Path.GetFullPath(fileNames[i]);
                            string fileNam = Path.GetFileNameWithoutExtension(fileNames[i]);
                            Image img = Image.FromFile(fFull);
                            if (!img.PropertyIdList.Contains(exifOrientationID))
                            {
                                Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                dataGridView1.Rows.Add(row);
                            }
                            else
                            {
                                RotImage(img);
                                Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                dataGridView1.Rows.Add(row);
                            }
                            
                        }
                        //capLength();
                        fileSize();
                        fileSizeTotal();
                        updatePictureBox();
                        BarExample();

                    }
                    // For loading into an already populated dataGridView1 at empty row
                    else
                    {
                        if (rowIndexOfItemUnderMouseToDrop == -1)
                        {
                            int totalRows = dataGridView1.Rows.Count;
                            for (int i = 0; i < fileNames.Length; i++)
                            {
                                string fFull = Path.GetFullPath(fileNames[i]);
                                string fileNam = Path.GetFileNameWithoutExtension(fileNames[i]);
                                Image img = Image.FromFile(fFull);
                                if (!img.PropertyIdList.Contains(exifOrientationID))
                                {
                                    Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                    dataGridView1.Rows.Add(row);
                                }
                                else
                                {
                                    RotImage(img);
                                    Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                    dataGridView1.Rows.Add(row);
                                }

                            }
                            //capLength();
                            fileSize();
                            fileSizeTotal();
                            updatePictureBox();
                            BarExample();
                            //MessageBox.Show("copy to populated empty cell");

                        }
                        // For loading into an already populated dataGridView1 at occupied row
                        else
                        {
                            for (int i = 0; i < fileNames.Length; i++)
                            {
                                string fFull = Path.GetFullPath(fileNames[i]);
                                string fileNam = Path.GetFileNameWithoutExtension(fileNames[i]);
                                Image img = Image.FromFile(fFull);
                                if (!img.PropertyIdList.Contains(exifOrientationID))
                                {
                                    Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                    dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop + i, row);
                                }
                                else
                                {
                                    RotImage(img);
                                    Object[] row = new object[] { fileNam, img, "Insert caption here", fFull };
                                    dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop + i, row);
                                }
                                
                            }

                            //dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Selected = true;
                            dataGridView1.CurrentCell = this.dataGridView1[1, rowIndexOfItemUnderMouseToDrop];
                            //capLength();
                            fileSize();
                            fileSizeTotal();
                            updatePictureBox();
                            BarExample();
                        }
                    }
                }
            }
        }


        // Allows the left and right click to highlight a row in dataGridView1
        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Image img;
                try
                {
                    var hti = dataGridView1.HitTest(e.X, e.Y);
                    dataGridView1.ClearSelection();
                    dataGridView1.CurrentCell = dataGridView1.Rows[hti.RowIndex].Cells[hti.ColumnIndex];
                    dataGridView1.Rows[hti.RowIndex].Selected = true;
                    string txt = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    img = new Bitmap(txt);
                    updatePictureBox();
                }
                catch
                {
                    return;
                }
            }
        }


        private void KeyEvent(object sender, KeyEventArgs e) //Keyup Event 
        {
            if (e.KeyCode == Keys.Enter)
            {
                updatePictureBox();
            }
            if (e.KeyCode == Keys.Up)
            {
                updatePictureBox();
            }
            if (e.KeyCode == Keys.Down)
            {
                updatePictureBox();
            }

        }


        // BUTTON UP
        private void button3_Click_1(object sender, EventArgs e)
        {
            MoveUp();
        }


        // BUTTON DOWN
        private void button4_Click(object sender, EventArgs e)
        {
            MoveDown();
        }


        // Sorting function
        private static int DataGridViewRowIndexCompare(DataGridViewRow x, DataGridViewRow y)
        {
            if (x == null)
            {
                if (y == null)
                {
                    // If x is null and y is null, they're
                    // equal. 
                    return 0;
                }
                else
                {
                    // If x is null and y is not null, y
                    // is greater. 
                    return -1;
                }
            }
            else
            {
                // If x is not null...
                //
                if (y == null)
                // ...and y is null, x is greater.
                {
                    return 1;
                }
                else
                {
                    // ...and y is not null, compare the 
                    // lengths of the two strings.
                    //
                    int retval = x.Index.CompareTo(y.Index);

                    if (retval != 0)
                    {
                        // If the strings are not of equal length,
                        // the longer string is greater.
                        //
                        return retval;
                    }
                    else
                    {
                        // If the strings are of equal length,
                        // sort them with ordinary string comparison.
                        //
                        return x.Index.CompareTo(y.Index);
                    }
                }
            }
        }


        // BUTTON - delete
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && dataGridView1.SelectedRows.Count > 0)
            {
                int rowIndex = dataGridView1.SelectedCells[0].OwningRow.Index;

                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    if (rowIndex != -1 && dataGridView1.Rows.Count >= 1)
                    {
                        dataGridView1.Rows.Remove(row);
                        dataGridView1.ClearSelection();
                    }
                }
            }
            textBox3.Text = "";
            textBox4.Text = "";
            pictureBox1.Image = null;
            BarExample();
            fileSizeTotal();
        }


        // BUTTON TOP
        void top_Click(object sender, EventArgs e)
        {

            if (dataGridView1.Rows.Count == 0) return;

            if (dataGridView1.SelectedRows.Count == 0)
            {
                if (dataGridView1.CurrentCell.RowIndex > -1)
                {
                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Selected = true;
                }
                else
                {
                    return;
                }
            }

            List<DataGridViewRow> SelectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow dgvr in dataGridView1.SelectedRows)
            {
                SelectedRows.Add(dgvr);
            }
            SelectedRows.Sort(DataGridViewRowIndexCompare);

            for (int i = 0; i <= SelectedRows.Count - 1; i++)
            {
                int selRowIndex = SelectedRows[i].Index;
                if (selRowIndex > 0)
                {
                    dataGridView1.Rows.Remove(SelectedRows[i]);
                    //dataGridView1.Rows.Insert(selRowIndex - 1, SelectedRows[i]);
                    dataGridView1.Rows.Insert(0 + i, SelectedRows[i]);
                    dataGridView1.CurrentCell.Selected = false;
                    dataGridView1.Rows[selRowIndex - 1].Selected = true;
                }
                else
                {
                    // if selRowIndex == 0
                    return;
                }
            }
            // now make selected rows the top SelectedRows.Count
            dataGridView1.ClearSelection();
            for (int j = 0; j < SelectedRows.Count; j++)
            {
                dataGridView1.Rows[j].Selected = true;
            }
            scrollGrid();
            BarExample();
        }


        // BUTTON BOTTOM
        private void bottom_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0) return;

            if (dataGridView1.SelectedRows.Count == 0)
            {
                if (dataGridView1.CurrentCell.RowIndex > -1)
                {
                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Selected = true;
                }
                else
                {
                    return;
                }
            }

            List<DataGridViewRow> SelectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow dgvr in dataGridView1.SelectedRows)
            {
                SelectedRows.Add(dgvr);
            }

            SelectedRows.Sort(DataGridViewRowIndexCompare);

            for (int i = 0; i < SelectedRows.Count; i++)
            {
                int selRowIndex = SelectedRows[i].Index;

                if (selRowIndex < dataGridView1.Rows.Count)
                {
                    dataGridView1.Rows.Remove(SelectedRows[i]);
                    dataGridView1.Rows.Add(SelectedRows[i]);
                }
                else
                {
                    return;
                }
            }
            // now make selected rows the bottom SelectedRows.Count


            dataGridView1.ClearSelection();
            for (int j = dataGridView1.Rows.Count - SelectedRows.Count; j < dataGridView1.Rows.Count; j++)
            {
                dataGridView1.Rows[j].Selected = true;
            }
            scrollGrid();
            BarExample();
        }


        // Keeps the view centered
        private void scrollGrid()
        {

            int halfWay = (dataGridView1.DisplayedRowCount(false) / 2);
            if (dataGridView1.FirstDisplayedScrollingRowIndex + halfWay > dataGridView1.SelectedRows[0].Index ||
                (dataGridView1.FirstDisplayedScrollingRowIndex + dataGridView1.DisplayedRowCount(false) - halfWay) <= dataGridView1.SelectedRows[0].Index)
            {
                int targetRow = dataGridView1.SelectedRows[0].Index;

                targetRow = Math.Max(targetRow - halfWay, 0);
                dataGridView1.FirstDisplayedScrollingRowIndex = targetRow;

            }
        }


        // DragEnter needed??
        private void dataGridView1_DragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }


        // MENU - Drag & Drop Individual Images
        private void dataGridView1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }


        // MENU - Drag & Drop Individual Images
        private void Form1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }


        // Allows theleft and right click to highlight a row in dataGridView1
        private void OnRowHeaderMouseClick(object sender, MouseEventArgs e)
        {
            int rowIndex = dataGridView1.SelectedCells[0].OwningRow.Index;
            dataGridView1.Rows[rowIndex].Selected = true;
            updatePictureBox();
        }


        // Image cell click = View larger image
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            updatePictureBox();
        }


        // update pictureBox
        private void updatePictureBox()
        {

            fileSize();

            Bitmap resizedImage;

            String txt = dataGridView1.CurrentRow?.Cells[3].Value.ToString();

            Image img_file = Image.FromFile(txt);

            if (!img_file.PropertyIdList.Contains(exifOrientationID))
            {
                Image img0;
                using (var bmpTemp = new Bitmap(img_file))
                {
                    img0 = new Bitmap(bmpTemp);

                    int rectHeight = pictureBox1.Height;
                    int rectWidth = pictureBox1.Width;

                    //if the image is squared set it's height and width to the smallest of the desired dimensions (our box). In the current example rectHeight<rectWidth
                    if (img0.Height == img0.Width)
                    {
                        resizedImage = new Bitmap(img0, rectHeight, rectHeight);
                    }
                    else
                    {
                        //calculate aspect ratio
                        float aspect = img0.Width / (float)img0.Height;
                        int newWidth, newHeight;
                        //calculate new dimensions based on aspect ratio
                        newWidth = (int)(rectWidth * aspect);
                        newHeight = (int)(newWidth / aspect);
                        //if one of the two dimensions exceed the box dimensions
                        if (newWidth > rectWidth || newHeight > rectHeight)
                        {
                            //depending on which of the two exceeds the box dimensions set it as the box dimension and calculate the other one based on the aspect ratio
                            if (newWidth > newHeight)
                            {
                                newWidth = rectWidth;
                                newHeight = (int)(newWidth / aspect);
                            }
                            else
                            {
                                newHeight = rectHeight;
                                newWidth = (int)(newHeight * aspect);
                            }
                        }
                        resizedImage = new Bitmap(img0, newWidth, newHeight);
                        pictureBox1.Image = resizedImage;

                        pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                        textBox4.Text = txt;
                    }

                }
            }
            else
            {
                var rot = RotateFlipType.RotateNoneFlipNone;
                var prop = img_file.GetPropertyItem(exifOrientationID);
                int val = BitConverter.ToUInt16(prop.Value, 0);
                //var rot = RotateFlipType.RotateNoneFlipNone;

                if (val == 3 || val == 4)
                    rot = RotateFlipType.Rotate180FlipNone;
                else if (val == 5 || val == 6)
                    rot = RotateFlipType.Rotate90FlipNone;
                else if (val == 7 || val == 8)
                    rot = RotateFlipType.Rotate270FlipNone;

                if (val == 2 || val == 4 || val == 5 || val == 7)
                    rot |= RotateFlipType.RotateNoneFlipX;


                float filesize = new FileInfo(dataGridView1.SelectedRows[0].Cells[3].Value.ToString()).Length;

                if (txt != null)
                    if (filesize > 5000000)
                    {
                        pictureBox1.Image = null;
                        textBox4.Text = txt;
                        StringBuilder myStringBuilder = new StringBuilder("This image exceeds 5 MB: \n\n");
                        myStringBuilder.Append(txt + "\n\n");
                        myStringBuilder.Append("It may cause problems if the Photolog app tries to view it. \n\n" +
                            "It is also likely to cause problems if you try and PUBLISH it in your Word Document. " +
                        "Maybe you could try compressing the image or use a different one?");
                        MessageBox.Show(myStringBuilder.ToString());
                    }

                    else
                    {
                        Image img;
                        using (var bmpTemp = new Bitmap(txt))
                        {


                            img = new Bitmap(bmpTemp);

                            int rectHeight = pictureBox1.Height;
                            int rectWidth = pictureBox1.Width;

                            //if the image is squared set it's height and width to the smallest of the desired dimensions (our box). In the current example rectHeight<rectWidth
                            if (img.Height == img.Width)
                            {
                                resizedImage = new Bitmap(img, rectHeight, rectHeight);
                            }
                            else
                            {
                                //calculate aspect ratio
                                float aspect = img.Width / (float)img.Height;
                                int newWidth, newHeight;
                                //calculate new dimensions based on aspect ratio
                                newWidth = (int)(rectWidth * aspect);
                                newHeight = (int)(newWidth / aspect);
                                //if one of the two dimensions exceed the box dimensions
                                if (newWidth > rectWidth || newHeight > rectHeight)
                                {
                                    //depending on which of the two exceeds the box dimensions set it as the box dimension and calculate the other one based on the aspect ratio
                                    if (newWidth > newHeight)
                                    {
                                        newWidth = rectWidth;
                                        newHeight = (int)(newWidth / aspect);
                                    }
                                    else
                                    {
                                        newHeight = rectHeight;
                                        newWidth = (int)(newHeight * aspect);
                                    }
                                }
                                resizedImage = new Bitmap(img, newWidth, newHeight);
                                pictureBox1.Image = resizedImage;


                                if (rot != RotateFlipType.RotateNoneFlipNone)
                                    resizedImage.RotateFlip(rot);

                                pictureBox1.Image = resizedImage; 
                                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                                textBox4.Text = txt;
                            }

                        }
                    }
            }
            //capLength();
        }


        // METHOD - Calculate Image size
        private void fileSize()
        {
            float n, d;
            if (dataGridView1.SelectedRows.Count > 0) // make sure user select at least 1 row 
            {
                float filesize = new FileInfo(dataGridView1.SelectedRows[0].Cells[3].Value.ToString()).Length;
                //float filesize = (File.OpenRead(dataGridView1.SelectedRows[0].Cells[3].Value.ToString())).Length;
                n = filesize / 1048576;
                d = filesize % 1048576;
                textBox3.Text = n.ToString("n2");
                if (n > 2)
                {
                    textBox3.BackColor = Color.Red;
                }
                else if (n < 2)
                {
                    textBox3.BackColor = Color.White;
                }
            }
        }


        // METHOD - calculate Total Image file size
        private void fileSizeTotal()
        {
            float n, d;
            float total = 0;
            foreach (DataGridViewRow row in dataGridView1.Rows)

            {
                float filesize = new FileInfo(row.Cells[3].Value.ToString()).Length;
                //float filesize = (File.OpenRead(dataGridView1.SelectedRows[0].Cells[3].Value.ToString())).Length;
                n = filesize / 1048576;
                d = filesize % 1048576;
                total += n;
                textBox5.Text = total.ToString("n2");
            }

        }



        //// Caption Cell click = update caption length
        //private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    if (dataGridView1.CurrentCell.ColumnIndex.Equals(2) && e.RowIndex != -1)
        //    {
        //        if (dataGridView1.CurrentCell != null && dataGridView1.CurrentCell.Value != null)
        //            //capLength();
        //    }
        //}


        //// METHOD - calculate caption Length
        //private void capLength()
        //{
        //    if (dataGridView1.SelectedRows.Count > 0) // make sure user select at least 1 row 
        //    {
        //        string cap = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
        //        textBox2.Text = cap.Length.ToString();
        //    }
        //}



        // Create Datable of datagridViewView1
        private System.Data.DataTable GetDataTableFromDGV1(DataGridView dgv)
        {
            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2.Columns.Add("fileName", typeof(string));
            dt2.Columns.Add("dataGridView1Bitmap", typeof(string));
            dt2.Columns.Add("dataGridView1Caption", typeof(string));
            dt2.Columns.Add("dataGridView1Path", typeof(string));
            //dt2.Columns.Add("dataGridView1ImageNumber", typeof(string));

            object[] cellValues2 = new object[dgv.Columns.Count];
            foreach (DataGridViewRow row in dgv.Rows)
            {
                for (int ii = 0; ii < row.Cells.Count; ii++)
                {
                    cellValues2[ii] = row.Cells[ii].Value;
                }
                dt2.Rows.Add(cellValues2);
            }
            return dt2;
        }


        // BUTTON - CREATE WORD DOC
        private void button2_Click(object sender, EventArgs e)
        {
            //AutoClosingMessageBox.Show("Creating Your Word Document", "In Progress...", 5000);
            //IsOn = !IsOn;
            CreateWordDoc(dataGridView1);

        }



        // METHOD - Create the Word doc
        private void CreateWordDoc(DataGridView DGV)
        {
            if (DGV.Rows.Count != 0)
            {
                //Create a missing variable for missing value
                object oMissing = Missing.Value;

                // \endofdoc is a predefined bookmark
                object oEndOfDoc = "\\endofdoc";

                //Start Word and create a new document.          
                _Application oWord = new Word.Application();
                oWord.Visible = true;
                _Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);

                int RowCount = DGV.Rows.Count;

                // Iterate over each DataGrid row and extract image path and caption text
                for (int i = 0; i <= RowCount - 1; i++)
                {

                    // make a para as numbered list
                    Paragraph oPara0 = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara0.KeepWithNext = 0;
                    oPara0.Format.SpaceAfter = 0;
                    Paragraph oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
                    Range rngTarget0 = oPara0.Range;
                    Range rngTarget1 = oPara1.Range;
                    rngTarget0.Font.Size = 12;
                    rngTarget0.Font.Name = "Tahoma";
                    rngTarget1.Font.Size = 12;
                    rngTarget1.Font.Name = "Tahoma";
                    object anchor = rngTarget1;

                    //int paraStartNumber = 4;


                    rngTarget0.Paragraphs.TabStops.Add(42, WdTabAlignment.wdAlignTabRight);
                    rngTarget0.ListFormat.ApplyNumberDefault();
                    

                    // Get image path and caption from dataGridView
                    string fileName1 = DGV.Rows[i].Cells[3].Value.ToString();
                    string caption = "" + DGV.Rows[i].Cells[2].Value.ToString();
                    caption = caption.TrimEnd('\r', '\n');

                    // Picture placement
                    //InlineShape pic = rngTarget1.InlineShapes.AddPicture(fileName1, ref oMissing, ref oMissing, ref anchor);
                    InlineShape pic = rngTarget1.InlineShapes.AddPicture(fileName1, ref oMissing, ref oMissing, ref oMissing);
                    Shape sh = pic.ConvertToShape();
                    sh.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;


                    //int exifOrientationID = 0x112; //274


                    Image img = Image.FromFile(fileName1);
                    if (!img.PropertyIdList.Contains(exifOrientationID))
                    {
                        // Scale the height here for pics with no EXIF
                        sh.Height = 252;

                        if (sh.Width > 400)
                        {
                            sh.Width = 400;
                        }
                    }
                    else
                    {
                        var prop = img.GetPropertyItem(exifOrientationID);
                        int val = BitConverter.ToUInt16(prop.Value, 0);
                        //MessageBox.Show(val.ToString());

                        if (val == 5 || val == 6 || val == 7 || val == 8)
                        {
                            sh.Width = 252;

                            if (sh.Height > 400)
                            {
                                sh.Height = 400;
                            }
                        }
                        else
                        {
                            sh.Height = 252;

                            if (sh.Width > 400)
                            {
                                sh.Width = 400;
                            }
                        }
                    }
                    
                    sh.Left = (float)WdShapePosition.wdShapeCenter;
                    sh.Top = (float)WdShapePosition.wdShapeTop;
                    //sh.Top = 0;

                    //Write substring into Word doc with a bullet before it.
                    rngTarget0.InsertBefore(caption + "\v");
                    oPara1.Format.SpaceAfter = 264;
                    //rngTarget1.InsertParagraphAfter();                
                }

            }
        }


        // Make the Chart
        public void BarExample()
        {
            //float n;
            List<float> points = new List<float>();
            foreach (DataGridViewRow dgvr in dataGridView1.Rows)
            {
                float sz = new FileInfo(dgvr.Cells[3].Value.ToString()).Length;
                points.Add(sz / 1048576);
                //n = filesize / 1048576;
            }
            float[] pointsArray = points.ToArray();

            chart1.Series.Clear();
            chart1.Titles.Clear();
            chart1.Titles.Add("Image Sizes (MB)");
            chart1.Palette = ChartColorPalette.SeaGreen;
            chart1.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.AlignDataPointsByAxisLabel();


            // Add series.
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                System.Windows.Forms.DataVisualization.Charting.Series series = this.chart1.Series.Add(dataGridView1.Rows[i].Cells[3].Value.ToString());
                series.Points.AddXY(i + 1, pointsArray[i]);
                if (pointsArray[i] >= 2)
                {
                    chart1.Series[i].Color = Color.Red;
                    chart1.Series[i].Label = (i + 1).ToString();
                }
            }
        }

        // RESUME FUNCTION
        private void resume(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            //ofd.Filter = "XML Files (*.xml)|*.xml";
            ofd.Filter = "photolog files (*.photolog; *.XML)| *.photolog; *.XML";
            ofd.FilterIndex = 1;

            // check user selects pass
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xmlFileName = ofd.FileName;
                // For updating textBox
                string xmlFile = Path.GetFileName(xmlFileName);
                //label7.Text = xmlFile;
                if (sender.ToString() == "Resume")
                {
                    dataGridView1.Rows.Clear();
                }
                else
                {
                    // empty braces == Python pass                
                }

                var array0 = dataGridView1.Rows.Cast<DataGridViewRow>()
                                             .Select(x => x.Cells[3].Value.ToString().Trim()).ToArray();
                //string[] collisions = { };
                List<string> collisions = new List<string>();


                // Read and load info from XML 
                XDocument doc = XDocument.Load(xmlFileName);
                if (doc.Root.Elements().Any())
                {
                    try
                    {
                        foreach (var dm2 in doc.Descendants("Table1"))
                        {
                            string fileName = dm2.Element("fileName").Value;
                            string fileNameFull = dm2.Element("dataGridView1Path").Value.ToString();

                            bool contains = array0.Contains(fileNameFull, StringComparer.OrdinalIgnoreCase);
                            if (contains == true)
                            {
                                // If collisions exist
                                collisions.Add(fileNameFull);
                            }
                            else
                            {
                                Image img = Image.FromFile(dm2.Element("dataGridView1Path").Value.ToString());
                                var capt = dm2.Element("dataGridView1Caption")?.Value;
                                var pth = dm2.Element("dataGridView1Path").Value;
                                //const int exifOrientationID = 0x112; //274


                                if (!img.PropertyIdList.Contains(exifOrientationID))
                                {
                                    dataGridView1.Rows.Add(fileName, img, capt, pth);
                                }
                                else
                                {
                                    RotImage(img);
                                    dataGridView1.Rows.Add(fileName, img, capt, pth);
                                }
                            }
                        }
                        if (collisions.Count > 0)
                        {
                            StringBuilder myStringBuilder = new StringBuilder("Whilst trying to load images from the following project: \n\n");
                            myStringBuilder.Append(xmlFileName + "\n\n");
                            myStringBuilder.Append("You tried to add the following file(s): \n\n");
                            foreach (string value in collisions)
                            {
                                myStringBuilder.Append(value);
                                myStringBuilder.Append(", ");
                            }

                            myStringBuilder.Append("\n\nBut they were already in the existing list and were NOT loaded. \n\n");
                            MessageBox.Show(myStringBuilder.ToString());
 
                            //capLength();
                            fileSize();
                            fileSizeTotal();
                            updatePictureBox();
                            textBox6.Text = "";
                            BarExample();
                        }
                        else
                        {
                            //capLength();
                            fileSize();
                            fileSizeTotal();
                            updatePictureBox();
                            BarExample();

                            if (sender.ToString() == "Resume")
                            {
                                textBox6.Text = xmlFileName;
                            }
                            else
                            {
                                textBox6.Text = "";
                            }
                        }
                    }
                    // Error
                    catch (Exception exc)
                    {
                        StringBuilder myStringBuilder = new StringBuilder(xmlFileName + "\n");
                        myStringBuilder.Append("tried to load the following file: \n\n");
                        myStringBuilder.Append(exc.Message + "\n\n");
                        myStringBuilder.Append("But this file does not exist in this folder location. Have you perhaps moved it? Or has it been renamed? \n\n");
                        myStringBuilder.Append("If you have moved all your files you could try changing the parent folder from the menu! \n\n");
                        MessageBox.Show(myStringBuilder.ToString());
                    }
                }
                else
                {
                    StringBuilder myStringBuilder = new StringBuilder("You tried to load images from the following saved project: \n\n");
                    myStringBuilder.Append(xmlFileName + "\n\n");
                    myStringBuilder.Append("But this file is empty. \n\n");
                    MessageBox.Show(myStringBuilder.ToString());
                }
            }
        }


        // ROTATE IMAGE - METHOD
        private void RotImage(Image img)
        {
            var prop = img.GetPropertyItem(exifOrientationID);
            int val = BitConverter.ToUInt16(prop.Value, 0);
            var rot = RotateFlipType.RotateNoneFlipNone;

            if (val == 3 || val == 4)
                rot = RotateFlipType.Rotate180FlipNone;
            else if (val == 5 || val == 6)
                rot = RotateFlipType.Rotate90FlipNone;
            else if (val == 7 || val == 8)
                rot = RotateFlipType.Rotate270FlipNone;

            if (val == 2 || val == 4 || val == 5 || val == 7)
                rot |= RotateFlipType.RotateNoneFlipX;

            if (rot != RotateFlipType.RotateNoneFlipNone)
                img.RotateFlip(rot);
        }


        // MOVE UP FUNCTION
        private void MoveUp()
        {

            if (dataGridView1.Rows.Count == 0) return;

            if (dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }

            List<DataGridViewRow> SelectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow dgvr in dataGridView1.SelectedRows)
            {
                SelectedRows.Add(dgvr);
            }
            SelectedRows.Sort(DataGridViewRowIndexCompare);

            for (int i = 0; i <= SelectedRows.Count - 1; i++)
            {
                int selRowIndex = SelectedRows[i].Index;

                if (selRowIndex > 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[selRowIndex - 1].Cells[1];
                    dataGridView1.Rows.Remove(SelectedRows[i]);
                    dataGridView1.Rows.Insert(selRowIndex - 1, SelectedRows[i]);
                    dataGridView1.CurrentCell.Selected = false;
                    dataGridView1.Rows[selRowIndex - 1].Selected = true;
                }
                else
                {            
                    return;

                }
            }
            scrollGrid();
            BarExample();
        }


        // MOVE DOWN FUNCTION
        private void MoveDown()
        {
            if (dataGridView1.Rows.Count == 0) return;

            if (dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }

            List<DataGridViewRow> SelectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow dgvr in dataGridView1.SelectedRows)
            {
                SelectedRows.Add(dgvr);
            }

            SelectedRows.Sort(DataGridViewRowIndexCompare);

            for (int i = SelectedRows.Count - 1; i >= 0; i--)
            {
                int selRowIndex = SelectedRows[i].Index;
                if (selRowIndex < dataGridView1.Rows.Count - 1)
                {
                    dataGridView1.Rows.Remove(SelectedRows[i]);
                    dataGridView1.Rows[selRowIndex].Selected = false;
                    dataGridView1.Rows.Insert(selRowIndex + 1, SelectedRows[i]);
                    dataGridView1.Rows[selRowIndex + 1].Selected = true;
                    //dataGridView1.CurrentCell = dataGridView1.Rows[selRowIndex + 1].Cells[1];
                }
                else
                {
                    return;
                }
            }
            scrollGrid();
            BarExample();
        }

        private void standardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[2].DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 10F);
        }

        private void largeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[2].DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 14F);
        }

    }
}





/*


     private void toSpell()
        {
            if (dataGridView1.SelectedRows.Count > 0) // make sure user select at least 1 row 
            {
                string cap = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                spellBox1.Text = cap;
            }
        }


        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            spellBox1.Text = textBox1.Text;
        }

        Control cnt;

        void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.TextChanged += new EventHandler(tb_TextChanged);
            cnt = e.Control;
            cnt.TextChanged += tb_TextChanged;
        }

        void tb_TextChanged(object sender, EventArgs e)
        {
            if (cnt.Text != string.Empty)
            {
                spellBox1.Text = cnt.Text;
               
            }
        }

        // METHOD - Vary image quality
        private void VaryQualityLevel()
        {
            // Get a bitmap. The using statement ensures objects  
            // are automatically disposed from memory after use.  
            using (Bitmap bmp1 = new Bitmap(@"C:\TestPhoto.jpg"))
            {
                ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);

                // Create an Encoder object based on the GUID  
                // for the Quality parameter category.  
                System.Drawing.Imaging.Encoder myEncoder =
                    System.Drawing.Imaging.Encoder.Quality;

                // Create an EncoderParameters object.  
                // An EncoderParameters object has an array of EncoderParameter  
                // objects. In this case, there is only one  
                // EncoderParameter object in the array.  
                EncoderParameters myEncoderParameters = new EncoderParameters(1);

                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 50L);
                myEncoderParameters.Param[0] = myEncoderParameter;
                bmp1.Save(@"C:\Users\dhaggerty\Desktop\TestPhotoQualityFifty.jpg", jpgEncoder, myEncoderParameters);

                myEncoderParameter = new EncoderParameter(myEncoder, 100L);
                myEncoderParameters.Param[0] = myEncoderParameter;
                bmp1.Save(@"C:\Users\dhaggerty\Desktop\TestPhotoQualityHundred.jpg", jpgEncoder, myEncoderParameters);

                // Save the bitmap as a JPG file with zero quality level compression.  
                myEncoderParameter = new EncoderParameter(myEncoder, 0L);
                myEncoderParameters.Param[0] = myEncoderParameter;
                bmp1.Save(@"C:\Users\dhaggerty\Desktop\TestPhotoQualityZero.jpg", jpgEncoder, myEncoderParameters);
            }
        }


        // METHOD Image quality Encoder
        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }


        // BUTTON - Vary image quality
        private void button1_Click(object sender, EventArgs e)
        {
            VaryQualityLevel();
        }


    */
