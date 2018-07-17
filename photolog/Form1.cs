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

        public Form1()
        {
            InitializeComponent();

            //Create right click menu using contextmenustrip for right click move top bottom
            ContextMenuStrip s = new ContextMenuStrip();

            // add one right click menu item named as top and bottom          
            ToolStripMenuItem top = new ToolStripMenuItem();
            ToolStripMenuItem bottom = new ToolStripMenuItem();
            top.Text = "Send to TOP";
            bottom.Text = "Send to BOTTOM";

            // add the clickevent of hello item
            top.Click += top_Click;
            bottom.Click += bottom_Click;

            // add the item in right click menu
            s.Items.Add(top);
            s.Items.Add(bottom);

            // attach the right click menu with form
            this.ContextMenuStrip = s;

            //this.dataGridView1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseDown);
            dataGridView1.RowHeaderMouseClick += new DataGridViewCellMouseEventHandler(OnRowHeaderMouseClick);

            this.AllowDrop = true;
            //this.DragOver += new DragEventHandler(Form1_DragOver);
            //this.DragDrop += new DragEventHandler(Form1_DragDrop);
            this.dataGridView1.DragOver += new DragEventHandler(dataGridView1_DragOver);
            this.dataGridView1.DragDrop += new DragEventHandler(dataGridView1_DragDrop);
            this.dataGridView1.DragEnter += new DragEventHandler(dataGridView1_DragEnter);
            this.dataGridView1.AllowDrop = true;
            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            // declare at form level

        }


        // Set Form listView and datGridView properties on load
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Load += new EventHandler(Form1_Load);

            // form resizing
            // Turn off form resizing and maximize
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;

            // dataGridView1
            dataGridView1.RowTemplate.Height = 64;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.Columns["Caption"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft;
            dataGridView1.AllowDrop = true;
            dataGridView1.MultiSelect = true;

            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            dataGridView1.Columns[1].DefaultCellStyle.Font = new System.Drawing.Font("Verdana", 10F);
            dataGridView1.Columns[2].DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 10F);

            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Blue;
            dataGridView1.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);


            // pictureBox1
            pictureBox1.SizeMode = PictureBoxSizeMode.CenterImage;

            ToolTip toolTip1 = new ToolTip();
            toolTip1.SetToolTip(button2, "Make a Word document");
            toolTip1.SetToolTip(button1, "Rotation is only saved in your photolog project and NOT to your computer's file system");


            // Chart
            chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisX.ScrollBar.Enabled = true;
           

            // photolog version
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            label5.Text = "PhotoLog Desktop Version: " + version;
        }


        /*
         MENU OPTIONS
         */


        // MENU - Save as
        private void saveProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataSet dS = new DataSet();
            //System.Data.DataTable dT1 = GetDataTableFromDGV0(dataGridView0);
            System.Data.DataTable dT2 = GetDataTableFromDGV1(dataGridView1);

            //dS.Tables.Add(dT1);
            dS.Tables.Add(dT2);


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML files(.xml)|*.xml|all Files(*.*)|*.*";
            saveFileDialog.AddExtension = true;
            saveFileDialog.Title = "Save work as .XML file";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                dS.WriteXml(File.Open(saveFileDialog.FileName, FileMode.Create));
                textBox6.Text = saveFileDialog.FileName;
                //label7.Text = Path.GetFileName(saveFileDialog.FileName);
            }
        }


        // MENU - Save
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {

            string path = textBox6.Text;

            if (path == "")
            {
                Console.WriteLine("empty");
                DataSet dS = new DataSet();
                //System.Data.DataTable dT1 = GetDataTableFromDGV0(dataGridView0);
                System.Data.DataTable dT2 = GetDataTableFromDGV1(dataGridView1);

                //dS.Tables.Add(dT1);
                dS.Tables.Add(dT2);


                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XML files(.xml)|*.xml|all Files(*.*)|*.*";
                saveFileDialog.AddExtension = true;
                saveFileDialog.Title = "Save work as .XML file";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    dS.WriteXml(File.Open(saveFileDialog.FileName, FileMode.Create));
                    textBox6.Text = saveFileDialog.FileName;

                }
            }
            else
            {
                DataSet dS = new DataSet();
                System.Data.DataTable dT2 = GetDataTableFromDGV1(dataGridView1);
                dS.Tables.Add(dT2);
                dS.WriteXml(File.Open(path, FileMode.Create));

                AutoClosingMessageBox.Show(textBox6.Text, "Saved", 3000);
            }
        }


        // MENU - Resume
        private void resumeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "XML Files (*.xml)|*.xml";
            ofd.FilterIndex = 1;

            // check user selects pass
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xmlFileName = ofd.FileName;
                // For updating textBox
                string xmlFile = Path.GetFileName(xmlFileName);
                //label7.Text = xmlFile;

                dataGridView1.Rows.Clear();

                // Read and load info from XML 
                XDocument doc = XDocument.Load(xmlFileName);
                try
                {
                    foreach (var dm2 in doc.Descendants("Table1"))
                    {
                        string fileName = dm2.Element("fileName").Value;
                        string fileNameFull = dm2.Element("dataGridView1Path").Value.ToString();
                        Image img = Image.FromFile(dm2.Element("dataGridView1Path").Value.ToString());
                        var capt = dm2.Element("dataGridView1Caption").Value;
                        var pth = dm2.Element("dataGridView1Path").Value;
                        dataGridView1.Rows.Add(fileName, img, capt, pth);
                    }
                    dgLength();
                    dgLength();
                    capLength();
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
        }


        // MENU - Change My parent folder


        // Overlay file name on top of image      
        private void dataGridView1_CellPainting(object sender,
                                DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;                  // no image in the header
            if (e.ColumnIndex == this.dataGridView1.Columns["imageColumn"].Index)

            {
                e.PaintBackground(e.ClipBounds, false);  // no highlighting
                e.PaintContent(e.ClipBounds);

                // calculate the location of your text..:
                int y = e.CellBounds.Bottom - 17;         // your  font height

                string mystring = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                var result = mystring.Substring(mystring.Length - Math.Min(4, mystring.Length));

                System.Drawing.Rectangle rect = new System.Drawing.Rectangle(e.CellBounds.Location.X + 1, e.CellBounds.Location.Y + 49, 35, 14);
                e.Graphics.FillRectangle(Brushes.White, rect);

                e.Graphics.DrawString(result, e.CellStyle.Font,
                Brushes.Crimson, e.CellBounds.Left, y);

                //e.PaintContent(rect);
                e.Handled = true;                        // done with the image column 
            }
        }


        // Overlay file name on top of image      
        private void pictureBox1_CellPainting(object sender,
                                DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;                  // no image in the header
            if (e.ColumnIndex == this.dataGridView1.Columns["imageColumn"].Index)

            {
                e.PaintBackground(e.ClipBounds, false);  // no highlighting
                e.PaintContent(e.ClipBounds);

                // calculate the location of your text..:
                int y = e.CellBounds.Bottom - 17;         // your  font height

                string mystring = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                var result = mystring.Substring(mystring.Length - Math.Min(4, mystring.Length));

                System.Drawing.Rectangle rect = new System.Drawing.Rectangle(e.CellBounds.Location.X + 1, e.CellBounds.Location.Y + 49, 35, 14);
                e.Graphics.FillRectangle(Brushes.White, rect);

                e.Graphics.DrawString(result, e.CellStyle.Font,
                Brushes.Crimson, e.CellBounds.Left, y);

                //e.PaintContent(rect);
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

            //Console.WriteLine(rowIndexOfItemUnderMouseToDrop);

            // Make an array of all files being dragged in
            string[] fileNames = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (fileNames != null && fileNames.Length != 0)
            {
                // Make an array of all files in the dataGridView
                var array0 = dataGridView1.Rows.Cast<DataGridViewRow>()
                                             .Select(x => x.Cells[3].Value.ToString().Trim()).ToArray();

                // Which files "collide"
                var intersection = array0.Intersect(fileNames);

                // If collisions exist
                if (intersection.Count() > 0)
                {
                    StringBuilder sb = new StringBuilder("The following file(s) are already in the list: \n");

                    foreach (string id in intersection)
                    {
                        sb.Append(id + "\n");
                    }
                    MessageBox.Show(sb.ToString());
                }
                else
                // Otherwise load the files
                {
                    // For loading into an empty dataGridView1
                    if (dataGridView1.Rows.Count == 0)
                    {
                        for (int i = 0; i < fileNames.Length; i++)
                        {
                            string fFull = Path.GetFullPath(fileNames[i]);
                            string fileNam = Path.GetFileNameWithoutExtension(fileNames[i]);
                            Bitmap bmp1 = new Bitmap(fFull);
                            Object[] row = new object[] { fileNam, bmp1, "Insert caption here", fFull };
                            dataGridView1.Rows.Add(row);
                        }
                    }
                    // For loading into an already populated dataGridView1
                    else
                    {
                        if (rowIndexOfItemUnderMouseToDrop == -1)
                        {
                            int totalRows = dataGridView1.Rows.Count;
                            for (int i = 0; i < fileNames.Length; i++)
                            {
                                string fFull = Path.GetFullPath(fileNames[i]);
                                string fileNam = Path.GetFileNameWithoutExtension(fileNames[i]);
                                Bitmap bmp1 = new Bitmap(fFull);
                                Object[] row = new object[] { fileNam, bmp1, "Insert caption here", fFull };
                                // Add at index under mouse for first and + i for rest
                                dataGridView1.Rows.Insert(totalRows + i, row);
                            }
                            // Move highlighted + slected to top of that index
                            //dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Selected = true;
                            dataGridView1.CurrentCell = this.dataGridView1[1, totalRows];
                        }
                        else
                        {
                            for (int i = 0; i < fileNames.Length; i++)
                            {
                                string fFull = Path.GetFullPath(fileNames[i]);
                                string fileNam = Path.GetFileNameWithoutExtension(fileNames[i]);
                                Bitmap bmp1 = new Bitmap(fFull);
                                Object[] row = new object[] { fileNam, bmp1, "Insert caption here", fFull };
                                // Add at index under mouse for first and + i for rest
                                dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop + i, row);
                            }
                            // Move highlighted + slected to top of that index
                            //dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Selected = true;
                            dataGridView1.CurrentCell = this.dataGridView1[1, rowIndexOfItemUnderMouseToDrop];
                        }

                    }
                }
                dgLength();
                capLength();
                fileSize();
                fileSizeTotal();
                updatePictureBox();
                BarExample();
            }
        }


        // Allows theleft and right click to highlight a row in dataGridView1
        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                try
                {
                    var hti = dataGridView1.HitTest(e.X, e.Y);
                    dataGridView1.ClearSelection();
                    dataGridView1.CurrentCell = dataGridView1.Rows[hti.RowIndex].Cells[hti.ColumnIndex];
                    dataGridView1.Rows[hti.RowIndex].Selected = true;
                    updatePictureBox();
                }
                catch
                {
                    return;
                }
            }
        }


        // BUTTON UP
        private void button3_Click_1(object sender, EventArgs e)
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
            //foreach (int value in SelectedRows.Index)
            //{
            //    Console.WriteLine(value);
            //}

            for (int i = 0; i <= SelectedRows.Count - 1; i++)
            {
                //Console.WriteLine(SelectedRows.Count);
                int selRowIndex = SelectedRows[i].Index;
                //Console.WriteLine(selRowIndex);
                if (selRowIndex > 0)
                {
                    dataGridView1.Rows.Remove(SelectedRows[i]);
                    dataGridView1.Rows.Insert(selRowIndex - 1, SelectedRows[i]);
                    dataGridView1.CurrentCell.Selected = false;
                    dataGridView1.Rows[selRowIndex - 1].Selected = true;
                }
                else
                {
                    // if selRowIndex == 0
                    return;
                }
            }
            /*
            if (dataGridView1.SelectedRows.Count == 1)
            {
                dataGridView1.CurrentCell = dataGridView1.Rows[SelectedRows[0].Index].Cells[0];
            }*/
            scrollGrid();
            BarExample();
        }


        // BUTTON DOWN
        private void button4_Click(object sender, EventArgs e)
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

            for (int i = SelectedRows.Count - 1; i >= 0; i--)
            {
                int selRowIndex = SelectedRows[i].Index;
                //if ((selRowIndex <= dataGridView1.Rows.Count - 1) && (!(selRowIndex == dataGridView1.Rows.Count - 1)))
                if (selRowIndex < dataGridView1.Rows.Count - 1)
                {
                    dataGridView1.Rows.Remove(SelectedRows[i]);
                    dataGridView1.Rows[selRowIndex].Selected = false;
                    dataGridView1.Rows.Insert(selRowIndex + 1, SelectedRows[i]);
                    dataGridView1.Rows[selRowIndex + 1].Selected = true;
                }
                else
                {
                    return;
                }
            }
            scrollGrid();
            BarExample();
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
            int cnt = 0;
            int rowIndex = dataGridView1.SelectedCells[0].OwningRow.Index;
            int totalRows = dataGridView1.Rows.Count;
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(row);
                dataGridView1.ClearSelection();
                cnt += 1;
            }
            dgLength();
            fileSizeTotal();
            textBox3.Text = "";
            pictureBox1.Image = null;
            BarExample();
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
            //foreach (int value in SelectedRows.Index)
            //{
            //    Console.WriteLine(value);
            //}

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

            String txt = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            float filesize = new FileInfo(dataGridView1.SelectedRows[0].Cells[3].Value.ToString()).Length;


            if (txt != null)
                if (filesize > 5000000)
                {

                    pictureBox1.Image = null;

                    StringBuilder myStringBuilder = new StringBuilder("This image exceeds 5 MB: \n\n");
                    myStringBuilder.Append(txt + "\n\n");
                    myStringBuilder.Append("It may cause problems if the Photolog app tries to view it. \n\n" +
                        "It is also likely to cause problems if you try and PUBLISH it in your Word Document. " +
                    "Maybe you could try compressing the image or use a different one?");
                    MessageBox.Show(myStringBuilder.ToString());

                    /*
                MessageBox.Show("This image exceeds 5 MB and may cause problems if the app tries to view it.\n\n" +
                "It is also likely to cause problems if you try and PUBLISH it in your Word Document. " +
                "Maybe you could try compressing the image or use a different one?");
                */
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
                            textBox4.Text = txt;
                        }
                        // Use default image viewer
                        //Process.Start(txt);
                    }
                }
            else
            {
                MessageBox.Show("No Item is selected");
            }
            capLength();
        }



        // BUTTON - Rotate Image
        private void button1_Click_2(object sender, EventArgs e)
        {
            RotateImage();

            // store rotation value in dgv1

            // retrieve that rotation value for pictureBox

            // update thumbnail rotation
        }



        // METHOD - Rotate Image
        public Image RotateImage()
        {
            //var bmp = new Bitmap(textBox4.Text);
            var bmp = pictureBox1.Image;

            bmp.RotateFlip(RotateFlipType.Rotate270FlipNone);
            return pictureBox1.Image = bmp;
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
                //textBox3.Text = filesize.ToString();
                textBox3.Text = n.ToString("n2");
                if (n > 2)
                {
                    textBox3.BackColor = System.Drawing.Color.Red;
                }
                else if (n < 2)
                {
                    textBox3.BackColor = System.Drawing.Color.White;
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



        // Caption Cell click = update caption length
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex.Equals(2) && e.RowIndex != -1)
            {
                if (dataGridView1.CurrentCell != null && dataGridView1.CurrentCell.Value != null)
                    capLength();
            }
        }


        // METHOD - calculate dataGridView1 Length
        private void dgLength()
        {
            int dgRows = dataGridView1.Rows.Count;
            //textBox1.Text = dgRows.ToString();
        }


        // METHOD - calculate caption Length
        private void capLength()
        {
            if (dataGridView1.SelectedRows.Count > 0) // make sure user select at least 1 row 
            {
                string cap = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                textBox2.Text = cap.Length.ToString();
            }
        }



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

        private bool _IsOn;
        public bool IsOn
        {
            get
            {
                return _IsOn;
            }
            set
            {
                _IsOn = value;
                button2.Text = _IsOn ? "On" : "Off";
            }
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

                    rngTarget0.ListFormat.ApplyNumberDefault();


                    // Get image path and caption from dataGridView
                    string fileName1 = DGV.Rows[i].Cells[3].Value.ToString();
                    string caption = DGV.Rows[i].Cells[2].Value.ToString();


                    // Picture placement
                    //InlineShape pic = rngTarget1.InlineShapes.AddPicture(fileName1, ref oMissing, ref oMissing, ref anchor);
                    InlineShape pic = rngTarget1.InlineShapes.AddPicture(fileName1, ref oMissing, ref oMissing, ref oMissing);

                    Shape sh = pic.ConvertToShape();
                    sh.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;


                    // Windows runs as default at 96dpi (display) Macs run as default at 72 dpi (display)
                    // Assuming 72 points per inch
                    // 3.5 inches is 3.5*72 = 252
                    // 3.25 inches is 3.25*72 = 234

                    sh.Height = 252;

                    if (sh.Width > 400)
                    {
                        sh.Width = 400;
                    }

                    sh.Left = (float)WdShapePosition.wdShapeCenter;
                    sh.Top = (float)WdShapePosition.wdShapeTop;
                    //sh.Top = 0;

                    //Write substring into Word doc with a bullet before it.
                    rngTarget0.InsertBefore(caption + "\v");
                    oPara1.Format.SpaceAfter = 264;
                    //rngTarget1.InsertParagraphAfter();                
                }


                /*
                foreach (Shape ilPicture in oDoc.Shapes)
                {
                    //ilPicture.{compress the picture}
                }
                */

                /*
                foreach (Shape ilPicture in oDoc.Shapes)
                {
                    //ilPicture.{ROTATE}
                }
                */
            }
        }




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


            //foreach (var item in pointsArray)
            //{            
            //    Console.WriteLine(item / 1048576);
            //}
            
            chart1.Series.Clear();
            chart1.Titles.Clear();
            chart1.Titles.Add("Image Sizes (MB)");
            chart1.Palette = ChartColorPalette.Grayscale;
            chart1.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.AlignDataPointsByAxisLabel();
            

            // Add series.
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                System.Windows.Forms.DataVisualization.Charting.Series series = this.chart1.Series.Add(dataGridView1.Rows[i].Cells[3].Value.ToString());
                series.Points.AddXY(i+1, pointsArray[i]);
                if (pointsArray[i] >= 2)
                {
                    chart1.Series[i].Color = Color.Red;
                    chart1.Series[i].Label = (i+1).ToString();
                }

            }
            
        }
    }
}





/*
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


                // Allows theleft and right click to highlight a row in dataGridView1
        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                try
                {
                    var hti = dataGridView1.HitTest(e.X, e.Y);
                    dataGridView1.ClearSelection();
                    dataGridView1.CurrentCell = dataGridView1.Rows[hti.RowIndex].Cells[hti.ColumnIndex];
                    dataGridView1.Rows[hti.RowIndex].Selected = true;
                    updatePictureBox();
                }
                catch
                {
                    return;
                }
            }
        }


    */
