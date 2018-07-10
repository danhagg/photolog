using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Data;
using System.Xml.Linq;
using System.Drawing.Imaging;
using System.Text;
using System.Linq;

namespace photolog
{

    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();

            this.AllowDrop = true;
            //this.DragOver += new DragEventHandler(Form1_DragOver);
            //this.DragDrop += new DragEventHandler(Form1_DragDrop);
            this.dataGridView1.DragOver += new DragEventHandler(dataGridView1_DragOver);
            this.dataGridView1.DragDrop += new DragEventHandler(dataGridView1_DragDrop);
            this.dataGridView1.DragEnter += new DragEventHandler(dataGridView1_DragEnter);

            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);

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
            //dataGridView1.MultiSelect = true;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[1].DefaultCellStyle.Font = new System.Drawing.Font("Verdana", 12F, FontStyle.Bold);
            dataGridView1.Columns[2].DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 10F);
            //dataGridView1.Columns[2].SpellCheck.IsEnabled = true;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Blue;
            dataGridView1.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);


            // pictureBox1
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;


            // photolog version
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Console.WriteLine(version);
            label5.Text = "Version: " + version;
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
            saveFileDialog.Title = "Save work as .XML file";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                dS.WriteXml(File.Open(saveFileDialog.FileName, FileMode.Create));
            }
        }


        // MENU - Save



        // MENU - Resume
        private void resumeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog fbd = new FolderBrowserDialog();
            //fbd.Description = "Choose an your photolog .XML file";

            OpenFileDialog ofd = new OpenFileDialog();


            ofd.Filter = "XML Files (*.xml)|*.xml";
            ofd.FilterIndex = 1;
            //ofd.Multiselect = false;

            // check user selects pass
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xmlFileName = ofd.FileName;

                Console.WriteLine(xmlFileName);

                XDocument doc = XDocument.Load(xmlFileName);
                //Console.WriteLine(doc);
                /*
                foreach (var dm1 in doc.Descendants("Table1"))
                {
                    string fileName = dm1.Element("fileName").Value; ;
                    Image img = Image.FromFile(dm1.Element("dataGridView0Path").Value.ToString());
                    var pth = dm1.Element("dataGridView0Path").Value;
                    dataGridView0.Rows.Add(fileName, img, "", pth);
                }
                */
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
                }
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
            dgLength();
            capLength();
            fileSize();
            fileSizeTotal();
        }


        // MENU - Change My parent folder



        // Paint the 
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
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
            
        
        // Image cell click = View larger image
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            String txt = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            Console.WriteLine(txt);
            if (txt != null)
            {             
                Image img;
                using (var bmpTemp = new Bitmap(txt))
                {
                    img = new Bitmap(bmpTemp);
                    pictureBox1.Image = img;
                    textBox4.Text = txt;
                }
                //Image newImage = Image.FromFile(txt);
                //Process.Start(txt);
                //pictureBox1.Image = Image.FromFile(txt);
                //textBox4.Text = txt;
            }
            else
            {
                MessageBox.Show("No Item is selected");
            }
            capLength();
            fileSize();
        }
        

        // BUTTON - Rotate Image
        private void button1_Click_2(object sender, EventArgs e)
        {
            RotateImage();
        }



        // METHOD - Rotate Image
        public Image RotateImage()
        {
            //var bmp = new Bitmap(textBox4.Text);
            var bmp = pictureBox1.Image;

            bmp.RotateFlip(RotateFlipType.Rotate270FlipNone);
            return pictureBox1.Image = bmp;
        }


        // Save rotated Image
        private void button5_Click(object sender, EventArgs e)
        {
            var fd = new SaveFileDialog();
            //fd.FileName = textBox4.Text;
            string sourceDirectory = Path.GetDirectoryName(textBox4.Text);
            fd.InitialDirectory = Path.GetFullPath(sourceDirectory);
            fd.RestoreDirectory = true;
            fd.OverwritePrompt = true;


            //fd.Filter = "Bmp(*.BMP;)|*.BMP;| Jpg(*Jpg)|*.jpg";
            fd.Filter =  "Jpg(*Jpg)|*.jpg";
            if (fd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                if (System.IO.File.Exists(textBox4.Text))
                    System.IO.File.Delete(textBox4.Text);

                pictureBox1.Image.Save(fd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                /*
                switch (Path.GetExtension(fd.FileName))
                {
                    case ".BMP":
                        
                        
                        //pictureBox1.Image.Save(fd.FileName, System.Drawing.Imaging.ImageFormat.Bmp);
                        break;
                    case ".Jpg":
                        pictureBox1.Image.Save(fd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                        break;
                    default:
                        break;
                } 
                */
            }
        }


        // MENU - Drag & Drop Individual Images
        //void Form1_DragDrop(object sender, DragEventArgs e)
        void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
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
                        Console.WriteLine(id);
                        sb.Append(id + "\n");
                    }
                    MessageBox.Show(sb.ToString());
                }
                else
                // Otherwise load the files
                {
                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        string fFull = Path.GetFullPath(fileNames[i]);

                        string fileNam = Path.GetFileNameWithoutExtension(fileNames[i]);
                        //Image img = Image.FromFile(fileNameFull);
                        Bitmap bmp1 = new Bitmap(fFull);                       

                        //Object[] row = new object[] { fileNam, img, "", fileNameFull };
                        Object[] row = new object[] { fileNam, bmp1, "Insert caption here", fFull };
                        dataGridView1.Rows.Add(row);
                        /* 
                        string fileNam = Path.GetFileNameWithoutExtension(fileNames[0]);
                        Image img = Image.FromFile(fileNameFull);
                        Object[] row = new object[] { fileNam, img, "Insert caption here", fileNameFull };
                        dataGridView1.Rows.Add(row);
                        dgLength();
                        */
                    }
                }
                dgLength();
                capLength();
                fileSize();
                fileSizeTotal();
            }
        }

        private void dataGridView1_DragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }


        // MENU - Drag & Drop Individual Images
        private void Form1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        // MENU - Drag & Drop Individual Images
        private void dataGridView1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }



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
                int y = e.CellBounds.Bottom - 25;         // your  font height

                string mystring = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                var result = mystring.Substring(mystring.Length - Math.Min(4, mystring.Length));

                e.Graphics.DrawString(result, e.CellStyle.Font,
                Brushes.Crimson, e.CellBounds.Left, y);

                //FillRectangle() and use MeasureString()
                System.Drawing.Rectangle rect = new System.Drawing.Rectangle(e.CellBounds.Location.X + 4, e.CellBounds.Location.Y + 4, 13, 13);

                e.Graphics.FillRectangle(Brushes.White, rect);
                //e.PaintContent(rect);


                e.Handled = true;                        // done with the image column 
            }
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


        // METHOD - calculate Image file size
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

        // METHOD - calculate dataGridView1 Length
        private void dgLength()
        {
            int dgRows = dataGridView1.Rows.Count;
            textBox1.Text = dgRows.ToString();
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
        


        // BUTTON - delete
        private void button1_Click_1(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(row);
                dataGridView1.ClearSelection();
                dgLength();
                fileSizeTotal();
            }
        }


        // BUTTON - UP
        private void button3_Click_1(object sender, EventArgs e)
        {
            DataGridView dgv = dataGridView1;
            try
            {
                int totalRows = dgv.Rows.Count;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                if (rowIndex == 0)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex - 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex - 1].Selected = true;

            }
            catch { }
        }


        // BUTTON - DOWN
        private void button4_Click(object sender, EventArgs e)
        {
            DataGridView dgv = dataGridView1;
            try
            {
                int totalRows = dgv.Rows.Count;
                Console.WriteLine("total rows {0}", totalRows);
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                Console.WriteLine("rowIndex {0}", rowIndex);
                if (rowIndex == totalRows - 1)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                Console.WriteLine("colIndex {0}", colIndex);
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex + 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex + 1].Selected = true;
                Console.WriteLine("dis{0}", rowIndex);
            }
            catch { }
        }


        // BUTTON - CREATE WORD DOC
        private void button2_Click(object sender, EventArgs e)
        {
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
                    oPara0.Format.SpaceAfter=0;
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

                    
                    InlineShape pic = rngTarget1.InlineShapes.AddPicture(fileName1, ref oMissing, ref oMissing, ref anchor);
                    
                    Shape sh = pic.ConvertToShape();
                    sh.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;
                    

                    // Windows runs as default at 96dpi (display) Macs run as default at 72 dpi (display)
                    // Assuming 72 points per inch
                    // 3.5 inches is 3.5*72 = 252
                    // 3.25 inches is 3.25*72 = 234

                    sh.Height = 252;

                    if (sh.Width >400)
                    {
                        sh.Width = 400;
                    }

                    sh.Left = (float)WdShapePosition.wdShapeCenter;
                    sh.Top = 0;

                    //Write substring into Word doc with a bullet before it.
                    rngTarget0.InsertBefore(caption + "\v");
                    oPara1.Format.SpaceAfter = 264;
                    //rngTarget1.InsertParagraphAfter();                
                }

                foreach (Shape ilPicture in oDoc.Shapes)
                {
                    //ilPicture.{compress the picture}
                }

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

















        /*
        // IMAGE CLICK - View larger image, dataGridView0
        private void dataGridView0_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            String txt = dataGridView0.CurrentRow.Cells[3].Value.ToString();
            Console.WriteLine(txt);
            if (txt != null)
            {
                Image newImage = Image.FromFile(txt);
                Process.Start(txt);
            }
            else
            {
                MessageBox.Show("No Item is selected");
            }
        }
        */
        //// BUTTON - Publish Word document
        //private void button3_Click(object sender, EventArgs e)
        //{
        //    //SaveFileDialog sfd = new SaveFileDialog();
        //    //sfd.Filter = "Word Documents (*.docx)|*.docx";
        //    //sfd.FileName = ".docx";
        //    //if (sfd.ShowDialog() == DialogResult.OK)
        //    {
        //        //CreateWordDoc(dataGridView1, sfd.FileName);              
        //        CreateWordDoc(dataGridView1);
        //    }
        //}



        //// Create Datable of datagridViewView0
        //private System.Data.DataTable GetDataTableFromDGV0(DataGridView dgv)
        //{
        //    System.Data.DataTable dt1 = new System.Data.DataTable();
        //    dt1.Columns.Add("fileName", typeof(string));
        //    dt1.Columns.Add("dataGridView0Bitmap", typeof(string));
        //    dt1.Columns.Add("dataGridView0Caption", typeof(string));
        //    dt1.Columns.Add("dataGridView0Path", typeof(string));

        //    object[] cellValues1 = new object[dgv.Columns.Count];
        //    foreach (DataGridViewRow row in dgv.Rows)
        //    {
        //        for (int ii = 0; ii < row.Cells.Count; ii++)
        //        {
        //            cellValues1[ii] = row.Cells[ii].Value;
        //        }
        //        dt1.Rows.Add(cellValues1);
        //    }
        //    return dt1;
        //}


        //// Link to readme on GitHub
        //private void readMeToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    // Launch browser to facebook...
        //    System.Diagnostics.Process.Start("https://github.com/danhagg/photolog_dataGrid");
        //}


        //// METHOD - Vary image quality
        //private void VaryQualityLevel()
        //{
        //    // Get a bitmap. The using statement ensures objects  
        //    // are automatically disposed from memory after use.  
        //    using (Bitmap bmp1 = new Bitmap(@"C:\TestPhoto.jpg"))
        //    {
        //        ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);

        //        // Create an Encoder object based on the GUID  
        //        // for the Quality parameter category.  
        //        System.Drawing.Imaging.Encoder myEncoder =
        //            System.Drawing.Imaging.Encoder.Quality;

        //        // Create an EncoderParameters object.  
        //        // An EncoderParameters object has an array of EncoderParameter  
        //        // objects. In this case, there is only one  
        //        // EncoderParameter object in the array.  
        //        EncoderParameters myEncoderParameters = new EncoderParameters(1);

        //        EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 50L);
        //        myEncoderParameters.Param[0] = myEncoderParameter;
        //        bmp1.Save(@"C:\Users\dhaggerty\Desktop\TestPhotoQualityFifty.jpg", jpgEncoder, myEncoderParameters);

        //        myEncoderParameter = new EncoderParameter(myEncoder, 100L);
        //        myEncoderParameters.Param[0] = myEncoderParameter;
        //        bmp1.Save(@"C:\Users\dhaggerty\Desktop\TestPhotoQualityHundred.jpg", jpgEncoder, myEncoderParameters);

        //        // Save the bitmap as a JPG file with zero quality level compression.  
        //        myEncoderParameter = new EncoderParameter(myEncoder, 0L);
        //        myEncoderParameters.Param[0] = myEncoderParameter;
        //        bmp1.Save(@"C:\Users\dhaggerty\Desktop\TestPhotoQualityZero.jpg", jpgEncoder, myEncoderParameters);
        //    }
        //}


        //// METHOD Image quality Encoder
        //private ImageCodecInfo GetEncoder(ImageFormat format)
        //{
        //    ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
        //    foreach (ImageCodecInfo codec in codecs)
        //    {
        //        if (codec.FormatID == format.Guid)
        //        {
        //            return codec;
        //        }
        //    }
        //    return null;
        //}


        //// BUTTON - Vary image quality
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    VaryQualityLevel();
        //}

        //private void readDocsToolStripMenuItem1_Click(object sender, EventArgs e)
        //{
        //    // Launch browser to facebook...
        //    System.Diagnostics.Process.Start("https://github.com/danhagg/photolog_dataGrid");
        //}


        /*
        // MENU - Drag & Drop Individual Images
        void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileNames = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (fileNames != null && fileNames.Length != 0)
            {

                //var array0 = dataGridView0.Rows.Cast<DataGridViewRow>()
                                             //.Select(x => x.Cells[3].Value.ToString().Trim()).ToArray();

                var array1 = dataGridView1.Rows.Cast<DataGridViewRow>()
                                             .Select(x => x.Cells[3].Value.ToString().Trim()).ToArray();

                string fileNameFull = fileNames[0];

                //int pos0 = Array.IndexOf(array0, fileNameFull);
                int pos1 = Array.IndexOf(array1, fileNameFull);
                //if (pos0 > -1)
                //{
                //    StringBuilder sb = new StringBuilder("The following file is already in the left-hand list: ");
                //    sb.AppendLine();
                //    sb.Append(fileNameFull);
                //    sb.AppendLine();
                //    MessageBox.Show(sb.ToString());
                //}
                if (pos1 > -1)
                {
                    StringBuilder sb = new StringBuilder("The following file is already in your list: ");
                    sb.AppendLine();
                    sb.Append(fileNameFull);
                    sb.AppendLine();
                    MessageBox.Show(sb.ToString());
                }
                else
                {
                    string fileNam = Path.GetFileNameWithoutExtension(fileNames[0]);
                    Image img = Image.FromFile(fileNameFull);
                    Object[] row = new object[] { fileNam, img, "Insert caption here", fileNameFull };
                    dataGridView1.Rows.Add(row);
                    dgLength();
                }
            }
        }
    */


        /*
    // METHOD - Move selected items from grid0 to grid1
    private void grid0_to_grid1(DataGridView source, DataGridView target)
    {
    foreach (DataGridViewRow row in dataGridView0.SelectedRows)
    {
        string file = dataGridView0.CurrentRow.Cells[0].Value.ToString();
        String pth = dataGridView0.CurrentRow.Cells[3].Value.ToString();
        Image img = Image.FromFile(pth);
        dataGridView1.Rows.Add(file, img, "Insert Caption Here", pth);
        dataGridView0.Rows.Remove(row);
        dataGridView0.ClearSelection();
    }
    dgLength();
    }
    */

        /*
    // METHOD - Move selected items from grid1 to grid0
    private void grid1_to_grid0(DataGridView source, DataGridView target)
    {
        foreach (DataGridViewRow row in dataGridView1.SelectedRows)
        {
            string file = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            String pth = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            Image img = Image.FromFile(pth);
            dataGridView0.Rows.Add(file, img, "", pth);
            dataGridView1.Rows.Remove(row);
            dataGridView1.ClearSelection();
        }
        dgLength();
    }


    // BUTTON - Move selected items from dataGridView0 to dataGridView1
    private void addDGButton_Click(object sender, EventArgs e)
    {
        grid0_to_grid1(dataGridView0, dataGridView1);
    }


    // BUTTON - Move selected items from dataGridView1 to dataGridView0
    private void removeDGButton_Click(object sender, EventArgs e)
    {
        grid1_to_grid0(dataGridView1, dataGridView0);
    }
    */

        //// BUTTON - Up
        //private void upButton_Click(object sender, EventArgs e)
        //{
        //    DataGridView dgv = dataGridView1;
        //    try
        //    {
        //        int totalRows = dgv.Rows.Count;
        //        // get index of the row for the selected cell
        //        int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
        //        if (rowIndex == 0)
        //            return;
        //        // get index of the column for the selected cell
        //        int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
        //        DataGridViewRow selectedRow = dgv.Rows[rowIndex];
        //        dgv.Rows.Remove(selectedRow);
        //        dgv.Rows.Insert(rowIndex - 1, selectedRow);
        //        dgv.ClearSelection();
        //        dgv.Rows[rowIndex - 1].Cells[colIndex].Selected = true;
        //    }
        //    catch { }
        //}


        //// BUTTON - Down
        //private void downButton_Click(object sender, EventArgs e)
        //{
        //    DataGridView dgv = dataGridView1;
        //    try
        //    {
        //        int totalRows = dgv.Rows.Count;
        //        // get index of the row for the selected cell
        //        int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
        //        if (rowIndex == totalRows - 1)
        //            return;
        //        // get index of the column for the selected cell
        //        int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
        //        DataGridViewRow selectedRow = dgv.Rows[rowIndex];
        //        dgv.Rows.Remove(selectedRow);
        //        dgv.Rows.Insert(rowIndex + 1, selectedRow);
        //        dgv.ClearSelection();
        //        dgv.Rows[rowIndex + 1].Cells[colIndex].Selected = true;
        //    }
        //    catch { }
        //}


        /*
    // Overlay file name on top of image
    private void dataGridView0_CellPainting(object sender,
                                DataGridViewCellPaintingEventArgs e)
    {
    if (e.RowIndex < 0) return;                  // no image in the header
    if (e.ColumnIndex == this.dataGridView0.Columns["Column1"].Index)

    {
        e.PaintBackground(e.ClipBounds, false);  // no highlighting
        e.PaintContent(e.ClipBounds);

        // calculate the location of your text..:
        int y = e.CellBounds.Bottom - 20;         // your  font height
        e.Graphics.DrawString(dataGridView0.Rows[e.RowIndex].Cells[0].Value.ToString(), e.CellStyle.Font,
        Brushes.Crimson, e.CellBounds.Left, y);
        e.Handled = true;                        // done with the image column 
    }
    }
    */

        //private void dataGridView1_DragOver(object sender, DragEventArgs e)
        //{
        //    e.Effect = DragDropEffects.All;
        //}


        //// MENU - Drag & Drop Individual Images
        //private void loadIndividualImagesToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog ofd = new OpenFileDialog();
        //    ofd.InitialDirectory = @"C:\";
        //    ofd.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png";
        //    ofd.Title = "Drag Images into Assembly Line";
        //    ofd.ShowDialog();
        //}



        //// MENU - Open Folder
        //private void openToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    // create instance of folderBrowserDialog class
        //    FolderBrowserDialog fbd = new FolderBrowserDialog();
        //    // set root folder
        //    // fbd.RootFolder = Environment.SpecialFolder.MyDocuments;
        //    fbd.Description = "Choose an UNZIPPED folder of pictures to upload";

        //    // check user selects pass
        //    if (fbd.ShowDialog() == DialogResult.OK)
        //    {
        //        // clear previous data
        //        // add code

        //        string[] patterns = new[] { "*.jpg", "*.jpeg", "*.jpe", "*.jif", "*.jfif", "*.jfi", "*.webp", "*.gif", "*.png", "*.apng", "*.bmp", "*.dib", "*.tiff", "*.tif", "*.svg", "*.svgz", "*.ico", "*.xbm" };
        //        string[] files = CustomDirectoryTools.GetFiles(fbd.SelectedPath, patterns);

        //        if (files.Length == 0)
        //        {
        //            MessageBox.Show("No images in folder");
        //        }

        //        // iterate over selected folders files and load ALL images to dataGridView0
        //        for (int i = 0; i < files.Length; i++)
        //        {
        //            string fileNameFull = Path.GetFullPath(files[i]);

        //            string fileNam = Path.GetFileNameWithoutExtension(files[i]);
        //            //Image img = Image.FromFile(fileNameFull);
        //            Bitmap bmp1 = new Bitmap(fileNameFull);

        //            //Object[] row = new object[] { fileNam, img, "", fileNameFull };
        //            Object[] row = new object[] { fileNam, bmp1, "", fileNameFull };
        //            dataGridView1.Rows.Add(row);
        //        }
        //    }
        //}

    }
}




