using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using System.IO;
using System.Drawing.Imaging;

using ZXing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Ghostscript.NET.Rasterizer;
using System.Diagnostics;
namespace streamDecoding
{
    public partial class Form1 : Form
    {
        public string outputFolder4img = @"\4img\";
        public string outputFolder4qr = @"\4code\";
        public string qrEmpty = @"\EmptyCode\";
        public string blureImagePath = @"\blur\";
        public string currentFolder = Directory.GetCurrentDirectory();

        public string qr2str = string.Empty;
        public string qr2strNext = string.Empty;
        public string q = string.Empty;
        public int qq = 0;
        public int x = 0, y = 0;
        //public int oz = 0;

        //Form f2 = new Form();
        public Graphics gBlack;
        public Graphics gRed;

        public int rectX;
        public int rectY;
        public int rectW;
        public int rectH;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Last().ToString() != @"\")
                textBox1.Text += @"\";
            if (textBox2.Text.Last().ToString() != @"\")
                textBox2.Text += @"\";
            string inputFolder = textBox2.Text;     //2
            string endOutputFolder = textBox1.Text; //1

                //1 собираем все PDF и бежим по ним, сохраняем их в виде jpeg постранично. формат: *имяPDF*_*НомерСтраницы*.jpeg
                string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf", SearchOption.TopDirectoryOnly);
            if (pdfFiles.Count() == 0)
            {
                MessageBox.Show("No *.pdf in InputFolder!");
            }
            else
            {
                progressBar1.Maximum = 0;
                progressBar1.Visible = true;
                int colDocs = pdfFiles.Count();

                // кол-во листов для всех документов. Для progressBar (по идее, их нужно умножать на 3 и добавлять +1 каждый раз по 3 раза соответственно 
                //(1-tojpeg, 2-jpegToRect, 3-decode(хотя, тут, наверно, может и больше)))
                foreach (string pdfFile in pdfFiles.OrderBy(pdfFile => File.GetCreationTime(pdfFile)))
                {
                    using (var rasterizer = new GhostscriptRasterizer())
                    {
                        rasterizer.Open(pdfFile);
                        progressBar1.Maximum += rasterizer.PageCount * 3;
                    }
                }

                foreach (string pdfFile in pdfFiles.OrderBy(pdfFile => File.GetCreationTime(pdfFile)))
                {
                    string fileName = Path.GetFileNameWithoutExtension(pdfFile);
                    GOpdfToJpeg(pdfFile, fileName, colDocs);
                    

                    //2 будем выделять область сканирования из сохранённых .jpeg
                    string[] jpegFiles = Directory.GetFiles(currentFolder + outputFolder4img, "*.jpeg");
                    foreach (string jpegFile in jpegFiles.OrderBy(jpegFile => File.GetCreationTime(jpegFile)))
                    {
                        string fileName1 = Path.GetFileNameWithoutExtension(jpegFile);
                        GOjpegToRect(jpegFile, fileName1);
                        progressBar1.Value += 1;
                    }
                    //3 пробуем декодировать изображения из outputFolder4qr
                    string[] qrFiles = Directory.GetFiles(currentFolder + outputFolder4qr, "*.jpeg");
                    foreach (string qrFile in qrFiles.OrderBy(qrFile => File.GetCreationTime(qrFile)))
                    {
                        string fileName2 = Path.GetFileNameWithoutExtension(qrFile);
                        GOqrDecoding(qrFile, fileName2, endOutputFolder, inputFolder);
                        progressBar1.Value += 1;
                    }

                    //удаляем temp файлы
                    string[] allPath = { currentFolder + @"\4img\", currentFolder + @"\blur\", currentFolder + @"\4code\" };
                    DeleteTempFiles(allPath);

                }
            }
            progressBar1.Visible = false;
            progressBar1.Value = 0;
            progressBar1.Maximum = 0;
            //progressBar1.Dispose();

            //string[] allPath = {currentFolder + @"\4img\", currentFolder + @"\blur\", currentFolder + @"\4code\"};
            MessageBox.Show("complete");
            //DeleteTempFiles(allPath);
            Process.Start("explorer.exe", endOutputFolder);

            if (Directory.GetFiles(currentFolder + qrEmpty).Count() != 0)
            {
                MessageBox.Show("not decode pages");
                Process.Start("explorer.exe", currentFolder + qrEmpty);
            }
            
        }

        private void DeleteTempFiles (string[] allPath)
        {
            //MessageBox.Show("1");
            foreach (string path in allPath)
            {
                var files = Directory.GetFiles(path, "*", SearchOption.TopDirectoryOnly);
                if (files.Count() != 0)
                {
                    foreach (string file in files)
                    {
                        File.Delete(file);
                    }
                }
            }
        }

        private void Decode(string pdfFileName, int pageInt, Bitmap qrimg, string inputFolder, string endOutputFolder)
        {
            //старый документ, забираем из него страницу, на которой нашли QR
            PdfDocument inPdfDoc = PdfReader.Open(inputFolder + pdfFileName + ".pdf", PdfDocumentOpenMode.Import);
            PdfPage inPdfPage = inPdfDoc.Pages[pageInt - 1]; //тут ХЗ какой индекс. Поставим = 0. Если нет, то поменять на = 1
            //создаём новый PDF
            PdfDocument nPdfDoc = new PdfDocument();
            nPdfDoc.Info.Title = qr2str; //в инфо запишем значение из QR. Может потом получится работать с ним
            PdfPage nPage = nPdfDoc.AddPage(inPdfPage);

            if (Directory.GetFiles(endOutputFolder, "*" + qr2str + ".pdf").Count() == 0)
            {
                nPdfDoc.Save(endOutputFolder + qr2str + ".pdf");
                qr2strNext = qr2str;
            }
            else
            {
                q = "_";
                qq = qq + 1;
                qr2str = q + qr2str;
                nPdfDoc.Save(endOutputFolder + qr2str + ".pdf");
                qr2strNext = qr2str;
                qr2str = qr2str.Substring(qq, 10);
                //ProcessVariable oqq = process.GetVariableByName("qq");
                //oqq.Value = qq;
            }

            //закрываем
            inPdfDoc.Close();
            inPdfDoc.Dispose();

            nPdfDoc.Close();
            nPdfDoc.Dispose();
            qrimg.Dispose();
        }

        private Bitmap Blur(Bitmap image, Int32 blurSize)
        {
            return Blured(image, new Rectangle(0, 0, image.Width, image.Height), blurSize);
        }

        private Bitmap Blured(Bitmap image, Rectangle rectangle, Int32 blurSize)
        {
            Bitmap blurred = new Bitmap(image.Width, image.Height);


            using (Graphics graphics = Graphics.FromImage(blurred))
                graphics.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height),
                    new Rectangle(0, 0, image.Width, image.Height), GraphicsUnit.Pixel);

            for (int xx = rectangle.X; xx < rectangle.X + rectangle.Width; xx++)
            {
                for (int yy = rectangle.Y; yy < rectangle.Y + rectangle.Height; yy++)
                {
                    int avgR = 0, avgG = 0, avgB = 0;
                    int blurPixelCount = 0;


                    for (int x = xx; (x < xx + blurSize && x < image.Width); x++)
                    {
                        for (int y = yy; (y < yy + blurSize && y < image.Height); y++)
                        {
                            Color pixel = blurred.GetPixel(x, y);
                            avgR += pixel.R;
                            avgG += pixel.G;
                            avgB += pixel.B;
                            blurPixelCount++;
                        }
                    }
                    avgR = avgR / blurPixelCount;
                    avgG = avgG / blurPixelCount;
                    avgB = avgB / blurPixelCount;


                    for (int x = xx; x < xx + blurSize && x < image.Width && x < rectangle.Width; x++)
                        for (int y = yy; y < yy + blurSize && y < image.Height && y < rectangle.Height; y++)
                            blurred.SetPixel(x, y, Color.FromArgb(avgR, avgG, avgB));
                }
            }
            image.Dispose();
            return blurred;
        }

        private void NoCode(string pdfFileName, int pageInt, string endOutputFolder, string inputFolder)
        {
            //тут нет qr
            //надо проверить, на наличте файла PDF. (только в том случае, если верхний лист был без QR, или если он не прочитался)
            //если его нет, то создать файл с листами, которые не получилось декодировать
            if (!File.Exists(endOutputFolder + qr2strNext + ".pdf"))
            {
                // если в папке с потерянными листами нет файла, то создаём его
                if (Directory.GetFiles(currentFolder + qrEmpty, "*.pdf").Length == 0)
                {
                    //будем делать новый PDF и добавлять в него empty pages
                    PdfDocument inPdfDoc = PdfReader.Open(inputFolder + pdfFileName + ".pdf", PdfDocumentOpenMode.Import);
                    PdfPage inPdfPage = inPdfDoc.Pages[pageInt - 1]; //тут ХЗ какой индекс. Поставим = 0. Если нет, то поменять на = 1

                    PdfDocument emptyPdfDoc = new PdfDocument();
                    PdfPage emptyPdfPage = emptyPdfDoc.AddPage(inPdfPage);
                    emptyPdfDoc.Save(Directory.GetCurrentDirectory() + qrEmpty + "emptyPages" + ".pdf");

                    //закрываем
                    inPdfDoc.Close();
                    inPdfDoc.Dispose();

                    emptyPdfDoc.Close();
                    emptyPdfDoc.Dispose();
                }
                else    //если он есть, то добавляем в него новый лист
                {
                    PdfDocument inPdfDoc = PdfReader.Open(inputFolder + pdfFileName + ".pdf", PdfDocumentOpenMode.Import);
                    PdfPage inPdfPage = inPdfDoc.Pages[pageInt - 1]; //тут ХЗ какой индекс. Поставим = 0. Если нет, то поменять на = 1

                    //будем делать новые страницы в nPdfDoc
                    PdfDocument nPdfDoc = PdfReader.Open(currentFolder + qrEmpty + "emptyPages" + ".pdf", PdfDocumentOpenMode.Modify);
                    PdfPage nPage = nPdfDoc.AddPage(inPdfPage);
                    nPdfDoc.Save(currentFolder + qrEmpty + "emptyPages" + ".pdf");

                    //закрываем
                    inPdfDoc.Close();
                    inPdfDoc.Dispose();

                    nPdfDoc.Close();
                    nPdfDoc.Dispose();
                }
            }
            else // если он есть...
            {
                //старый документ, забираем из него страницу на которой находимся
                PdfDocument inPdfDoc = PdfReader.Open(inputFolder + pdfFileName + ".pdf", PdfDocumentOpenMode.Import);
                PdfPage inPdfPage = inPdfDoc.Pages[pageInt - 1]; //тут ХЗ какой индекс. Поставим = 0. Если нет, то поменять на = 1

                //будем делать новые страницы в nPdfDoc
                PdfDocument nPdfDoc = PdfReader.Open(endOutputFolder + qr2strNext + ".pdf", PdfDocumentOpenMode.Modify);
                PdfPage nPage = nPdfDoc.AddPage(inPdfPage);
                nPdfDoc.Save(endOutputFolder + qr2strNext + ".pdf");

                //закрываем
                inPdfDoc.Close();
                inPdfDoc.Dispose();

                nPdfDoc.Close();
                nPdfDoc.Dispose();
            }
        }

        private void GOqrDecoding(string inputFile, string outputFileName, string endOutputFolder, string inputFolder)
        {
            //сохраняем имя исходного файла и текущую страницу
            string pdfFileName = outputFileName.Substring(0, outputFileName.LastIndexOf('_'));
            string page = outputFileName.Substring(outputFileName.LastIndexOf('_') + 1);
            int pageInt = Int32.Parse(page);

            //Если получается, то будем создавать первую страницу Документа	
            Bitmap qrimg = Image.FromFile(inputFile) as Bitmap;
            if (CodeInPage(qrimg) == true)
            {

                qr2strNext = string.Empty;
                IBarcodeReader decoder = new BarcodeReader { AutoRotate = true, TryInverted = true };
                decoder.Options.PossibleFormats = new List<BarcodeFormat>();
                decoder.Options.PossibleFormats.Add(BarcodeFormat.CODE_128);
                try
                {
                    qr2str = decoder.Decode(qrimg).ToString();
                    Decode(pdfFileName, pageInt, qrimg, inputFolder, endOutputFolder);
                    //
                    qrimg.Dispose();
                }
                catch
                {
                    //пробуем ещё раз
                    for (int z = 1; qr2str.Length != 10 && z <= 5; z++)
                    {	
                        qrimg = Blur(qrimg, 2);
                        qrimg.Save(currentFolder + blureImagePath + pdfFileName + page + z.ToString() + "blure.jpeg");
                        try
                        {
                            qr2str = decoder.Decode(qrimg).ToString();
                            Decode(pdfFileName, pageInt, qrimg, inputFolder, endOutputFolder);
                            //
                            qrimg.Dispose();
                        }
                        catch
                        {
                            //qrimg.Dispose();
                        }
                    }
                }
                if (qr2str.Length != 10)
                {
                    NoCode(pdfFileName, pageInt, endOutputFolder, inputFolder);
                    qrimg.Dispose();
                }
                else
                {
                    qr2str = string.Empty;
                    qrimg.Dispose();
                }
            }
            else
            {
                NoCode(pdfFileName, pageInt, endOutputFolder, inputFolder);
                qrimg.Dispose();
            }
            qrimg.Dispose();
        }

        private bool CodeInPage(Bitmap img)
        {
            using (Graphics graphics = Graphics.FromImage(img))
                graphics.DrawImage(img, new Rectangle(0, 0, img.Width, img.Height),
                    new Rectangle(0, 0, img.Width, img.Height), GraphicsUnit.Pixel);
            int x = 0;
            int y = 0;
            int z = 0;
            int black = 0;
            for (x = 0; x < img.Width; x++)
            {
                for (y = 0; y < img.Height; y++)
                {
                    Color pixelColor = img.GetPixel(x, y);
                    z = z + 1;
                    if (pixelColor.G < 100)
                    {
                        black = black + 1;
                    }
                }
            }
            float pr = (((float)black / (float)z) * 100);
            if (pr > 5)
                return true;
            else
                return false;
            
        }

        private void GOjpegToRect(string inputFile, string outputFileName)
        {
            //область сканирования
            Bitmap pdf2Jpeg = Image.FromFile(inputFile) as Bitmap;
            Rectangle rect;
            Bitmap nBitmap;
            if (checkBox1.Checked == true)
            {
                rect = new Rectangle(pdf2Jpeg.Width / 4,
                    0,
                    pdf2Jpeg.Width - (pdf2Jpeg.Width / 2),
                    pdf2Jpeg.Height / 32);
                nBitmap = new Bitmap(rect.Width, rect.Height);
            }
            else
            {
                if (pdf2Jpeg.Width < ((pdf2Jpeg.Width / rectW) + (pdf2Jpeg.Width / rectX)))
                {
                    //MessageBox.Show("!!!");
                    rect = new Rectangle(pdf2Jpeg.Width / rectX,
                        pdf2Jpeg.Height / rectY,
                        (pdf2Jpeg.Width / rectW) - (pdf2Jpeg.Width / rectX),
                        pdf2Jpeg.Height / rectH);
                    nBitmap = new Bitmap(rect.Width, rect.Height);
                }
                else
                {
                    rect = new Rectangle(pdf2Jpeg.Width / rectX,
                        pdf2Jpeg.Height / rectY,
                        pdf2Jpeg.Width / rectW,
                        pdf2Jpeg.Height / rectH);
                    nBitmap = new Bitmap(rect.Width, rect.Height);
                }

            }
            //Bitmap nBitmap = new Bitmap(rect.Width, rect.Height);

            //делаем новую картинку для поиска QR
            Graphics g = Graphics.FromImage(nBitmap);
            g.DrawImage(pdf2Jpeg, -rect.X, -rect.Y);
            nBitmap.Save(currentFolder + outputFolder4qr + outputFileName + ".jpeg", ImageFormat.Jpeg);
            pdf2Jpeg.Dispose();
            g.Dispose();
            nBitmap.Dispose();
        }

        private void GOpdfToJpeg(string inputFile, string outputFileName, int colDocs)
        {
            var xDpi = 150; //set the x DPI
            var yDpi = 150; //set the y DPI

            using (var rasterizer = new GhostscriptRasterizer()) //create an instance for GhostscriptRasterizer
            {
                rasterizer.Open(inputFile); //opens the PDF file for rasterizing
                for (int page = 1; page <= rasterizer.PageCount; page++)
                {
                    //set the output image(jpeg's) complete path
                    string outputJpegPath = Path.Combine(currentFolder + outputFolder4img, string.Format("{0}.jpeg", outputFileName + "_" + page.ToString()));
                    string outputQrPath = Path.Combine(currentFolder + outputFolder4qr, string.Format("{0}.jpeg", outputFileName + "_" + page.ToString()));

                    //converts the PDF pages to jpeg's 
                    var pdf2Jpeg = rasterizer.GetPage(xDpi, yDpi, page);

                    //save the jpeg's
                    pdf2Jpeg.Save(outputJpegPath, ImageFormat.Jpeg);
                    pdf2Jpeg.Dispose();
                    progressBar1.Value += 1;
                }
                rasterizer.Dispose();
            }
        }

        private void CreateFolder(string path)
        {
            if (Directory.Exists(currentFolder + path) == false)
                Directory.CreateDirectory(currentFolder + path);
        }

        private string FolderBro(string text)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (text == string.Empty)
            {
                DialogResult result = folderBrowser.ShowDialog();
                text = folderBrowser.SelectedPath.ToString();
            }
            return text;
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = FolderBro(textBox1.Text);
            textBox1.Select(textBox1.Text.Length, textBox1.Text.Length);
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = FolderBro(textBox1.Text);
            textBox2.Select(textBox2.Text.Length, textBox2.Text.Length);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] folders = {outputFolder4img, outputFolder4qr, qrEmpty, blureImagePath};

            foreach (string folder in folders)
            {
               string folderPath = currentFolder + folder;
                if (Directory.Exists(folderPath) == false)
                    CreateFolder(folder);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            f2.Width = 210 *2;     //210;
            f2.Height = 297 *2;    //297;
            f2.MinimumSize = f2.Size;
            f2.MaximumSize = f2.Size;
            f2.BackColor = Color.White;

            f2.ShowDialog();
        }

        public void f2_MouseDown(object sender, EventArgs e)
        {
            x = MousePosition.X - (f2.Location.X + 3);
            y = MousePosition.Y - f2.Location.Y - 25;
            gBlack = f2.CreateGraphics();
            f2.MouseMove += new System.Windows.Forms.MouseEventHandler(this.f2_MouseMove);

            rectX = f2.Width / (x - 5);
            rectY = f2.Height / (y - 6);
        }

        public void f2_MouseUp(object sender, EventArgs e)
        {
            f2.MouseMove -= new System.Windows.Forms.MouseEventHandler(this.f2_MouseMove);
            DelGr(gRed);
            Pen pen = new Pen(Color.Black, 10);
            Gr(gBlack, pen);
            //rectX = f2.Width / x;
            //rectY = y - 10;
            rectW = (int)(f2.Width / ((MousePosition.X - (f2.Location.X + 3)) - x - pen.Width));
            rectH = f2.Height / ((MousePosition.Y - f2.Location.Y - 25) - y - (int)pen.Width);
            //MessageBox.Show(rectX.ToString() + " " + rectY.ToString() + " " + rectW.ToString() + " " + rectH.ToString() + " ");
        }

        public void f2_MouseMove(object sender, EventArgs e)
        {
            gRed = f2.CreateGraphics();
            Pen pen = new Pen(Brushes.DeepPink, 2);
            DelGr(gRed);
            Gr(gRed, pen);
        }

        private void Gr(Graphics g, Pen p)
        {
            Rectangle rec;
            if (x < MousePosition.X - (f2.Location.X + 3))
            {
                rec = new Rectangle(x, y, (MousePosition.X - (f2.Location.X + 3)) - x - (int)p.Width, (MousePosition.Y - f2.Location.Y - 25) - y - (int)p.Width);
                

                
            }
            else
            {
                rec = new Rectangle((MousePosition.X - (f2.Location.X + 3)), (MousePosition.Y - f2.Location.Y - 25), x - (MousePosition.X - (f2.Location.X + 3)), y - (MousePosition.Y - (f2.Location.Y + 3)));

            }
            g.DrawRectangle(p, rec);
        }

        private void f2_Load(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void DelGr(Graphics g)
        {
            if (g != null)
                g.Clear(Color.White);
        }
        
    }
}
