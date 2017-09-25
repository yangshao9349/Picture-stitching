using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ChatTool.Common;
using System.IO;
using System.Drawing.Imaging;

namespace 图片处理
{
    public partial class Main : Form
    {
        private List<string> files = new List<string>();
        private double bili = 1.5;
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {



        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;

                txtimgpath.Text = foldPath;
            }

        }

        private void btnstart_Click(object sender, EventArgs e)
        {
            if (txtimgpath.Text.ToString().Equals(""))
            {
                MessageBox.Show("请选择图片文件夹");
            }
            else if (txtxls.Text.ToString().Equals("")) {
                MessageBox.Show("请选择模板文件");
            }
            else if (txtout.Text.ToString().Equals(""))
            {
                MessageBox.Show("请选择输出文件夹");
            }
            else {
             
              

                chuli();
            }
        }

        private void chuli()
        {
            //获取所有图片
            DirectoryInfo TheFolder = new DirectoryInfo(txtimgpath.Text.ToString());
            foreach (FileInfo NextFile in TheFolder.GetFiles()) {

                if (NextFile.Extension.ToString().ToLower().Equals(".jpg")) {
                    string file = NextFile.FullName.ToString();
                    files.Add(file);
                }
              
            }
               

            DataTable dt = Excel.ExcelToDS(txtxls.Text.ToString()).Tables[0];
            //设置初始图片尺寸，比例固定1：1.5
            int w = 640;
            int h = 960;
            int column = 4;  //每行4列
            int height = 213;
       
            //开始循环处理图片
            foreach (var item in files)
            {
             
                var img = Image.FromFile(item);
            
                //旋转270°
                img = ImageHelper.RotateImg(img, 270);
                string filename = System.IO.Path.GetFileNameWithoutExtension(item);

               
                //画一个背景色为白色的画布
                Bitmap map = new Bitmap(w, h);
                Graphics g = Graphics.FromImage(map);
                g.FillRectangle(Brushes.White, 0, 0, w, h);



                int width = (w-20) / column;
           
                //第一行
                for (int i = 0; i < column; i++)
                {
                    g.DrawImage(img, width * i+10, 0+10, width+1, height);
                }
                //第二行
                for (int i = 0; i < column; i++)
                {
                    int k = 0;
                    if (i == 0)
                    {
                        k = height;
                    }
                    else
                    {
                        k = height * 1;
                    }
                    g.DrawImage(img, i * width+10, k+10, width+1, height);
                }


                int heigh2 = 16;
                //第三行 每行2列
                img = ImageHelper.RotateImg(img, 270);
                g.DrawImage(img, 0 + 10, height * 2 + heigh2, width * 2, height + 1);
                g.DrawImage(img, width * 2 + 10, height * 2 + heigh2, width * 2, height + 1);

                //第四行 每行2列
                g.DrawImage(img, 0 + 10, height * 3 + heigh2, width * 2, height);
                g.DrawImage(img, 320, height * 3 + heigh2, width * 2, height);
                
                //获取xls上面的数据 
                int n = 0;
                string schoolname = dt.Rows[n][7].ToString();
                string username = dt.Rows[n ][3].ToString();
                string sex = dt.Rows[n][4].ToString();
                string id = dt.Rows[n][5].ToString();
                string zhuanye = dt.Rows[n][11].ToString();
                string haoma = dt.Rows[n][9].ToString();
                 

                string str1 = schoolname+"      "+haoma+"    "+zhuanye;

                string str2 = username+"  "+sex+"  "+id;

                string str3 = txtbanquan.Text.ToString();

                //输出文字
               
                Font drawFont = new Font("宋体", 18);//显示的字符串使用的字体
                SolidBrush drawBrush = new SolidBrush(Color.Black);//写字符串用的刷子
                PointF drawPoint = new PointF(20, 880);//显示的字符串左上角的坐标
                g.DrawString(str1, drawFont, drawBrush, drawPoint);

                PointF drawPoint2 = new PointF(20, 910);//显示的字符串左上角的坐标
                g.DrawString(str2, drawFont, drawBrush, drawPoint2);

                //写蓝色的字
                Font bluefont = new Font("华文行楷", 14);//显示的字符串使用的字体
                SolidBrush bluebrush = new SolidBrush(Color.FromArgb(2,32,254));//写字符串用的刷子
                PointF drawPoint3 = new PointF(450, 920);//显示的字符串左上角的坐标
                g.DrawString(str3, bluefont, bluebrush, drawPoint3);
                

                //map.Save(txtout.Text + "\\" + id+".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
         
                string name = "";
                //如何选中则已身份证命名
                if (checkBox1.Checked)
                {
                    name = id;
                }
                else {
                    name = filename;
                }
                map =ImageHelper.ResizeImage(map, Convert.ToInt32(txtkuan.Text.ToString()), Convert.ToInt32(txtgao.Text), 0);

             //   KiSaveAsJPEG(map, txtout.Text + "\\" + name + ".jpg", bili);

                map.Save(txtout.Text + "\\" + name + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                g.Dispose();

                n++;
            }
            MessageBox.Show("处理完成");
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;//允许同时选择多个文件

            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "所有文件(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                for (int fi = 0; fi < fileDialog.FileNames.Length; fi++)
                {
                    string file = fileDialog.FileNames[fi].ToString();
                    txtxls.Text = file;

                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                txtout.Text = foldPath;
            }
        }

        public  bool KiSaveAsJPEG(Bitmap bmp, string FileName, int Qty)
        {
            try
            {
                EncoderParameter p;
                EncoderParameters ps;

                ps = new EncoderParameters(1);

                p = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, Qty);
                ps.Param[0] = p;

                bmp.Save(FileName, GetCodecInfo("image/jpeg"), ps);

                return true;
            }
            catch
            {
                return false;
            }

        }
        /**/
        /// <summary>
        /// 保存JPG时用
        /// </summary>
        /// <param name="mimeType"></param>
        /// <returns>得到指定mimeType的ImageCodecInfo</returns>
        private  ImageCodecInfo GetCodecInfo(string mimeType)
        {
            ImageCodecInfo[] CodecInfo = ImageCodecInfo.GetImageEncoders();
            foreach (ImageCodecInfo ici in CodecInfo)
            {
                if (ici.MimeType == mimeType) return ici;
            }
            return null;
        }

        private void txtgao_TextChanged(object sender, EventArgs e)
        {
            double width = Convert.ToInt32(txtgao.Text) / bili;
            txtkuan.Text = Convert.ToInt32(width).ToString();  
        }

        private void txtkuan_TextChanged(object sender, EventArgs e)
        {
            double height = Convert.ToInt32(txtkuan.Text) * bili;
            txtgao.Text = Convert.ToInt32(height).ToString();  
        }


    }
}
