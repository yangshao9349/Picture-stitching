namespace 图片处理
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtxls = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtout = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.btnstart = new System.Windows.Forms.Button();
            this.txtimgpath = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtbanquan = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtgao = new System.Windows.Forms.TextBox();
            this.txtkuan = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(456, 22);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(97, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "选择图片文件夹";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "图片";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "模板文件";
            // 
            // txtxls
            // 
            this.txtxls.Location = new System.Drawing.Point(79, 53);
            this.txtxls.Name = "txtxls";
            this.txtxls.Size = new System.Drawing.Size(356, 21);
            this.txtxls.TabIndex = 4;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(456, 53);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 5;
            this.button2.Text = "选择文件";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "输出文件夹";
            // 
            // txtout
            // 
            this.txtout.Location = new System.Drawing.Point(79, 94);
            this.txtout.Name = "txtout";
            this.txtout.Size = new System.Drawing.Size(356, 21);
            this.txtout.TabIndex = 7;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(456, 92);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 8;
            this.button3.Text = "选择文件夹";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btnstart
            // 
            this.btnstart.Location = new System.Drawing.Point(191, 273);
            this.btnstart.Name = "btnstart";
            this.btnstart.Size = new System.Drawing.Size(147, 48);
            this.btnstart.TabIndex = 9;
            this.btnstart.Text = "执行";
            this.btnstart.UseVisualStyleBackColor = true;
            this.btnstart.Click += new System.EventHandler(this.btnstart_Click);
            // 
            // txtimgpath
            // 
            this.txtimgpath.Location = new System.Drawing.Point(79, 24);
            this.txtimgpath.Name = "txtimgpath";
            this.txtimgpath.Size = new System.Drawing.Size(356, 21);
            this.txtimgpath.TabIndex = 10;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(79, 141);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(60, 16);
            this.checkBox1.TabIndex = 11;
            this.checkBox1.Text = "身份证";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 141);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "命名方式";
            // 
            // txtbanquan
            // 
            this.txtbanquan.Location = new System.Drawing.Point(79, 168);
            this.txtbanquan.Name = "txtbanquan";
            this.txtbanquan.Size = new System.Drawing.Size(356, 21);
            this.txtbanquan.TabIndex = 13;
            this.txtbanquan.Text = "贵州高校采集组制作";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(34, 171);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 12);
            this.label5.TabIndex = 14;
            this.label5.Text = "版权";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(34, 212);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 15;
            this.label6.Text = "高";
            // 
            // txtgao
            // 
            this.txtgao.Location = new System.Drawing.Point(79, 209);
            this.txtgao.Name = "txtgao";
            this.txtgao.Size = new System.Drawing.Size(60, 21);
            this.txtgao.TabIndex = 16;
            this.txtgao.Text = "960";
            this.txtgao.TextChanged += new System.EventHandler(this.txtgao_TextChanged);
            // 
            // txtkuan
            // 
            this.txtkuan.Location = new System.Drawing.Point(191, 209);
            this.txtkuan.Name = "txtkuan";
            this.txtkuan.Size = new System.Drawing.Size(97, 21);
            this.txtkuan.TabIndex = 17;
            this.txtkuan.Text = "640";
            this.txtkuan.TextChanged += new System.EventHandler(this.txtkuan_TextChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(157, 212);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(17, 12);
            this.label7.TabIndex = 18;
            this.label7.Text = "宽";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(579, 333);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtkuan);
            this.Controls.Add(this.txtgao);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtbanquan);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.txtimgpath);
            this.Controls.Add(this.btnstart);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.txtout);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.txtxls);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "Main";
            this.Text = "图片处理";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtxls;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtout;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button btnstart;
        private System.Windows.Forms.TextBox txtimgpath;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtbanquan;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtgao;
        private System.Windows.Forms.TextBox txtkuan;
        private System.Windows.Forms.Label label7;

    }
}

