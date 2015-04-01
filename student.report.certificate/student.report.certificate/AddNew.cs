using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SH.Student.Diploma
{
    public partial class AddNew : FISCA.Presentation.Controls.BaseForm
    {
        public string name = "";
        public Byte[] Template = null;
        public AddNew()
        {
            InitializeComponent();
        }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.name = textBoxX1.Text;
            this.Close();
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            DevComponents.DotNetBar.Controls.CheckBoxX cb = (DevComponents.DotNetBar.Controls.CheckBoxX)sender;
            switch (cb.Name)
            {
                case "check_default"://"使用預設樣板":
                    if (cb.Checked)
                        Template = null;
                    break;
                case "check_custom"://使用自訂樣板
                    if (cb.Checked && Template == null)
                    {
                        try
                        {
                            OpenFileDialog ofd = new OpenFileDialog();
                            ofd.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";

                            if (ofd.ShowDialog() == DialogResult.OK)
                                Template = File.ReadAllBytes(ofd.FileName);
                            else
                                check_default.Checked = true;
                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("檔案開啟錯誤,請檢查檔案是否開啟中!!");
                        }
                    }
                    break;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel ll = (LinkLabel)sender;
            Byte[] doc = null;
            switch (ll.Text)
            {
                case "檢視樣板":
                    doc = Properties.Resources.證明書範本;
                    break;
                case "檢視合併欄位總表":
                    doc = Properties.Resources.證明書範本;
                    break;
                default:
                    return;
            }
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = ll.Text.Replace("檢視", "");
                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Aspose.Words.Document doc2 = new Aspose.Words.Document(new MemoryStream(doc));
                    doc2.Save(SaveFileDialog1.FileName);
                    System.Diagnostics.Process.Start(SaveFileDialog1.FileName);
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
            }
        }
    }
}
