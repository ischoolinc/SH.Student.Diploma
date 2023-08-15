using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Aspose.Words;
using Aspose.Words.Drawing;
using Campus.Report2014;
using K12.Data;
using SHSchool.Data;
using System.IO;
using System.Text.RegularExpressions;
namespace SH.Student.Diploma
{
    public partial class MainForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bgw = new BackgroundWorker();

        private const string config = "plugins.student.report.certificate.huangwc.config";
        private Dictionary<string, ReportConfiguration> custConfigs = new Dictionary<string, ReportConfiguration>();
        ReportConfiguration conf = new Campus.Report2014.ReportConfiguration(config);
        public string current = "";

        public MainForm()
        {
            InitializeComponent();
            this.Text = "學籍相關證明書";
            #region 設定comboBox選單
            foreach (string item in getCustConfig())
            {
                if (!string.IsNullOrWhiteSpace(item) && !custConfigs.Keys.Contains(item))
                {
                    custConfigs.Add(item, new ReportConfiguration(configNameRule(item)));
                    comboBoxEx1.Items.Add(item);
                }
            }
            comboBoxEx1.Items.Add("新增");
            comboBoxEx1.SelectedIndex = 0;
            #endregion
            _bgw.DoWork += new DoWorkEventHandler(_bgw_DoWork);
            _bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgw_RunWorkerCompleted);
        }

        void _bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            Document document = new Document();
            Byte[] template = (custConfigs[current].Template != null) //單頁範本
                 ? custConfigs[current].Template.ToBinary()
                 : new Campus.Report2014.ReportTemplate(Properties.Resources.證明書範本, Campus.Report2014.TemplateType.docx).ToBinary();
            List<string> ids = K12.Presentation.NLDPanels.Student.SelectedSource;


            List<SHStudentRecord> srl = SHStudent.SelectByIDs(ids);
            //離校資訊
            Dictionary<string, SHLeaveInfoRecord> dshlir = SHLeaveInfo.SelectByStudentIDs(ids).ToDictionary(x => x.RefStudentID, x => x);
            //畢業異動
            Dictionary<string, SHUpdateRecordRecord> dshurr = new Dictionary<string, SHUpdateRecordRecord>();

            // 新生與轉入異動，主要取得入學年
            Dictionary<string, SHUpdateRecordRecord> dsEntryRceDict = new Dictionary<string, SHUpdateRecordRecord>();

            foreach (SHUpdateRecordRecord shurr in SHUpdateRecord.SelectByStudentIDs(ids))
            {
                if (shurr.UpdateCode == "501")
                {
                    if (dshurr.ContainsKey(shurr.StudentID))
                    {
                        if (dshurr[shurr.StudentID].UpdateDate.CompareTo(shurr.UpdateDate) == 1)
                        {
                            dshurr[shurr.StudentID] = shurr;
                        }
                    }
                    else dshurr.Add(shurr.StudentID, shurr);
                }

                int intCode;
                if (int.TryParse(shurr.UpdateCode, out intCode))
                {
                    // 新生或轉入相關異動
                    if (intCode < 200)
                    {
                        if (dsEntryRceDict.ContainsKey(shurr.StudentID))
                        {
                            DateTime dt1, dt2;
                            DateTime.TryParse(shurr.UpdateDate, out dt1);
                            DateTime.TryParse(dsEntryRceDict[shurr.StudentID].UpdateDate, out dt2);
                            if (dt1 > dt2)
                            {
                                dsEntryRceDict[shurr.StudentID] = shurr;
                            }
                        }
                        else
                            dsEntryRceDict.Add(shurr.StudentID, shurr);
                    }
                }
            }
            //入學照片
            Dictionary<string, string> dphoto_p = K12.Data.Photo.SelectFreshmanPhoto(K12.Presentation.NLDPanels.Student.SelectedSource);
            Dictionary<string, string> dphoto_g = K12.Data.Photo.SelectGraduatePhoto(K12.Presentation.NLDPanels.Student.SelectedSource);
            //科別中英文對照表
            Dictionary<string, string> dic_dept_ch_en = new Dictionary<string, string>();
            XmlElement Data = SmartSchool.Customization.Data.SystemInformation.Configuration["科別中英文對照表"];
            foreach (XmlElement var in Data)
            {
                if (!dic_dept_ch_en.ContainsKey(var.GetAttribute("Chinese")))
                {
                    dic_dept_ch_en.Add(var.GetAttribute("Chinese"), var.GetAttribute("English"));
                }
            }
            Dictionary<string, object> mailmerge = new Dictionary<string, object>();
            string 校內字號 = textBoxX1.Text, 校內字號英文 = textBoxX2.Text, 校長姓名 = "", 校長姓名英文 = "";
            if (K12.Data.School.Configuration["學校資訊"] != null && K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName") != null)
                校長姓名 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText;
            if (K12.Data.School.Configuration["學校資訊"] != null && K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName") != null)
                校長姓名英文 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorEnglishName").InnerText;

            Document each;
            foreach (SHStudentRecord sr in srl)
            {
                mailmerge.Clear();
                mailmerge.Add("學校全銜", School.ChineseName);
                mailmerge.Add("學校英文全銜", School.EnglishName);
                mailmerge.Add("目前學期", School.DefaultSemester);
                mailmerge.Add("目前學年度", School.DefaultSchoolYear);
                mailmerge.Add("校長姓名", 校長姓名);
                mailmerge.Add("校長姓名英文", 校長姓名英文);

                mailmerge.Add("民國年", DateTime.Today.Year - 1911);
                mailmerge.Add("英文年", DateTime.Today.Year);
                mailmerge.Add("月", DateTime.Today.Month);
                mailmerge.Add("英文月", DateTime.Today.ToString("MMMM", new System.Globalization.CultureInfo("en-US")));
                mailmerge.Add("英文月3", DateTime.Today.ToString("MMM", new System.Globalization.CultureInfo("en-US")));
                mailmerge.Add("日上標", daySuffix(DateTime.Today.Day.ToString()));
                mailmerge.Add("日", DateTime.Today.Day);

                mailmerge.Add("校內字號", 校內字號);
                mailmerge.Add("校內字號英文", 校內字號英文);

                #region 學生資料
                mailmerge.Add("學生姓名", sr.Name);
                mailmerge.Add("學生英文姓名", sr.EnglishName);
                mailmerge.Add("學生身分證號", sr.IDNumber);
                mailmerge.Add("學生目前學號", sr.StudentNumber);
                mailmerge.Add("性別", sr.Gender);

                SHClassRecord tmpcr;
                if ((tmpcr = sr.Class) != null)
                {
                    mailmerge.Add("學生目前班級", tmpcr.Name);
                    mailmerge.Add("學生目前年級", tmpcr.GradeYear);
                    if (tmpcr.Department != null)
                    {
                        mailmerge.Add("學生目前科別", tmpcr.Department.Name);
                        if (dic_dept_ch_en.ContainsKey(tmpcr.Department.Name))
                            mailmerge.Add("科別英文名稱", dic_dept_ch_en[tmpcr.Department.Name]);
                        else
                            mailmerge.Add("科別英文名稱", "");
                    }

                }
                mailmerge.Add("學生目前座號", sr.SeatNo);
                if (sr.Birthday.HasValue)
                {
                    mailmerge.Add("學生生日民國年", sr.Birthday.Value.Year - 1911);
                    mailmerge.Add("學生生日英文年", sr.Birthday.Value.Year);
                    mailmerge.Add("學生生日月", sr.Birthday.Value.Month);
                    mailmerge.Add("學生生日英文月", sr.Birthday.Value.ToString("MMMM", new System.Globalization.CultureInfo("en-US")));
                    mailmerge.Add("學生生日英文月3", sr.Birthday.Value.ToString("MMM", new System.Globalization.CultureInfo("en-US")));
                    mailmerge.Add("學生生日上標", daySuffix(sr.Birthday.Value.Day.ToString()));
                    mailmerge.Add("學生生日日", sr.Birthday.Value.Day);
                }
                if (dphoto_p.ContainsKey(sr.ID))
                {
                    mailmerge.Add("入學照片1吋", dphoto_p[sr.ID]);
                    mailmerge.Add("入學照片2吋", dphoto_p[sr.ID]);
                }
                if (dphoto_g.ContainsKey(sr.ID))
                {
                    mailmerge.Add("畢業照片1吋", dphoto_g[sr.ID]);
                    mailmerge.Add("畢業照片2吋", dphoto_g[sr.ID]);
                }
                //畢業資訊
                if (dshlir.ContainsKey(sr.ID))
                {
                    mailmerge["畢業資訊西元年"] = dshlir[sr.ID].SchoolYear + 1911;
                    mailmerge["畢業資訊學年度"] = dshlir[sr.ID].SchoolYear;
                    mailmerge["畢業資訊證書字號"] = dshlir[sr.ID].DiplomaNumber;
                    mailmerge["畢業資訊證書字號數字"] = getCertificateNumberNumber(dshlir[sr.ID].DiplomaNumber);
                    mailmerge["畢業資訊科別中文"] = dshlir[sr.ID].DepartmentName;
                    string tmp_dept = dshlir[sr.ID].DepartmentName;
                    mailmerge["畢業資訊科別英文"] = (tmp_dept != null && dic_dept_ch_en.ContainsKey(tmp_dept)) ? dic_dept_ch_en[tmp_dept] : "";

                }

                // 入學學年
                if (dsEntryRceDict.ContainsKey(sr.ID))
                {
                    DateTime dt;
                    if (DateTime.TryParse(dsEntryRceDict[sr.ID].UpdateDate, out dt))
                    {
                        if (dt.Year > 1911)
                            mailmerge["入學西元年"] = dt.Year;
                    }
                }

                //畢業異動
                if (dshurr.ContainsKey(sr.ID))
                {
                    //int ExpectGraduateSchoolYear;
                    //if (int.TryParse(dshurr[sr.ID].ExpectGraduateSchoolYear, out ExpectGraduateSchoolYear))
                    //    mailmerge["畢業異動西元年"] = ExpectGraduateSchoolYear + 1911;

                    if (!string.IsNullOrEmpty(dshurr[sr.ID].UpdateDate))
                    {
                        int ADYear;
                        int.TryParse(dshurr[sr.ID].UpdateDate.Split('/')[0], out ADYear);
                        mailmerge["畢業異動西元年"] = ADYear;
                        int republicYaer;
                        if (ADYear > 1911)
                        {
                            republicYaer = ADYear - 1911;
                            mailmerge["畢業異動民國年"] = republicYaer;
                        }

                        int UpdateDateMonth;
                        int.TryParse(dshurr[sr.ID].UpdateDate.Split('/')[1], out UpdateDateMonth);
                        int UpdateDateDate;
                        int.TryParse(dshurr[sr.ID].UpdateDate.Split('/')[2], out UpdateDateDate);

                        mailmerge["畢業異動月"] = UpdateDateMonth;
                        mailmerge["畢業異動日"] = UpdateDateDate;

                        mailmerge["畢業異動英文月"] = ToEngMonth(UpdateDateMonth);
                        mailmerge["畢業異動日上標"] = daySuffix(UpdateDateDate.ToString());
                    }


                    mailmerge["畢業異動學年度"] = dshurr[sr.ID].ExpectGraduateSchoolYear;
                    mailmerge["畢業異動證書字號"] = dshurr[sr.ID].GraduateCertificateNumber;
                    mailmerge["畢業異動證書字號數字"] = getCertificateNumberNumber(dshurr[sr.ID].GraduateCertificateNumber);
                    mailmerge["畢業異動科別中文"] = dshurr[sr.ID].Department;
                    string tmp_dept = dshurr[sr.ID].Department;
                    mailmerge["畢業異動科別英文"] = (tmp_dept != null && dic_dept_ch_en.ContainsKey(tmp_dept)) ? dic_dept_ch_en[tmp_dept] : "";
                }
                #endregion

                each = new Document(new MemoryStream(template));

                //each.MailMerge.CleanupOptions = Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveUnusedFields;
                each.MailMerge.FieldMergingCallback = new merge();
                each.MailMerge.Execute(mailmerge.Keys.ToArray(), mailmerge.Values.ToArray());
                each.MailMerge.DeleteFields();
                document.Sections.Add(document.ImportNode(each.FirstSection, true));
            }
            document.Sections.RemoveAt(0);
            e.Result = document;
        }
        class merge : Aspose.Words.Reporting.IFieldMergingCallback
        {
            public void FieldMerging(Aspose.Words.Reporting.FieldMergingArgs e)
            {

                if (e.FieldName == "入學照片1吋" || e.FieldName == "入學照片2吋")
                {
                    int tmp_width;
                    int tmp_height;
                    if (e.FieldName == "入學照片1吋")
                    {
                        tmp_width = 25;
                        tmp_height = 35;
                    }
                    else
                    {
                        tmp_width = 35;
                        tmp_height = 45;
                    }
                    #region 入學照片
                    if (e.FieldValue != null && !string.IsNullOrEmpty(e.FieldValue.ToString()))
                    {
                        byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                        if (photo != null && photo.Length > 0)
                        {
                            DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                            photoBuilder.MoveToField(e.Field, true);
                            e.Field.Remove();
                            //Paragraph paragraph = photoBuilder.InsertParagraph();// new Paragraph(e.Document);
                            Shape photoShape = new Shape(e.Document, ShapeType.Image);
                            photoShape.ImageData.ImageBytes = photo;
                            photoShape.WrapType = WrapType.Inline;
                            //Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                            //cell.CellFormat.LeftPadding = 0;
                            //cell.CellFormat.RightPadding = 0;

                            photoShape.Width = ConvertUtil.MillimeterToPoint(tmp_width);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(tmp_height);
                            //paragraph.AppendChild(photoShape);
                            photoBuilder.InsertNode(photoShape);
                        }
                    }
                    #endregion
                }
                else if (e.FieldName == "畢業照片1吋" || e.FieldName == "畢業照片2吋")
                {
                    int tmp_width;
                    int tmp_height;
                    if (e.FieldName == "畢業照片1吋")
                    {
                        tmp_width = 25;
                        tmp_height = 35;
                    }
                    else
                    {
                        tmp_width = 35;
                        tmp_height = 45;
                    }
                    #region 畢業照片
                    if (e.FieldValue != null && !string.IsNullOrEmpty(e.FieldValue.ToString()))
                    {
                        byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                        if (photo != null && photo.Length > 0)
                        {
                            DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                            photoBuilder.MoveToField(e.Field, true);
                            e.Field.Remove();
                            //Paragraph paragraph = photoBuilder.InsertParagraph();// new Paragraph(e.Document);
                            Shape photoShape = new Shape(e.Document, ShapeType.Image);
                            photoShape.ImageData.ImageBytes = photo;
                            photoShape.WrapType = WrapType.Inline;
                            //Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                            //cell.CellFormat.LeftPadding = 0;
                            //cell.CellFormat.RightPadding = 0;
                            photoShape.Width = ConvertUtil.MillimeterToPoint(tmp_width);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(tmp_height);
                            //paragraph.AppendChild(photoShape);
                            photoBuilder.InsertNode(photoShape);
                        }
                    }
                    #endregion
                }
            }

            public void ImageFieldMerging(Aspose.Words.Reporting.ImageFieldMergingArgs args)
            {
                throw new NotImplementedException();
            }
        }
        void _bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document inResult = (Document)e.Result;
            btnPrint.Enabled = true;
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Word (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = current;

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    inResult.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(SaveFileDialog1.FileName + ",列印完成!!");
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                    return;
                }
            }
            catch
            {
                string msg = "檔案儲存錯誤,請檢查檔案是否開啟中!!";
                FISCA.Presentation.Controls.MsgBox.Show(msg);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(msg);
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            if (value == "新增") return;
            //畫面內容(範本內容,預設樣式
            Campus.Report2014.TemplateSettingForm TemplateForm;
            if (custConfigs[current].Template == null)
            {
                custConfigs[current].Template = new Campus.Report2014.ReportTemplate(Properties.Resources.證明書範本, Campus.Report2014.TemplateType.docx);
            }
            Campus.Report2014.ReportTemplate defaultDoc;
            switch (value)
            {
                case "畢業證書_普通高中":
                    defaultDoc = new Campus.Report2014.ReportTemplate(Properties.Resources.畢業證書_普通高中, Campus.Report2014.TemplateType.docx);
                    break;
                case "畢業證書_高職":
                    defaultDoc = new Campus.Report2014.ReportTemplate(Properties.Resources.畢業證書_高職, Campus.Report2014.TemplateType.docx);
                    break;
                case "補發證明書_普通高中":
                    defaultDoc = new Campus.Report2014.ReportTemplate(Properties.Resources.補發證明書_普通高中, Campus.Report2014.TemplateType.docx);
                    break;
                case "補發證明書_高職":
                    defaultDoc = new Campus.Report2014.ReportTemplate(Properties.Resources.補發證明書_高職, Campus.Report2014.TemplateType.docx);
                    break;
                default:
                    defaultDoc = new Campus.Report2014.ReportTemplate(Properties.Resources.證明書範本, Campus.Report2014.TemplateType.docx);
                    break;
            }
            TemplateForm = new Campus.Report2014.TemplateSettingForm(custConfigs[current].Template, defaultDoc);
            //預設名稱
            TemplateForm.DefaultFileName = current + "樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                custConfigs[current].Template = TemplateForm.Template;
                custConfigs[current].Save();
            }
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            if (value == "新增") return;
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count < 1)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請先選擇學生");
                return;
            }
            btnPrint.Enabled = false;
            _bgw.RunWorkerAsync();
        }
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var value = (string)comboBoxEx1.SelectedItem;
            switch (value)
            {
                case "新增":
                    //第一次使用時加入
                    if (custConfigs.Count == 0)
                    {
                        addCustConfig("畢業證書_普通高中");
                        ReportConfiguration custConf;
                        custConf = new Campus.Report2014.ReportConfiguration(configNameRule("畢業證書_普通高中"));
                        custConf.Template = new Campus.Report2014.ReportTemplate(Properties.Resources.畢業證書_普通高中, Campus.Report2014.TemplateType.docx);
                        custConf.Save();
                        custConfigs.Add("畢業證書_普通高中", custConf);
                        comboBoxEx1.Items.Insert(0, "畢業證書_普通高中");

                        addCustConfig("畢業證書_高職");
                        custConf = new Campus.Report2014.ReportConfiguration(configNameRule("畢業證書_高職"));
                        custConf.Template = new Campus.Report2014.ReportTemplate(Properties.Resources.畢業證書_高職, Campus.Report2014.TemplateType.docx);
                        custConf.Save();
                        custConfigs.Add("畢業證書_高職", custConf);
                        comboBoxEx1.Items.Insert(0, "畢業證書_高職");

                        addCustConfig("補發證明書_普通高中");
                        custConf = new Campus.Report2014.ReportConfiguration(configNameRule("補發證明書_普通高中"));
                        custConf.Template = new Campus.Report2014.ReportTemplate(Properties.Resources.補發證明書_普通高中, Campus.Report2014.TemplateType.docx);
                        custConf.Save();
                        custConfigs.Add("補發證明書_普通高中", custConf);
                        comboBoxEx1.Items.Insert(0, "補發證明書_普通高中");

                        addCustConfig("補發證明書_高職");
                        custConf = new Campus.Report2014.ReportConfiguration(configNameRule("補發證明書_高職"));
                        custConf.Template = new Campus.Report2014.ReportTemplate(Properties.Resources.補發證明書_高職, Campus.Report2014.TemplateType.docx);
                        custConf.Save();
                        custConfigs.Add("補發證明書_高職", custConf);
                        comboBoxEx1.Items.Insert(0, "補發證明書_高職");

                        conf.Save();
                        comboBoxEx1.SelectedIndex = 0;
                    }
                    else
                    {
                        AddNew input = new AddNew();
                        if (input.ShowDialog() == DialogResult.OK)
                        {
                            input.name = System.Text.RegularExpressions.Regex.Replace(input.name, @"[\W]+", "");
                            if (string.IsNullOrWhiteSpace(input.name))
                                FISCA.Presentation.Controls.MsgBox.Show("請輸入樣板名稱(中文或英文字母)");
                            else if (custConfigs.ContainsKey(input.name))
                                FISCA.Presentation.Controls.MsgBox.Show("樣板名稱已存在");
                            else
                            {
                                ReportConfiguration tmp_conf = new ReportConfiguration(configNameRule(input.name));
                                if (input.Template != null)
                                    tmp_conf.Template = new ReportTemplate(input.Template, TemplateType.docx);
                                tmp_conf.Save();
                                custConfigs.Add(input.name, tmp_conf);
                                addCustConfig(input.name);
                                comboBoxEx1.Items.Insert(0, input.name);
                                comboBoxEx1.SelectedIndex = 0;
                            }
                        }
                    }
                    break;
                default:
                    current = value;
                    break;
            }
        }
        private void delete_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            switch (value)
            {
                case "新增":
                    break;
                default:
                    if (custConfigs.ContainsKey(value))
                    {
                        custConfigs[value].Template = null;
                        custConfigs[value].Save();
                        custConfigs.Remove(value);
                        comboBoxEx1.Items.Remove(value);
                        delCustConfig(value);
                    }
                    break;
            }
            comboBoxEx1.SelectedIndex = 0;
            current = (string)comboBoxEx1.SelectedItem;
        }
        private void addCustConfig(string custConfig)
        {
            List<string> tmp = conf.GetString("customs", "").Split(';').ToList<string>();
            tmp.Add(System.Text.RegularExpressions.Regex.Replace(custConfig, @"[\W]+", ""));
            conf.SetString("customs", string.Join(";", tmp));
            conf.Save();
        }
        private void delCustConfig(string custConfig)
        {
            List<string> tmp = conf.GetString("customs", "").Split(';').ToList<string>();
            tmp.Remove(custConfig);
            tmp.Remove(custConfig);
            tmp.Remove(custConfig);
            tmp.Remove(custConfig);
            conf.SetString("customs", string.Join(";", tmp));
            conf.Save();
        }
        private string[] getCustConfig()
        {
            return conf.GetString("customs", "").Split(';');
        }
        private static string configNameRule(string custConfigName)
        {
            return config + "." + custConfigName;
        }
        public static string daySuffix(string date)
        {
            //switch (int.Parse(date) % 10)
            //{
            //    case 1: return "st";
            //    case 2: return "nd";
            //    case 3: return "rd";
            //    default: return "th";
            //}

            //2022-05-25 Cynthia 11、12、13應是th
            switch (int.Parse(date))
            {
                case 1: return "st";
                case 2: return "nd";
                case 3: return "rd";

                case 21: return "st";
                case 22: return "nd";
                case 23: return "rd";

                case 31: return "st";

                default: return "th";
            }
        }

        public static string ToEngMonth(int month)
        {
            switch (month)
            {
                case 1: return "January";
                case 2: return "February";
                case 3: return "March";
                case 4: return "April";
                case 5: return "May";
                case 6: return "June";
                case 7: return "July";
                case 8: return "August";
                case 9: return "September";
                case 10: return "October";
                case 11: return "November";
                case 12: return "December";

                default: return "";
            }
        }
        public static string getCertificateNumberNumber(string cn)
        {
            if (cn != null)
            {
                MatchCollection matches = Regex.Matches(cn, @"\d+");
                if (matches.Count > 0)
                    return matches[matches.Count - 1].Value;
            }
            return cn;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
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
                SaveFileDialog1.Filter = "Word (*.docx)|*.docx|所有檔案 (*.*)|*.*";
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
