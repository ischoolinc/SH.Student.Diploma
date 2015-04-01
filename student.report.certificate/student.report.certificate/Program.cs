using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;

namespace SH.Student.Diploma
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            MenuButton item = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["學籍相關報表"];
            item["學籍相關證明書"].Enable = Permissions.學籍相關證明書權限;
            item["學籍相關證明書"].Click += delegate
            {
                new MainForm().ShowDialog();
            };
            Catalog detail1 = RoleAclSource.Instance["學生"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.學籍相關證明書, "學籍相關證明書"));
        }
    }
}
