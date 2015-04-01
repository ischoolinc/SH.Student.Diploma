using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SH.Student.Diploma
{
    class Permissions
    {
        public static string 學籍相關證明書 { get { return "plugins.student.report.certificate.2.cs"; } }
        public static bool 學籍相關證明書權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學籍相關證明書].Executable;
            }
        }
    }
}
