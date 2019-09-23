using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Aspose.Email;
using Aspose.Email.Mapi;

namespace ReplaceHyperlinkInMSG
{
    class Program
    {
        public static void ReadAndReplaceMessageBody()
        {
           // String path = @"C:\Email\";
            MapiMessage msg = MapiMessage.FromFile("Test Links.msg");

            String URLToReplace = "www.aspose.com";
            if (msg.BodyType == BodyContentType.Html)
            {
                //Extract HTML Body
                var bodyHtml = msg.BodyHtml;

                //Setting Regex for finding hyperlink
                Regex rs = new Regex(@"<a.*?href=(""|')(?<href>.*?)(""|').*?>(?<value>.*?)</a>");

                foreach (Match match in rs.Matches(bodyHtml))
                {
                    //Extrct URL
                    string url = match.Groups["href"].Value;

                    //Replace URL
                    bodyHtml = bodyHtml.Replace(url, URLToReplace);
                }

                //Setting new body
                msg.Body = bodyHtml;
            }

            //Save MSG
            Aspose.Email.SaveOptions options = new MsgSaveOptions(MailMessageSaveType.OutlookMessageFormatUnicode);
            msg.Save(@"Test Links_SavedMsg.msg", options);
        }

        public static void LoadLicense()
        {
            String licPath = @"C:\Aspose Data\Licenses_2019\";
            Aspose.Email.License lic = new Aspose.Email.License();
            lic.SetLicense(licPath + "Aspose.Total.Net.lic");
        }

        static void Main(string[] args)
        {
            //LoadLicense();
            ReadAndReplaceMessageBody();
        }
    }
}
