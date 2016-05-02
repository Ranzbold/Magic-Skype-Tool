using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Magic_Skype_Tool
{
    class QuoteUtils
    {
        public static DataObject getObject(String name, String message, String timestring)
        {
            dynamic ms = Convert.ToInt64(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds);
            DateTime time = DateTime.Parse(timestring);
            string str = DateTime.Now.ToString("u");
            DataObject data = new DataObject();
            data.SetData("System.String", str);
            data.SetData("UnicodeText", str);
            data.SetData("Text", str);
            data.SetData("SkypeMessageFragment", new System.IO.MemoryStream(Encoding.UTF8.GetBytes(string.Concat(new object[] {
    "<quote author=\"",
    name,
    " \" timestamp=\"",
    ms,
    "\">",
    message,
    "</quote>"
}))));
            data.SetData("Locale");
            data.SetData("OEMText", str);


            return data;
        }
    }
}
