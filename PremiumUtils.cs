using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Magic_Skype_Tool
{
    class PremiumUtils
    {
        public static Boolean isPremium(String user)
        {

            string adress = "http://holger-reuter.de/skypepremium/users.txt";
            WebRequest request = WebRequest.Create(adress);
            WebResponse response = request.GetResponse();

            List<string> users = new List<string>();


            using (StreamReader sr = new StreamReader((response.GetResponseStream())))
            {
                string line;
                int counter = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    users.Add(line);
                    counter++;
                }
            }

            foreach (String u in users)
            {
                if (u.Equals(CryptoUtils.SHA512(user)))
                {
                    return true;
                }
            }
            return false;

        }
    }
}
