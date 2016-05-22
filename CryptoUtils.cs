using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Magic_Skype_Tool
{
    class CryptoUtils
    {
        public static string SHA512(string s)
        {
            SHA512 shaM = new SHA512Managed();
            byte[] hash =
             shaM.ComputeHash(Encoding.ASCII.GetBytes(s));

            StringBuilder stringBuilder = new StringBuilder();
            foreach (byte b in hash)
            {
                stringBuilder.AppendFormat("{0:x2}", b);
            }
            return stringBuilder.ToString();
        }
    }
}
