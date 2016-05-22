using SKYPE4COMLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magic_Skype_Tool
{
    class SkypeUtils
    {
        public static void sendMessages(Skype skype,List<string> kontakte, string message, int anzahl)
        {
            for (int a = 0; a <= anzahl; a++)
            {
                foreach (string s in kontakte)
                {
                    skype.SendMessage(s, message);

                }

            }

        }


    }
}
