using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Magic_Skype_Tool
{
    class PerfUtils
    {
        public static int getCpuUsage() {
            System.Diagnostics.PerformanceCounter cpuCounter;
            cpuCounter = new System.Diagnostics.PerformanceCounter();
            cpuCounter.CategoryName = "Processor";
            cpuCounter.CounterName = "% Processor Time";
            cpuCounter.InstanceName = "_Total";
            float percfloat = cpuCounter.NextValue();
            int percentage = (int)percfloat;
            return percentage;
        }


  

    }
}
