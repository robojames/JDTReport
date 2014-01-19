using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDTReport
{
    class Job
    {
        int job_ID { get; set; }
        
        string Engineer { get; set; }

        DateTime TestEnd { get; set; }

        DateTime ReportWrite { get; set; }
    }

}
