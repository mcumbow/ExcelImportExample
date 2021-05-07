using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelExample.Models
{
    public class StaffInfoViewModel
    {
        public string Name { get; set; }
        public string Roll { get; set; }
        public string Email { get; set; }
        public List<StaffInfoViewModel> StaffList { get; set; }

        public StaffInfoViewModel()
        {
            StaffList = new List<StaffInfoViewModel>();
        }
    }
}
