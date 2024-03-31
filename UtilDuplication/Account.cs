using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UtilDuplication
{
    public class Account
    {
        public string Name { get; set; }
        public string AccountId { get; set; }
        public string A2C_ZoomInfoAccountId { get; set; }
        public string A2C_AccountSalesStage { get; set; }
        public string A2C_AccountType { get; set; }
        public string A2C_EmployeeBand { get; set; }
        public string A2C_Domain { get; set; }
        public string Owner { get; set; }
        public Guid OwnerId { get; set; }
        public string Country { get; set; }
        public string Lastmodified { get; set; }
        public string AccountPriority { get; set; }
        public string AccountPotential { get; set; }
        public string PrimaryContact { get; set; }
        public Guid PrimaryContactId { get; set; }
        public string ScheduledFollowup { get; set; }
        public string ParentChild { get; set; }
    }

}
