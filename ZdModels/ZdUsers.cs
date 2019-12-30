using System.Collections.Generic;

namespace VerifyZdUserEmailAddresses.ZdModels
{
    public class ZdUsers
    {
        public List<ZdUser> results { get; set; }
        public object facets { get; set; }
        public object next_page { get; set; }
        public object previous_page { get; set; }
        public int count { get; set; }
    }
}
