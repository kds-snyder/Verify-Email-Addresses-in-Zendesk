using System;

namespace VerifyZdUserEmailAddresses.ZdModels
{
    public class ZdUserIdentity
    {
        public string url { get; set; }
        public long id { get; set; }
        public long user_id { get; set; }
        public string type { get; set; }
        public string value { get; set; }
        public bool verified { get; set; }
        public bool primary { get; set; }
        public DateTime created_at { get; set; }
        public DateTime updated_at { get; set; }
        public int undeliverable_count { get; set; }
        public string deliverable_state { get; set; }
    }
}
