using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using System.Linq;

namespace Bizarre.PowiadomieniaTimerJob.Features.PowiadomieniaTimerJobFeature
{
    [Guid("9e5c11f8-11ee-423a-968b-d1b2a7831dbb")]
    public class PowiadomieniaTimerJobFeatureEventReceiver : SPFeatureReceiver
    {

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            var site = properties.Feature.Parent as SPSite;
            PowiadomieniaTimerJob.CreateTimerJob(site);
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            var site = properties.Feature.Parent as SPSite;
            PowiadomieniaTimerJob.DelteTimerJob(site);
        }
    }
}
