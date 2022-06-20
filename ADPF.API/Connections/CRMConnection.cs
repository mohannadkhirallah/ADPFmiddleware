using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ADPF.API.Connections
{
    public static class CRMConnection
    {
        static IOrganizationService _ADPFCRM;
        static IOrganizationService mutkamlahClientService;
        static DateTime serviceExpiryTime;
        static DateTime _serviceProxyCreatedTime;

        /// <summary>
        /// Create Static // Singlton Object for CRM connection 
        /// </summary>
        public static IOrganizationService ADPFCRM
        {
            get { return _ADPFCRM; }
        }

    }
}