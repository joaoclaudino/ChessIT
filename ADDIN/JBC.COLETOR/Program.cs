using Castle.Core.Logging;
using JBC.Framework.Attribute;
using JBC.Framework.DAO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JBC.COLETOR
{
    [ResourceBOM("JBC.COLETOR.resources.BOM.UDF.xml", ResourceType.UserField)]
    [AddIn(
        Name = "JBC.COLETOR"
        , Description = "JBC.COLETOR - Claudino Soluções em Software"
        , Namespace = "JBC"
        , LicenseFile = "JBC.COLETOR.publicKey.xml"
        )]
    class Program
    {
        private SAPbobsCOM.Company Sapcompany { get; set; }
        private SAPbouiCOM.Application sapApp { get; set; }
        private BusinessOneDAO JBCDAO { get; set; }
        private ILogger Logger { get; set; }

        public Program()
        {

        }
    }
}
