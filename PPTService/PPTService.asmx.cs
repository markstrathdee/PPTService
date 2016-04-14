using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Web;
using System.Web.Services;
using System.Web.Script.Serialization;
using System.Web.Script.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.IO;
using System.Web.WebSockets;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PPTService
{
    /// <summary>
    /// Summary description for PPTService
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class PPTService : System.Web.Services.WebService
    {

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }

        [WebMethod]

        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]

        public void createPPT()
        {
            var app = new Microsoft.Office.Interop.PowerPoint.Application();

            var pres = app.Presentations;

            //var file = pres.Open(@"C:\presentation1.ppt", MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
            var file = pres.Open(@"C:\presentation1.ppt");

            file.SaveCopyAs(@"C:\Users\Mark C Strathdee\Desktop\presentation1.jpg", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsJPG, MsoTriState.msoTrue);

        }

    }
}
