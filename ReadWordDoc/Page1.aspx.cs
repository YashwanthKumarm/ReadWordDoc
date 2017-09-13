using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ReadWordDoc
{
    public partial class Page1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            WordFileToRead.SaveAs(Server.MapPath("/Repository/" + WordFileToRead.FileName));
            object filename = Server.MapPath("/Repository/" + WordFileToRead.FileName);
            Application AC = new Application();
            Document doc = new Document();
            object readOnly = false;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;
            try
            {
                doc = AC.Documents.Open(ref filename, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref isVisible, ref missing, ref missing, ref missing);
                string fileContent = doc.Content.Text;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}