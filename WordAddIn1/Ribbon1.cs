using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;


namespace WordAddIn1
{
    public partial class Ribbon1
    {

        public static Word.Application e_application; 

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            e_application = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show( ReadDocumentProperty("transducerSN1"), "transducerSN1",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Information);
            System.Windows.Forms.MessageBox.Show( ReadDocumentCustomProperties(), "Document properties",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Information);
        }

        private string ReadDocumentCustomProperty(string propertyName)
        {
            DocumentProperties properties;
            properties = (DocumentProperties)e_application.ActiveDocument.CustomDocumentProperties;

            foreach (DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }

        private string ReadDocumentCustomProperties()
        {
            DocumentProperties properties;
            string props = null;
            properties = (DocumentProperties)e_application.ActiveDocument.CustomDocumentProperties;

            foreach (DocumentProperty prop in properties)
            {
                props += prop.Name + ": " + prop.Value.ToString()+ "\n";      
            }
            return props;
        }

        private string ReadDocumentProperties()
        {
            DocumentProperties properties;
            string props = null;
            properties = (DocumentProperties)e_application.ActiveDocument.BuiltInDocumentProperties;
            string temp;
            foreach (DocumentProperty prop in properties)
            {
                try
                {
                    if (prop != null & prop.Value != null)
                        temp = prop.Value.ToString();
                    else
                        temp = "";
                    props += prop.Name + ": " + temp + "\n";
                }
                catch(Exception ex)
                {

                }
            }
            return props;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(ReadDocumentProperties(), "Document properties", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
