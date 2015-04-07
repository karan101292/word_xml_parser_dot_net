using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections;

namespace Spike2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {

            XDocument xDoc = XDocument.Load("2.xml");
            XNamespace w = "http://schemas.microsoft.com/office/word/2003/wordml";

            

            // Find all paragraphs in the document.
            var paragraphs =
                from para in xDoc
                             .Root
                             .Element(w + "body")
                             .Descendants(w + "p")
                let styleNode = para
                                .Elements(w + "pPr")
                                .Elements(w + "pStyle")
                                .FirstOrDefault()
                select new
                {
                    ParagraphNode = para,
                    StyleName = styleNode 
                };

            // Following is the new query that retrieves the text of
            // each paragraph.
            var paraWithText =
                from para in paragraphs
                select new
                {
                    ParagraphNode = para.ParagraphNode,
                    StyleName = para.StyleName,
                    Text = para
                           .ParagraphNode
                           .Elements(w + "r")
                           .Elements(w + "t")
                           .Aggregate(
                               new StringBuilder(),
                               (s, i) => s.Append((string)i),
                               s => s.ToString()
                           )
                };
            Hashtable elements = new Hashtable();
            String[] authors = {};
            String[] affiliations = {};
            String title="";
            DateTime start_time, end_time;

            //start_time = DateTime.Now;
            foreach (var p in paraWithText){
                if (p.StyleName.Attribute(w + "val").Value.ToString().Equals("Author_28_s_29_"))
                {
                    authors = p.Text.ToString().Split(',');
                }
                if (p.StyleName.Attribute(w + "val").Value.ToString().Equals("Affiliation"))
                {
                    affiliations = p.Text.ToString().Split(',');
                }
                if (p.StyleName.Attribute(w + "val").Value.ToString().Equals("P5"))
                {
                    title = p.Text.ToString();
                }

            }
            //end_time = DateTime.Now;

            //textBox2.Clear();
            //textBox2.AppendText("Element count \t :: \t" + elements.Count + Environment.NewLine);
            //textBox2.AppendText("Start Time \t :: \t" + start_time.TimeOfDay + Environment.NewLine);
            //textBox2.AppendText("End Time \t :: \t" + end_time.TimeOfDay + Environment.NewLine);
            //textBox2.AppendText("Time Taken \t :: \t" + end_time.Subtract(start_time).ToString());


            

            XDocument doc = new XDocument();
            XElement root = new XElement("document");
            root.Add(new XElement("title", title));
            root.Add(new XElement("authors", authors.Select(x => new XElement("author", x))));
            root.Add(new XElement("affiliations", affiliations.Select(x => new XElement("affiliation", x))));

            textBox1.Text = root.ToString();
        
                              


                        
   
        }
    }
}


