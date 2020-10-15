/* Author: Jonathan Brown
 * Date: October 15th, 2020
 * Directions:
 * 
 * Create a solution that a user can use to aid them in generating a power 
 * point slide. They want the solution to give them suggestions of images 
 * to use form the internet (Google) based on the contents of the 
 * information they are using for the slide. They want to improve 
 * efficiency and save time not having to search for images for every 
 * slide they are making for their presentations.
 * 
 * Create a solution that accepts user input
 * 
 *     Title area (input)
 *     Text area (input)
 *     image suggestion area (multiple selection); utilize words in the 
 *          title, and bold words in the text area to bring suggested images in, 
 *          with ability to select multiple images to be included in the slide 
 * 
 *  
 * 
 * Final solution should use title, text areas, and using selected images to 
 * build power point slide.
 * Power point slide(s) should be the output.
 * 
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace PowerPoint_Maker
{
    public partial class Form1 : Form
    {
        private string title;
        private string textArea;

        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            title = textBox1.Text;
            GoogleImageSearch(title);
        }

        //If given time, I will develope a full WordPad like editor for this richTextBox,
        //however, if there is not enough time, then the user should use powerpoint to create
        //their text area, copy then paste that into the Text Area richTextBox. 10/14/2020
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            textArea = richTextBox1.Text;
            GoogleImageSearch(title);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*https://www.free-power-point-templates.com/articles/create-powerpoint-ppt-programmatically-using-c/
            * The above link is used for reference in this section of the applicaiton
            * 10/15/2020
            */
            PowerPoint.Application pptApplication = new PowerPoint.Application();

            Microsoft.Office.Interop.PowerPoint.Slides slides;
            Microsoft.Office.Interop.PowerPoint._Slide slide;
            Microsoft.Office.Interop.PowerPoint.TextRange objText;

            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Create new Slide
            slides = pptPresentation.Slides;
            slide = slides.AddSlide(1, customLayout);

            // Add title
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = title;
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;

            objText = slide.Shapes[2].TextFrame.TextRange;
            objText.Text = textArea;

            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "test";

            pptPresentation.SaveAs(@"c:\temp\SEH_test.pptx", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation.Close();
            pptApplication.Quit();
        }

        //Search Google Images using a passed in String, updating the listView. 10/14/2020
        private void GoogleImageSearch(String GoogleSearchString)
        {
            //Check if input is good. 10/14/2020
            if(GoogleSearchString != "")
            {

            }
        }
    }
}
