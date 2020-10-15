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
 * Author notes:
 * Google JSON API key:  AIzaSyAgk3eXP0oE1gpdhfrMLHAnuQ4UYoJwrmE 
 * Google search engine: 7f90a3ed8b3499910
 * https://customsearch.googleapis.com/customsearch/v1?cx=7f90a3ed8b3499910&fileType=.jpg&gl=us&lr=lang_en&num=10&q=square&safe=active&searchType=image&key=[YOUR_API_KEY]
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
using System.Net;
using System.IO;
using Google.Apis.Customsearch.v1.Data;
using Newtonsoft.Json;

namespace PowerPoint_Maker
{
    public partial class Form1 : Form
    {
        //variables and constants related to the form.
        private string title;
        private string textArea;
        private const string API_KEY = "key=AIzaSyAgk3eXP0oE1gpdhfrMLHAnuQ4UYoJwrmE";
        private const string SEARCH_ENGINE = "cx=7f90a3ed8b3499910";
        private const string API = "https://customsearch.googleapis.com/customsearch/v1?fileType=.jpg&gl=us&lr=lang_en&num=10&safe=active&searchType=image&" + API_KEY + "&" + SEARCH_ENGINE;

        //Auto generated.
        public Form1()
        {
            InitializeComponent();
        }

        //If the user sets a new title, then start an image search, and repopulate the images shown. 10/15/2020 - JB
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            title = textBox1.Text;
            GoogleImageSearch(title);
        }

        //If given time, I will develope a full WordPad like editor for this richTextBox,
        //however, if there is not enough time, then the user should use powerpoint to create
        //their text area, copy then paste that into the Text Area richTextBox. 10/14/2020 - JB
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            textArea = richTextBox1.Text;
            GoogleImageSearch(title);
        }

        //This button will generate a .ppt file with the slide using the title and text area.
        private void button1_Click(object sender, EventArgs e)
        {
            /*https://www.free-power-point-templates.com/articles/create-powerpoint-ppt-programmatically-using-c/
            * The above link is the reference where I grabbed the following code related to generating a .ppt file.
            * 10/15/2020 - JB
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

        //Search Google Images using a passed in String, updating the listView. 10/14/2020 - JB
        //using customsearch API from Google: the query is "GoogleSearchString", and the API key, custom search engine, and url is listed above. 10/15/2020 - JB
        private void GoogleImageSearch(string GoogleSearchString)
        {
            //Check if input is not empty. 10/14/2020
            if(GoogleSearchString != "")
            {

            }

            //Complete the API url with the search query. 10/15/2020 - JB
            string apiURL = API + "&q=" + GoogleSearchString;

            //https://www.youtube.com/watch?v=3NLfhEjq9Tk
            //the above link is a reference on the following code. I am struggling to figure out
            //how to get the JSON data using the API above.
            var request = WebRequest.Create(apiURL);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string responseString = reader.ReadToEnd();
            /*
             * dynamic jsonData = Json.Convert.DeserializeObject(reponseString);
            var results = new List<Result>();
            foreach(var item in jsonData.items)
            {
                results.Add(new Result {
                    Title = item.title,
                    Link = item.link,
                    Snippet = item.snippet,
                }) ;

            }
           return View(results.ToList());
            */

        }
    }
}
