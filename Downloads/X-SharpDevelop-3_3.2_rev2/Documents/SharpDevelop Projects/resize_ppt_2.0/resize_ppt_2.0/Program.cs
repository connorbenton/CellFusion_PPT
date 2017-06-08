 /*
 * Created by SharpDevelop.
 * User: cbenton
 * Date: 5/26/2017
 * Time: 3:28 PM
 */

using System;
using System.Windows.Forms;
using System.Linq;
using System.Threading;

namespace resize_ppt_2._
{
			
	static class Program
	{
		public static SplashScreen splashscreen = null;
		
		public static MainForm mainForm = null;
		
		// Program entry point.
		[STAThread]
		public static void Main()
		{
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
 
        //show splash
        Thread splashThread = new Thread(new ThreadStart(
            delegate
            {
                splashscreen = new SplashScreen();
                Application.Run(splashscreen);
            }
            ));
        splashThread.SetApartmentState(ApartmentState.STA);
        splashThread.Start();
 
        //run form - time taking operation
        mainForm = new MainForm();
        mainForm.Load += new EventHandler(mainForm_Load);
        mainForm.Shown += new EventHandler(mainForm_Shown);
        mainForm.Activate();
        Application.Run(mainForm);
		}
		
	    static void mainForm_Load(object sender, EventArgs e)
	    {
	        //close splash
	        if (splashscreen == null)
	        {
	            return;
	        }
	        
	        splashscreen.Invoke(new Action(splashscreen.Close));
	        splashscreen.Dispose();
	        splashscreen = null;
	        
	    }
	    
	    static void mainForm_Shown(object sender, EventArgs e)
	    {
	    	mainForm.WindowState = FormWindowState.Minimized;
	    	mainForm.Show();
	    	mainForm.WindowState = FormWindowState.Normal;
	    }
	}
	
}
		