/*
 * Created by SharpDevelop.
 * User: cbenton
 * Date: 6/8/2017
 * Time: 12:22 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Microsoft.Office.Core;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
//using System.Runtime.InteropServices; //Not needed in C# 5.0
using System.IO;
using System.Linq;
using System.Collections.Generic;
//using System.Threading; //Not needed in C# 5.0


namespace CellFusion_PPT
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	/// 
	

	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();			

				
//			System.Threading.Timer
//			Thread.Sleep(500);
			
//			this.TopMost = false;
			
//			Activate();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void TextBox1TextChanged(object sender, EventArgs e)
		{
	
		}
		void Button1Click(object sender, EventArgs e)
		{
			float percentofTotalSlide;
			bool res = float.TryParse(textBox1.Text, out percentofTotalSlide);
			string pptPresPath;
			
			if (res == false || (percentofTotalSlide <= 0.0F) || (percentofTotalSlide > 1.1F))
			    {
			    	MessageBox.Show("You appear to have an invalid value entered for scale. Please enter a value between 0 and 1.1.");
				    return;
			    }
			
				//Pick presentation for resizing.
				OpenFileDialog openFileDialog1 = new OpenFileDialog();
				openFileDialog1.Filter = "PowerPoint Files|*.pptx;*.ppt;*.pptm";
	//			openFileDialog1.Filter = "Text Files|*.txt"; //Debugging filter
				openFileDialog1.Title = "Select the presentation for resizing";
				
				if (openFileDialog1.ShowDialog() == DialogResult.OK)
			    	pptPresPath = openFileDialog1.FileName;
			    else
			    {
			    	pptPresPath = string.Empty;
			    	return;
			    }
			    
			    PowerPoint.Application objApp; 
				PowerPoint.Presentations objPresSet;
				PowerPoint._Presentation objPres;
				PowerPoint._Presentation objPresNew;
				PowerPoint.Slides objSlides;
//				PowerPoint._Slide objSlide;
//				PowerPoint.TextRange objTextRng;
//				PowerPoint.Shapes objShapes;
//				PowerPoint.Shape objShape;
//				PowerPoint.SlideShowWindows objSSWs;
//				PowerPoint.SlideShowTransition objSST;
//				PowerPoint.SlideShowSettings objSSS;
//				PowerPoint.SlideRange objSldRng;
//				PowerPoint.ShapeRange objShpRng;

				string pptPresName;
				string pptCorrectedPresPath;
//				string objPresName; //don't need, is covered under string parsing
			    
			    pptPresName = Path.GetFileNameWithoutExtension(pptPresPath);
			    
			    pptCorrectedPresPath = Path.GetDirectoryName(pptPresPath) + "\\" + pptPresName + "_corrected" + Path.GetExtension(pptPresPath);
			    
			    if (pptPresName.Contains("_corrected"))
			    {
			    	DialogResult dialogResult1 = MessageBox.Show("It looks like you selected a corrected version of this file - you selected:" +
			    	    Environment.NewLine + Environment.NewLine + pptPresName + Environment.NewLine + Environment.NewLine +
			    	    "Are you sure you want to proceed (this will create a _corrected_corrected version)?", "Correct corrected file", 
			    	    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2,(MessageBoxOptions)0x40000);
			    	
			    	if(dialogResult1 == DialogResult.No)
				    	return;
			    }
			    
			    if (File.Exists(pptCorrectedPresPath))
			    {
			    	DialogResult dialogResult1 = MessageBox.Show("A corrected version of this file appears to exist already in this folder." +
			    	    " Do you want to proceed (this will overwrite existing corrected version)?", "Overwrite corrected file",
			    	    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2,(MessageBoxOptions)0x40000);
				    if(dialogResult1 == DialogResult.No)
				    	return;
			    }
			    
				//Open selected presentation.
				objApp = new PowerPoint.Application();
				objApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
				objPresSet = objApp.Presentations;
				//clear the original version if one is already open
				IsOpen_Close(pptPresName, objApp, false);
				objPres = objPresSet.Open(pptPresPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
				objSlides = objPres.Slides;
				
	//			objPresName = objPres.Name.ToString();
				
	
				
				//clear the fixed version if one is already open
				IsOpen_Close(pptPresName + "_corrected", objApp, true);
				
	//			try{
				
				//Initial notification window
				NativeWindow pptWindow = new NativeWindow();
				pptWindow.AssignHandle(new IntPtr(objApp.HWND));
		
				NotificationForm resizeDialog = new NotificationForm();
				resizeDialog.StartPosition = FormStartPosition.CenterParent;
				resizeDialog.Text = "Resizing Shapes";
				resizeDialog.label1.AutoSize = true;
				resizeDialog.label1.Location = new System.Drawing.Point(13, 13);
				resizeDialog.label1.Text = "Resizing presentation; please do not close PowerPoint or any presentations until the 'Resizing Done' message box appears." +
					Environment.NewLine + Environment.NewLine + "Please close this box to continue.";
				resizeDialog.Controls.Add(resizeDialog.label1);
				resizeDialog.AutoSize = true;
				resizeDialog.ShowDialog(pptWindow);
				
				//Add new presentation for resize result
				objPresNew = objPresSet.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
				objPresNew.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
				
				//Copy and paste old small presentation slides into new size presentation
	//			var seq = Enumerable.Range(1,objPres.Slides.Count).ToArray(); //old method all slides at once
	//			objPres.Slides.Range(seq).Copy();
				int i = 1;
				while (i <= objPres.Slides.Count) {
					objPresNew.Slides.Add(objPresNew.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
					objPres.Slides[i].Shapes.Range(Type.Missing).Copy();
					objPresNew.Slides[i].Select();
					objPresNew.Slides[i].Application.CommandBars.ExecuteMso("PasteSourceFormatting");
		//			objPresNew.Slides[1].Select();
		//			objPresNew.Slides[1].Application.CommandBars.ExecuteMso("PasteSourceFormatting");
					Application.DoEvents();
					i = i + 1;
//					PowerPoint.ShapeRange shpRng;
					bool deleteflag;
					int tablecount = 0;
					objApp.ActiveWindow.Selection.Unselect();
					PowerPoint.Slide currentSlide = objPresNew.Slides[i-1];
					foreach (PowerPoint.Shape f in currentSlide.Shapes) {
						deleteflag = false;
						if (f.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue) {
							tablecount = tablecount + 1;
							foreach (PowerPoint.Cell colheader in f.Table.Rows[1].Cells) {
	//							Console.WriteLine(colheader.Shape.TextFrame.TextRange.Text.ToString()); //Used for debugging tech signoff delete
	//							Console.WriteLine(colheader.Shape.TextFrame.TextRange.Text.ToString().Contains("Technician Sign"));
								if (colheader.Shape.TextFrame.TextRange.Text.ToString().Contains("Technician Sign")) {
									deleteflag = true;
								}
							}
						}
						else if (f.Type != MsoShapeType.msoTable){
							f.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
						}
						if (deleteflag){
							f.Delete();
							tablecount = tablecount - 1;
						}
					}
					
					if ((objApp.ActiveWindow.Selection.ShapeRange.Count >= 2) && (tablecount <= 1)) {
						PowerPoint.Shape slideShapeGroup = objApp.ActiveWindow.Selection.ShapeRange.Group();
//						float percentofTotalSlide = 0.95F; //Now taken from textbox
						slideShapeGroup.LockAspectRatio = MsoTriState.msoTrue;
						if ((objPresNew.PageSetup.SlideHeight / objPresNew.PageSetup.SlideWidth) >= (slideShapeGroup.Height / slideShapeGroup.Width)) {
							slideShapeGroup.ScaleWidth(objPresNew.PageSetup.SlideWidth / slideShapeGroup.Width * percentofTotalSlide, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft);
						}
						else {
							slideShapeGroup.ScaleHeight(objPresNew.PageSetup.SlideHeight / slideShapeGroup.Height * percentofTotalSlide, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft);
						}
						slideShapeGroup.Select(MsoTriState.msoTrue);
						objApp.ActiveWindow.Selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, MsoTriState.msoTrue);
						objApp.ActiveWindow.Selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignMiddles, MsoTriState.msoTrue);
						
						objApp.ActiveWindow.Selection.Unselect();
					}
	//				else{
	//					currentSlide.Shapes.SelectAll();
	//					objApp.ActiveWindow.Selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, MsoTriState.msoTrue);
	//					objApp.ActiveWindow.Selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignMiddles, MsoTriState.msoTrue);
	//				}
				}
				
	//			objPresNew.Slides[1].Delete();
	//			objApp.ActiveWindow.Selection.Unselect();
				
	
				
				//Resize stuck shapes - superseded by going to old method of copy/pasting slide w/ formatting start to finish					
	//			foreach (PowerPoint._Slide d in objPresNew.Slides) {
	//				foreach (PowerPoint.Shape e in d.Shapes) {
	//					Console.WriteLine(e.Name);
	//					
	//					if (e.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape) {
	//						PowerPoint.Shape eOld = objPres.Slides[d.SlideNumber].Shapes[e.Name];
	//						
	//						if (e.Height != eOld.Height) {
	//							e.Height = eOld.Height;
	//						}
	//						
	//					}
	//				}
	//			}
				
				//Delete notes master if it exists
				var seq2 = Enumerable.Range(1,objPresNew.SlideMaster.Shapes.Count).ToArray();
				objPresNew.SlideMaster.Shapes.Range(seq2).Delete();
				
				//Delete last hanging slide
				objPresNew.Slides[objPresNew.Slides.Count].Delete();
		
				//Save new pres
				objPresNew.SaveAs(pptCorrectedPresPath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoFalse);
				
				//Displaying "Resizing Done" on successful resizing, and closing original box if needed
				if (!resizeDialog.IsDisposed){
				resizeDialog.Close();
				}
	
				NotificationForm resizeDialog2 = new NotificationForm();
				resizeDialog2.StartPosition = FormStartPosition.CenterParent;
				resizeDialog2.TopMost = false;
				resizeDialog2.Text = "Resizing Done";
				resizeDialog2.label1.AutoSize = true;
				resizeDialog2.label1.Location = new System.Drawing.Point(13, 13);
				resizeDialog2.label1.Text = "PowerPoint file successfully resized; you can find the new file at " + 
					Environment.NewLine + Environment.NewLine + pptCorrectedPresPath + Environment.NewLine + Environment.NewLine + "Please close this box to continue.";
				resizeDialog2.Controls.Add(resizeDialog2.label1);
				resizeDialog2.AutoSize = true;
				resizeDialog2.ShowDialog(pptWindow);
	//			resizeDialog2.Activate();
	
				IsOpen_Close(pptPresName, objApp, false);
			}
		
			public static void IsOpen_Close(string presentation, PowerPoint.Application pptApp, bool correctedflag)
		    {
				if (!correctedflag)
		        {		
			        foreach (PowerPoint._Presentation pptCheck in pptApp.Presentations)
			        {
		
			    		if ((pptCheck.Name.Contains(presentation)) && !(pptCheck.Name.Contains("_corrected")))
			        	{
			            	pptCheck.Close();
			            }
			        }
		    	}
				else
				{
			        foreach (PowerPoint._Presentation pptCheck in pptApp.Presentations)
			        {
		
			    		if (pptCheck.Name.Contains(presentation))
			        	{
			            	pptCheck.Close();
			            }
			        }				
				}
			}
		void Label2Click(object sender, EventArgs e)
		{
		
		}
		void Button2Click(object sender, EventArgs e)
		{
			string pptPresPath;
			
				//Pick presentation for resizing.
				OpenFileDialog openFileDialog1 = new OpenFileDialog();
				openFileDialog1.Filter = "PowerPoint Files|*.pptx;*.ppt;*.pptm";
				openFileDialog1.Title = "Select the presentation for renumbering";
				
				if (openFileDialog1.ShowDialog() == DialogResult.OK)
			    	pptPresPath = openFileDialog1.FileName;
			    else
			    {
			    	pptPresPath = string.Empty;
			    	return;
			    }
			    
			    PowerPoint.Application objApp; 
				PowerPoint.Presentations objPresSet;
				PowerPoint._Presentation objPres;
				PowerPoint._Presentation objPresNew;
				PowerPoint.Slides objSlides;

				string pptPresName;
				string pptCorrectedPresPath;
			    
			    pptPresName = Path.GetFileNameWithoutExtension(pptPresPath);
			    
			    pptCorrectedPresPath = Path.GetDirectoryName(pptPresPath) + "\\" + pptPresName + "_renumbered" + Path.GetExtension(pptPresPath);
			    
			    if (pptPresName.Contains("_corrected"))
			    {
			    	DialogResult dialogResult1 = MessageBox.Show("It looks like you selected a renumbered version of this file - you selected:" +
			    	    Environment.NewLine + Environment.NewLine + pptPresName + Environment.NewLine + Environment.NewLine +
			    	    "Are you sure you want to proceed (this will create a _renumbered_renumbered version)?", "Correct corrected file", 
			    	    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2,(MessageBoxOptions)0x40000);
			    	
			    	if(dialogResult1 == DialogResult.No)
				    	return;
			    }
			    
			    if (File.Exists(pptCorrectedPresPath))
			    {
			    	DialogResult dialogResult1 = MessageBox.Show("A renumbered version of this file appears to exist already in this folder." +
			    	    " Do you want to proceed (this will overwrite existing renumbered version)?", "Overwrite corrected file",
			    	    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2,(MessageBoxOptions)0x40000);
				    if(dialogResult1 == DialogResult.No)
				    	return;
			    }
			    
				//Open selected presentation.
				objApp = new PowerPoint.Application();
				objApp.Visible = MsoTriState.msoTrue;
				objPresSet = objApp.Presentations;
				
				//clear the original version if one is already open
				IsOpen_Close(pptPresName, objApp, false);
				
				objPresNew = objPresSet.Open(pptPresPath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
				
				//clear the fixed version if one is already open
				IsOpen_Close(pptPresName + "_renumbered", objApp, true);
				
				objPresNew.SaveAs(pptCorrectedPresPath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoFalse);
								
				//Initial notification window
				NativeWindow pptWindow = new NativeWindow();
				pptWindow.AssignHandle(new IntPtr(objApp.HWND));
		
				NotificationForm resizeDialog = new NotificationForm();
				resizeDialog.StartPosition = FormStartPosition.CenterParent;
				resizeDialog.Text = "Renumbering";
				resizeDialog.label1.AutoSize = true;
				resizeDialog.label1.Location = new System.Drawing.Point(13, 13);
				resizeDialog.label1.Text = "Renumbering presentation; please do not close PowerPoint or this presentation until the 'Renumbering Done' message box appears." +
					Environment.NewLine + Environment.NewLine + "Please close this box to continue.";
				resizeDialog.Controls.Add(resizeDialog.label1);
				resizeDialog.AutoSize = true;
				resizeDialog.ShowDialog(pptWindow);
				
				//Copy and paste old small presentation slides into new size presentation
	//			var seq = Enumerable.Range(1,objPres.Slides.Count).ToArray(); //old method all slides at once
	//			objPres.Slides.Range(seq).Copy();
//				int y = 1;
//				while (y <= 8)
//				{
					int i = 1;
					int prtCounter;
					int keyrow;
					y = y + 1;
					bool tableflag = false;
					
					var partTbl = new List<partRef>();
					
					while (i <= objPresNew.Slides.Count) {
						PowerPoint.Slide currentSlide = objPresNew.Slides[i];
						foreach (PowerPoint.Shape f in currentSlide.Shapes) {
							if (f.HasTable == MsoTriState.msoTrue) {
								keyrow = 1;
								foreach (PowerPoint.Cell colheader in f.Table.Rows[1].Cells) {
									if ((colheader.Shape.TextFrame.TextRange.Text.Contains("Item") || colheader.Shape.TextFrame.TextRange.Text.Contains("Find"))
									   && colheader.Shape.TextFrame.TextRange.Text.Length < 7) {
										tableflag = true;
										prtCounter = 1;
										foreach (PowerPoint.Cell prt in f.Table.Rows[keyrow].Cells.Count) {
											partTbl.Add( new partRef {
											        SlidePartRef = prt.Shape.TextFrame.TextRange.TrimText();
											    	partTrueNumber = f.Table.Rows[keyrow+1].Cells[prtCounter].Shape.TextFrame.TextRange.TrimText();
											});
								            prtCounter++;
										}
										break;
									}
									keyrow++;
								}
							}
						}
						i++;
					}
					
					while (i <= objPresNew.Slides.Count) {
						PowerPoint.Slide currentSlide = objPresNew.Slides[i];
						foreach (PowerPoint.Shape f in currentSlide.Shapes) {
							if (f.HasTable == MsoTriState.msoTrue) {
								keyrow = 1;
								foreach (PowerPoint.Cell colheader in f.Table.Rows[1].Cells) {
									if ((colheader.Shape.TextFrame.TextRange.Text.Contains("Item") || colheader.Shape.TextFrame.TextRange.Text.Contains("Find"))
									   && colheader.Shape.TextFrame.TextRange.Text.Length < 7) {
										tableflag = true;
										prtCounter = 1;
										foreach (PowerPoint.Cell prt in f.Table.Rows[keyrow].Cells.Count) {
											partTbl.Add( new partRef {
											        SlidePartRef = prt.Shape.TextFrame.TextRange.TrimText();
											    	partTrueNumber = f.Table.Rows[keyrow+1].Cells[prtCounter].Shape.TextFrame.TextRange.TrimText();
											});
								            prtCounter++;
										}
										break;
									}
									keyrow++;
								}
							}
						}
						i++;
					}
//				}//trying without group destroyer if possible

				//Delete notes master if it exists
				var seq2 = Enumerable.Range(1,objPresNew.SlideMaster.Shapes.Count).ToArray();
				objPresNew.SlideMaster.Shapes.Range(seq2).Delete();
				
				//Delete last hanging slide
				objPresNew.Slides[objPresNew.Slides.Count].Delete();
		
				//Save new pres
				objPresNew.SaveAs(pptCorrectedPresPath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoFalse);
				
				//Displaying "Resizing Done" on successful resizing, and closing original box if needed
				if (!resizeDialog.IsDisposed){
				resizeDialog.Close();
				}
	
				NotificationForm resizeDialog2 = new NotificationForm();
				resizeDialog2.StartPosition = FormStartPosition.CenterParent;
				resizeDialog2.TopMost = false;
				resizeDialog2.Text = "Resizing Done";
				resizeDialog2.label1.AutoSize = true;
				resizeDialog2.label1.Location = new System.Drawing.Point(13, 13);
				resizeDialog2.label1.Text = "PowerPoint file successfully resized; you can find the new file at " + 
					Environment.NewLine + Environment.NewLine + pptCorrectedPresPath + Environment.NewLine + Environment.NewLine + "Please close this box to continue.";
				resizeDialog2.Controls.Add(resizeDialog2.label1);
				resizeDialog2.AutoSize = true;
				resizeDialog2.ShowDialog(pptWindow);
	//			resizeDialog2.Activate();
				
	
				IsOpen_Close(pptPresName, objApp, false);
				
				}
			}
	
	public class partRef
	{
		public string SlidePartRef { get; set; }
		public string partTrueNumber { get; set; }
	}
	
	public class NotificationForm: Form
	{
		public Label label1 = new Label();
	}

}
