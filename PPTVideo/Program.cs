using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.IO;

// Based on article
// http://support.microsoft.com/kb/303718

namespace PPTVideo
{
	class Program
	{
		static PowerPoint.Application objApp;
		
		static void Main(string[] args)
		{
			PowerPoint._Presentation objPres;
			objApp = new PowerPoint.Application();
			//objApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
			objPres = objApp.Presentations.Open(Directory.GetCurrentDirectory() + "\\" + args[0], MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoTrue);
			objPres.SaveAs(Directory.GetCurrentDirectory() + "\\" + args[1], PowerPoint.PpSaveAsFileType.ppSaveAsWMV, MsoTriState.msoTriStateMixed);
			long len = 0;
			do{
				System.Threading.Thread.Sleep(500);
				try {
					FileInfo f = new FileInfo(args[1]);
					len = f.Length;
				} catch {
					continue;
				}
			} while (len == 0);
			objApp.Quit();
			File.Delete(args[0]);
		}
	}
}
