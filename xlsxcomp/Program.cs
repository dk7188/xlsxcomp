/*
 * Created by SharpDevelop.
 * User: 53785
 * Date: 2017/12/18
 * Time: 下午 02:43
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
namespace xlsxcomp
{
	class Program
	{
		public static void Main(string[] args)
		{
			Console.WriteLine("=== Program Start ===");
			Stopwatch stopWatch = new Stopwatch();
			stopWatch.Start();
			
			
			string updatedFilename = "demo-new.xlsx";
			string previousFilename = "demo-old.xlsx";
			ExcelComparator xlcp = new ExcelComparator();
			xlcp.UpdatedFileFullPath = updatedFilename;
			xlcp.PreviousFileFullPath = previousFilename;
			xlcp.LoadWorkbook();
			xlcp.test();
			xlcp.ReleaseResource();

			// TODO: Implement Functionality Here

			stopWatch.Stop();
			// Get the elapsed time as a TimeSpan value.
			TimeSpan ts = stopWatch.Elapsed;

			// Format and display the TimeSpan value.
			string elapsedTime = String.Format("{0:00}:{1:00}",
				                     ts.Seconds,
				                     ts.Milliseconds / 10);
			Console.WriteLine("*** Overall RunTime : " + elapsedTime);
			
			Console.WriteLine("=== Program End ===");
			Console.ReadKey(true);
		}
	}
}