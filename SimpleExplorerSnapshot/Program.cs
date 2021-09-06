using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SimpleExplorerSnapshot
{
	class Program
	{
		//icon source - https://icons8.com/icon/3jnrVS52owjR/card-file-box

		static void Main()
		{
			var controlPressed = ((Control.ModifierKeys & Keys.Control) != 0);
			var shiftPressed = ((Control.ModifierKeys & Keys.Shift) != 0);

			//

			var app_dir = Path.GetFullPath(Debugger.IsAttached ? @"..\..\..\..\" : @".\");
			var output_dir = Path.Combine(app_dir, "ExplorerSnapshots");
			Directory.CreateDirectory(output_dir);

			//

			var capture = true;
			var open_output = true;
			var close_explorer_windows = false;

			//

			if (controlPressed) capture = false; //if the control key is pressed, do not capture explorer addresses, only open output dir...
			if (shiftPressed) close_explorer_windows = true; //if the shift key is pressed, gracefully close all explorer windows

			//

			if (capture) CaptureAllExplorerAddresses(output_dir);
			if (close_explorer_windows) CloseAllExplorerWindows();
			if (open_output) Process.Start("explorer.exe", output_dir);
		}

		private static void CloseAllExplorerWindows()
		{
			var script = @"(New-Object -comObject Shell.Application).Windows() | foreach-object {$_.quit()}";
			var process = new Process
			{
				StartInfo = new ProcessStartInfo("powershell.exe", script)
				{
					WorkingDirectory = Environment.CurrentDirectory,
					RedirectStandardOutput = true,
					CreateNoWindow = true,
				}
			};
			process.Start();

			var reader = process.StandardOutput;
			var result = reader.ReadToEnd();

			if (!string.IsNullOrWhiteSpace(result))
			{
				EventLog.WriteEntry(".NET Runtime", AppDomain.CurrentDomain.FriendlyName + " - Error executing powershell.exe script - " + result, EventLogEntryType.Error, 1000 /*1000 <= eventID < 1029*/);
			}

			process.Dispose();
		}

		private static void CaptureAllExplorerAddresses(string output_dir)
		{
			var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss.fff");
			var filePath = Path.Combine(output_dir, $"ExplorerSnapshot.{timestamp}.bat");
			var contents = WindowList.GetAllWindowsExtendedInfo()
				.Where(x => x?.Class == "ToolbarWindow32") // filter
				.Where(x => x?.Parent?.Class == "Breadcrumb Parent") // filter
				.Select(x => x.Caption.Substring("Address: ".Length)) // capture directory path
				.Where(x => Directory.Exists(x)) // filter unusable
				.Select(x => string.Format("start \"\" /B \"{0}\"", x)) // generate command from template
				.Distinct() // only one of each
				.OrderBy(x => x) // sort
				.ToArray();
			File.WriteAllLines(filePath, contents);
		}
	}
}
