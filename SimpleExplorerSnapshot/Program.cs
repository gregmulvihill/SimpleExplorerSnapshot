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
			var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss.fff");
			var app_dir = Path.GetFullPath(Debugger.IsAttached ? @"..\..\..\..\" : @".\");
			var output_dir = Path.Combine(app_dir, "ExplorerSnapshots");
			Directory.CreateDirectory(output_dir);

			//if the shift key is pressed, don't capture, just open output dir
			if ((Control.ModifierKeys & Keys.Shift) == 0)
			{
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

			Process.Start("explorer.exe", output_dir);
		}
	}
}
