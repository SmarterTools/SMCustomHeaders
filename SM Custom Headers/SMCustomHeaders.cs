using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace SM_Custom_Headers
{
	// This service 
	public partial class SMCustomHeaders : ServiceBase
	{
		private string procFolder = @"C:\SmarterMail16\Spool\proc";
		private string dropFolder = @"C:\SmarterMail16\Spool\Drop";

		public SMCustomHeaders()
		{
			InitializeComponent();
		}

		protected override void OnStart(string[] args)
		{
			// Set up a watcher for the folder that watches when file names change
			FileSystemWatcher watcher = new FileSystemWatcher();
			watcher.Path = procFolder;
			watcher.NotifyFilter = NotifyFilters.FileName;
			watcher.Renamed += new RenamedEventHandler(OnChanged);
			watcher.EnableRaisingEvents = true;

			// Deal with the messages already in the folder
			string[] files = Directory.GetFiles(procFolder);

			foreach (string file in files)
			{
				AddHeaders(file);
			}
		}

		private void OnChanged(object source, FileSystemEventArgs e)
		{
			AddHeaders(e.FullPath);
		}

		private void AddHeaders(string filePath)
		{
			// The more try catches the merrier, because if this thing breaks, so does your spool
			try
			{
				if (!filePath.EndsWith(".eml")) return;
				string nameNoExtension = Path.GetFileNameWithoutExtension(filePath);

				try
				{
					// Parse .hdr file into a dictionary
					Dictionary<string, string> hdr = new Dictionary<string, string>();
					using (StreamReader sr = new StreamReader(File.OpenRead($"{procFolder}\\{nameNoExtension}.hdr"), Encoding.UTF8))
					{
						string text = "";
						while (!sr.EndOfStream)
						{
							text = sr.ReadLine();
							// .hdr variables are in the form of VariableName: VariableValue
							// So, if a variable exists on this line
							// Grab everything before the ':' as the dictionary key and everything after as the value
							int index = text.IndexOf(':');
							if (index > -1)
								hdr.Add(text.Substring(0, index), text.Substring(index + 1, text.Length - index - 1).Trim());
						}
					}

					// Read in the email
					string eml = "";
					using (StreamReader sr = new StreamReader(File.OpenRead(filePath), Encoding.UTF8))
					{
						eml = sr.ReadToEnd();
					}

					// Write headers to the top of the message
					eml = eml.Insert(0, $"X-Src-Sender: ({hdr["connectedHostName"]}) [{hdr["connectedIP"]}]:{hdr["connectedPort"]}\r\n");

					// Write out email
					using (StreamWriter sw = new StreamWriter(File.OpenWrite(filePath), Encoding.UTF8))
					{
						sw.Write(eml);
					}

					// NOTE:
					// If there is another program, such as Declude, that is also using the proc folder it should use another folder
					// Then, instead of moving these to the drop folder it should move them to another proc folder specifically for the other program
					File.Move($"{procFolder}\\{nameNoExtension}.eml", $"{dropFolder}\\{nameNoExtension}.eml");
					File.Move($"{procFolder}\\{nameNoExtension}.hdr", $"{dropFolder}\\{nameNoExtension}.hdr");
				}
				catch
				{
					File.Move($"{procFolder}\\{nameNoExtension}.eml", $"{dropFolder}\\{nameNoExtension}.eml");
					File.Move($"{procFolder}\\{nameNoExtension}.hdr", $"{dropFolder}\\{nameNoExtension}.hdr");
				}
			}
			catch { }
		}

		protected override void OnStop()
		{
		}
	}
}
