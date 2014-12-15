using System;
using Gtk;
using VPPListBuddy.Workflow;

public partial class MainWindow: Gtk.Window
{
	//Helper Functions
	public void AlertDialog(string Message)
	{
		var messagedialog = new MessageDialog (null, DialogFlags.Modal, MessageType.Error, ButtonsType.Ok, Message);
		messagedialog.Response += (o, args) => messagedialog.Destroy();
		messagedialog.Run ();
	}
	//Setup Drag Drop and OpenXLSWorkflow
	public TargetEntry [] targettable = new TargetEntry[] {new TargetEntry("text/uri-list",0,0)};
	public OpenXLSWorkFlow XLSParser = new OpenXLSWorkFlow();
	public MainWindow () : base (Gtk.WindowType.Toplevel)
	{
		Build ();
		XLSParser.OnNonExistentFile = new FileError (path => AlertDialog (String.Format ("File {0} does not exist!", path)));
		XLSParser.OnParseFailure = new FileError (path => AlertDialog (String.Format ("File at {0} could not be read as an XLS spreadsheet.", path)));
		XLSParser.OnEmptyWorkbook = new FileError (path => AlertDialog (String.Format ("XLS at {0} is an empty workbook.", path)));
		XLSParser.OnInvalidVPP = new Alert (() => AlertDialog ("Worksheet 1 does not contain Apple VPP Information or is formatted incorectly."));
		XLSParser.OnInvalidPartitionSheet = new Alert (() => AlertDialog("Worksheet 2 does not exist or is not formatted correctly."));
		Gtk.Drag.DestSet (labelDragTarget, DestDefaults.All, targettable, Gdk.DragAction.Copy);
	}
	protected void OnMenubar1DragDataReceived (object o, DragDataReceivedArgs args)
	{  
		string data = System.Text.Encoding.UTF8.GetString (args.SelectionData.Data);
		switch (args.Info) {
		case 0:  // uri-list
			string[  ] uri_list = System.Text.RegularExpressions.Regex.Split (data, "\r\n");
			var uri = new Uri (uri_list [0]);
			var filename = uri.LocalPath;
			OpenVPPPartition (filename);
			break;
		}
	}
	//Actual Program
	public void OpenVPPPartition(string Path) {
		PartitionWorkflow partitioner;
		if (XLSParser.TryOpenVPP (Path, out partitioner)) {
			Console.WriteLine("Success!");
		}
	}
	public string SaveFolderChooser()
	{
		//Should return null or empty string for failure.
		//Build File Chooser
		var title = "Save to"; 
		Window parent = null;
		var choosertype = FileChooserAction.SelectFolder;
		var filechooser = new FileChooserDialog (title, parent, choosertype,"Cancel",Gtk.ResponseType.Cancel,
			"Open",Gtk.ResponseType.Accept);

		var startdirectory = System.Environment.GetFolderPath (System.Environment.SpecialFolder.UserProfile);

		filechooser.SelectMultiple = false;
		filechooser.SetCurrentFolder(startdirectory);
		do {
			//stuff
		} while ((ResponseType)(filechooser.Run ()) != ResponseType.Accept);
		var directory = filechooser.CurrentFolder;
		filechooser.Destroy ();
		return directory;
	}
	//Etc.
	protected void OnDeleteEvent (object sender, DeleteEventArgs a)
	{
		Application.Quit ();
		a.RetVal = true;
	}
}
