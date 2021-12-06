using System.Windows;
using System.Windows.Input;
using EnvDTE;
using Window = System.Windows.Window;

namespace KsWare.VSOpenFile {

	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window {

		public MainWindow() {
			InitializeComponent();
			FileNameTextBox.Focus();
		}

		private void FileNameTextBox_KeyDown(object sender, KeyEventArgs e) {
			if (e.Key == Key.Enter) Open_Click(null, null);
		}

		private void Open_Click(object sender, RoutedEventArgs e) {
			var file = FileNameTextBox.Text;
			var lineNumber = 0;
			var linePosition = 0;
			if (file.Contains(";")) {
				var fragments = file.Split(";");
				file = fragments[0];
				lineNumber = int.Parse(fragments[1]);
				if (fragments.Length >= 3) linePosition = int.Parse(fragments[2]);
			}

			var dte = Tools.GetVsDteContainingFile(file, out var fullFileName);

			if (dte == null) {
				MessageBox.Show("No matching VS found.");
				return;
			}

			file = fullFileName;
			// dte.ExecuteCommand("Edit.OpenFile", $"\"{filePath
			dte.ItemOperations.OpenFile(file);
			(dte.ActiveDocument.Selection as TextSelection).MoveToLineAndOffset(lineNumber, linePosition);
			dte.MainWindow.Activate();
		}

		private void About_Click(object sender, RoutedEventArgs e) {
			new About().ShowDialog();
		}
	}

}
