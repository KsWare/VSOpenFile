using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.VisualStudio.Setup.Configuration;
using Process = System.Diagnostics.Process;
using Thread = System.Threading.Thread;

namespace KsWare.VSOpenFile {

	public static class Tools {

		public static string[] ProjectFileExtensions = { ".csproj", ".vbproj" };

		public static EnvDTE.DTE GetVsDte() {
			var names = GetRunningObjectNames("!VisualStudio.DTE.16.0:");
			var runningObjectDisplayName = names.First();
			object runningObject;
			IEnumerable<string> runningObjectDisplayNames = null;
			try {
				runningObject = GetRunningObject(runningObjectDisplayName, out runningObjectDisplayNames);
			}
			catch {
				runningObject = null;
			}

			if (runningObject != null) {
				return (EnvDTE.DTE)runningObject;
			}

			throw new TimeoutException(
				$"Failed to retrieve DTE object. Current running objects: {string.Join(";", runningObjectDisplayNames)}");
		}

		public static EnvDTE.DTE GetVsDte(string runningObjectDisplayName) {
			object runningObject;
			IEnumerable<string> runningObjectDisplayNames = null;
			try {
				runningObject = GetRunningObject(runningObjectDisplayName, out runningObjectDisplayNames);
			}
			catch {
				runningObject = null;
			}

			if (runningObject != null) {
				return (EnvDTE.DTE)runningObject;
			}

			throw new TimeoutException(
				$"Failed to retrieve DTE object. Current running objects: {string.Join(";", runningObjectDisplayNames)}");
		}

		public static EnvDTE.DTE GetVsDteContainingFile(string file, out string fullFileName) {
			var names = GetRunningObjectNames("!VisualStudio.DTE.16.0:");
			if (Path.IsPathRooted(file)) {
				if (IsSolutionFile(file)) {
					foreach (var runningObjectDisplayName in names) {
						var dte = GetVsDte(runningObjectDisplayName);
						if (string.Equals(dte.Solution.FullName, file, StringComparison.OrdinalIgnoreCase)) {
							fullFileName = file;
							return dte;
						}
					}
				} else if (IsProjectFile(file)) {
					foreach (var runningObjectDisplayName in names) {
						var dte = GetVsDte(runningObjectDisplayName);
						foreach (EnvDTE.Project project in dte.Solution.Projects) {
							if (string.Equals(project.FullName, file, StringComparison.OrdinalIgnoreCase)) {
								fullFileName = file;
								return dte;
							}
						}
					}
				}
				else {
					foreach (var runningObjectDisplayName in names) {
						var dte = GetVsDte(runningObjectDisplayName);
						var solutionPath = Path.GetDirectoryName(dte.Solution.FullName);
						if (file.StartsWith(solutionPath)) {
							fullFileName = file;
							return dte;
						}
					}
				}

				fullFileName = null;
				return null;
			}
			else {
				foreach (var runningObjectDisplayName in names) {
					var dte = GetVsDte(runningObjectDisplayName);
					foreach (EnvDTE.Project project in dte.Solution.Projects) {
						var projectPath = Path.GetDirectoryName(project.FullName);
						var possibleFile = Path.Combine(projectPath, file);
						if (File.Exists(possibleFile)) {
							fullFileName = possibleFile;
							return dte;
						}
					}
				}
				fullFileName = null;
				return null;
			}

		}

		private static bool IsProjectFile(string file) {
			return ProjectFileExtensions.Contains(Path.GetExtension(file));
		}

		private static bool IsSolutionFile(string file) {
			return Path.GetExtension(file) == ".sln";
		}

		public static EnvDTE.DTE LaunchVsDte(bool isPreRelease) {
			var setupInstance = GetSetupInstance(isPreRelease);
			var installationPath = setupInstance.GetInstallationPath();
			var executablePath = Path.Combine(installationPath, @"Common7\IDE\devenv.exe");
			var vsProcess = Process.Start(executablePath);
			var runningObjectDisplayName = $"VisualStudio.DTE.16.0:{vsProcess.Id}";

			IEnumerable<string> runningObjectDisplayNames = null;
			object runningObject;
			for (var i = 0; i < 60; i++) {
				try {
					runningObject = GetRunningObject(runningObjectDisplayName, out runningObjectDisplayNames);
				}
				catch {
					runningObject = null;
				}

				if (runningObject != null) {
					return (EnvDTE.DTE)runningObject;
				}

				Thread.Sleep(millisecondsTimeout: 1000);
			}

			throw new TimeoutException(
				$"Failed to retrieve DTE object. Current running objects: {string.Join(";", runningObjectDisplayNames)}");
		}

		public static object GetRunningObject(string displayName, out IEnumerable<string> runningObjectDisplayNames) {
			IBindCtx bindContext = null;
			NativeMethods.CreateBindCtx(0, out bindContext);

			IRunningObjectTable runningObjectTable = null;
			bindContext.GetRunningObjectTable(out runningObjectTable);

			IEnumMoniker monikerEnumerator = null;
			runningObjectTable.EnumRunning(out monikerEnumerator);

			object runningObject = null;
			var runningObjectDisplayNameList = new List<string>();
			var monikers = new IMoniker[1];
			var numberFetched = IntPtr.Zero;
			while (monikerEnumerator.Next(1, monikers, numberFetched) == 0) {
				var moniker = monikers[0];

				string objectDisplayName = null;
				try {
					moniker.GetDisplayName(bindContext, null, out objectDisplayName);
				}
				catch (UnauthorizedAccessException) {
					// Some ROT objects require elevated permissions.
				}

				if (!string.IsNullOrWhiteSpace(objectDisplayName)) {
					runningObjectDisplayNameList.Add(objectDisplayName);
					if (objectDisplayName.EndsWith(displayName, StringComparison.Ordinal)) {
						runningObjectTable.GetObject(moniker, out runningObject);
						if (runningObject == null) {
							throw new InvalidOperationException(
								$"Failed to get running object with display name {displayName}");
						}
					}
				}
			}

			runningObjectDisplayNames = runningObjectDisplayNameList;
			return runningObject;
		}

		public static IEnumerable<string> GetRunningObjectNames(string partialDisplayName) {
			NativeMethods.CreateBindCtx(0, out var bindContext);
			bindContext.GetRunningObjectTable(out var runningObjectTable);
			runningObjectTable.EnumRunning(out var monikerEnumerator);

			var runningObjectDisplayNameList = new List<string>();
			var monikers = new IMoniker[1];
			var numberFetched = IntPtr.Zero;
			while (monikerEnumerator.Next(1, monikers, numberFetched) == 0) {
				var moniker = monikers[0];

				string objectDisplayName = null;
				try {
					moniker.GetDisplayName(bindContext, null, out objectDisplayName);
				}
				catch (UnauthorizedAccessException) {
					// Some ROT objects require elevated permissions.
				}

				if (!string.IsNullOrWhiteSpace(objectDisplayName)) {
					if (objectDisplayName.Contains(partialDisplayName))
						runningObjectDisplayNameList.Add(objectDisplayName);
				}
			}

			return runningObjectDisplayNameList;
		}

		private static ISetupInstance GetSetupInstance(bool isPreRelease) {
			return GetSetupInstances().First(i => IsPreRelease(i) == isPreRelease);
		}

		private static IEnumerable<ISetupInstance> GetSetupInstances() {
			ISetupConfiguration setupConfiguration = new SetupConfiguration();
			var enumerator = setupConfiguration.EnumInstances();

			int count;
			do {
				var setupInstances = new ISetupInstance[1];
				enumerator.Next(1, setupInstances, out count);
				if (count == 1 && setupInstances[0] != null) {
					yield return setupInstances[0];
				}
			} while (count == 1);
		}

		private static bool IsPreRelease(ISetupInstance setupInstance) {
			var setupInstanceCatalog = (ISetupInstanceCatalog)setupInstance;
			return setupInstanceCatalog.IsPrerelease();
		}

		private static class NativeMethods {
			[DllImport("ole32.dll")]
			public static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
		}

		
	}

}
