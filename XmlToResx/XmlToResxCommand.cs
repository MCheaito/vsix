using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace XmlToResx
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class XmlToResxCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ed559473-4866-4d49-8b5f-ce06c501dad0");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="XmlToResxCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private XmlToResxCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static XmlToResxCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in XmlToResxCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new XmlToResxCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            MenuItemCallbackAsync();

        }

        private async void MenuItemCallbackAsync()
        {
            var dte = (DTE2)await ServiceProvider.GetServiceAsync(typeof(DTE));

            if (CanCreateResx(dte, out string xmlfile))
            {
                var destPath = $@"{Path.GetDirectoryName(xmlfile)}\Properties";

                GenerateResxFile(xmlfile, destPath, null);
                Console.WriteLine("Default Resources file  is done!!");

                GenerateResxFile(xmlfile, destPath, "en-US");
                Console.WriteLine($"Resources file for 'en-US' is done!!");

                GenerateResxFile(xmlfile, destPath, "fr-CA");
                Console.WriteLine($"Resources file for 'fr-CA' is done!!");

                GenerateResxFile(xmlfile, destPath, "ar-LB");
                Console.WriteLine($"Resources file for 'ar-LB' is done!!");
            }

        }


        public void ShowMessage(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "XmlToResxCommand";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
        public static bool CanCreateResx(DTE2 dte, out string xmlFile)
        {
            var items = GetSelectedFiles(dte);

            xmlFile = items.ElementAtOrDefault(0);

            return !string.IsNullOrEmpty(xmlFile) && !string.IsNullOrWhiteSpace(xmlFile) && Path.GetExtension(xmlFile) == ".xml";

        }


        public static IEnumerable<string> GetSelectedFiles(DTE2 dte)
        {
            var items = (Array)dte.ToolWindows.SolutionExplorer.SelectedItems;

            return from item in items.Cast<UIHierarchyItem>()
                   let pi = item.Object as ProjectItem
                   select pi.FileNames[1];
        }


        private static void GenerateResxFile(string filePathName, string destPath, string culture)
        {
            var culturesearch = culture ?? "en-US";
            var runGenerate = culture == null;
            var xdoc = XDocument.Load(filePathName);
            var captions = from seg
                         in xdoc.Descendants("caption")
                           select new
                           {
                               Id = ((XElement)seg.Parent).Attributes("id").FirstOrDefault().Value.TrimEnd() + "_caption",
                               Text = seg.Descendants("text")
                                             .First(i => (string)i.Attribute("id") == culturesearch)
                                             .Attributes("value")
                                             .FirstOrDefault().Value
                           };
            var descriptions = from seg
                         in xdoc.Descendants("description")
                               select new
                               {
                                   Id = ((XElement)seg.Parent).Attributes("id").FirstOrDefault().Value.TrimEnd() + "_description",
                                   Text = seg.Descendants("text")
                                                 .First(i => (string)i.Attribute("id") == culturesearch)
                                                 .Attributes("value")
                                                 .FirstOrDefault().Value
                               };

            var strings = from seg in xdoc.Descendants("string")
                          select new
                          {
                              Id = ((XElement)seg).Attributes("id").FirstOrDefault().Value.TrimEnd()
                              ,
                              Text = seg.Descendants("text")
                                              .First(i => (string)i.Attribute("id") == culturesearch)
                                              .Attributes("value")
                                              .FirstOrDefault().Value
                          };

            var resourceFilePath = $@"{destPath}\Resources.";

            resourceFilePath += (culture != null) ? $"{culture}.resx" : "resx";

            using (var resx = new ResXResourceWriter(resourceFilePath))
            {
                foreach (var item in captions)
                {
                    resx.AddResource(item.Id, item.Text);
                }

                foreach (var item in descriptions)
                {
                    resx.AddResource(item.Id, item.Text);
                }

                foreach (var item in strings)
                {
                    resx.AddResource(item.Id, item.Text);
                }

                if (runGenerate) resx.Generate();
            }

        }

    }
}
