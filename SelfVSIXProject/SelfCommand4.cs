using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace SelfVSIXProject
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SelfCommand4
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 256;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("b8cecac6-d457-491e-a942-d806e2db23ff");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SelfCommand4"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SelfCommand4(AsyncPackage package, OleMenuCommandService commandService)
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
        public static SelfCommand4 Instance
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
            // Switch to the main thread - the call to AddCommand in SelfCommand4's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SelfCommand4(package, commandService);
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
            ThreadHelper.ThrowIfNotOnUIThread();

            //自定义功能
            //操作选中功能, 自动补全方法或属性或字段左边的变量 (.cs文件 )
            DTE dte = ServiceProvider.GetServiceAsync(typeof(DTE)).Result as DTE;
            string newStr = string.Empty;
            if (dte?.ActiveDocument?.Name?.EndsWith(".cs", StringComparison.CurrentCultureIgnoreCase) ?? false)
            {
                var selection = (TextSelection)dte.ActiveDocument.Selection;
                selection.SelectLine();
                string curRowContent = selection?.Text;
                if (!string.IsNullOrWhiteSpace(curRowContent))
                {
                    //当前行有多种情况需要补全：
                    //1 new Class().Method() ;   只有补全var aaa = xxx;
                    //2 obj.Method();            只有补全var aaa = xxx;
                    //3 StaticClass.Method();       只有补全var aaa = xxx;
                    //4 Method();
                    //5 new Class();
                    newStr = $"var aaa = {curRowContent.TrimStart()}";
                    string matchStr = string.Empty;
                    //1. new Class().Method(); @Regex ^\s*new\s+\w+\s*\(\s*.*\s*\)\s*\.\s*\w+\s*\(\s*.*\s*\)\s*;\s*$
                    if (Regex.IsMatch(curRowContent, @"^\s*new\s+\w+\s*\(\s*.*\s*\)\s*\.\s*\w+\s*\(\s*.*\s*\)\s*;\s*$"))
                    {
                        selection.Text = newStr;
                    }
                    //2. obj.Method(); @Regex ^\s*\w+\s*\.\s*\w+\s*\(\s*.*\s*\)\s*;\s*$
                    //3. StaticClass.Method(); @Regex ^\s*\w+\s*\.\s*\w+\s*\(\s*.*\s*\)\s*;\s*$
                    else if (Regex.IsMatch(curRowContent, @"^\s*\w+\s*\.\s*\w+\s*\(\s*.*\s*\)\s*;\s*$"))
                    {
                        selection.Text = newStr;
                    }
                    //4. Method(); @Regex ^\s*\w+\s*\(\s*.*\s*\)\s*;\s*$
                    else if (Regex.IsMatch(curRowContent, @"^\s*\w+\s*\(\s*.*\s*\)\s*;\s*$"))
                    {
                        selection.Text = newStr;
                    }
                    //5. new Class(); @Regex ^\s*new\s*\w+\s*\(\s*.*\s*\)\s*;\s*$
                    else if (Regex.IsMatch(curRowContent, @"^\s*new\s*\w+\s*\(\s*.*\s*\)\s*;\s*$"))
                    {
                        selection.Text = newStr;
                    }

                }
            }
            else
            {
                ShowMsgBox("请在.cs文件中使用变量自动补全功能");
            }
        }

        public void ShowMsgBox(string msg)
        {
            VsShellUtilities.ShowMessageBox(this.package, msg, "", OLEMSGICON.OLEMSGICON_NOICON, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
