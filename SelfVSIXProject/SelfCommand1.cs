using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace SelfVSIXProject
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SelfCommand1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("93d4d243-8331-475a-aead-f962c6e967f9");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SelfCommand1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SelfCommand1(AsyncPackage package, OleMenuCommandService commandService)
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
        public static SelfCommand1 Instance
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
            // Switch to the main thread - the call to AddCommand in SelfCommand1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SelfCommand1(package, commandService);
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
            //操作选中功能, 用正则表达式判断是否匹配成功
            //格式例：(inputStr)(pattern)  =>  (123qwe)([0-9]+) 则提示true
            DTE dte = ServiceProvider.GetServiceAsync(typeof(DTE)).Result as DTE;
            string selectTXT = string.Empty;
            if (dte.ActiveDocument != null && dte.ActiveDocument.Type == "Text")
            {
                var selection = (TextSelection)dte.ActiveDocument.Selection;
                string text = selection?.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    ShowMsgBox("您正在使用正则表达式匹配判断功能，请先选中内容，且格式为:\n inputStr @Regex pattern ");
                }
                else
                {
                    text = text.Trim();
                    bool isMatch = Regex.IsMatch(text, @"^.+\@Regex.+$");
                    if (!isMatch)
                    {
                        ShowMsgBox("您正在使用正则表达式匹配判断功能，请使用格式为:\n inputStr @Regex pattern ");
                    }
                    else
                    {
                        //inputStr @Regex pattern
                        string[] Strs = text.Split(new string[] { "@Regex" }, StringSplitOptions.RemoveEmptyEntries);
                        if (Strs == null || Strs.Length != 2)
                        {
                            ShowMsgBox("您正在使用正则表达式匹配判断功能，请使用格式为:\n inputStr @Regex pattern ");
                        }
                        else
                        {
                            ShowMsgBox($"匹配结果：{(Regex.IsMatch(Strs[0].Trim(), Strs[1].Trim()) ? "匹配成功" : "匹配失败")}");
                        }

                    }
                    //selection.Text = text; //相当于修改当前代码，在两端加内容
                    //selectTXT = text; //直接用selection.Text无效
                }
            }

        }

        public void ShowMsgBox(string msg)
        {
            VsShellUtilities.ShowMessageBox(this.package, msg, "", OLEMSGICON.OLEMSGICON_NOICON, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
