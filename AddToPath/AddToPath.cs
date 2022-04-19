using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace AddToPath
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Directory)]
    public class AddToPath : SharpContextMenu
    {
        private string SelectedItemPath => SelectedItemPaths.First();
        private string ItemRegexPattern =>  $"(?<=^|;){Regex.Escape(SelectedItemPath)}(?=$|;)";

        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();
            var isInPath = IsInPath();

            var itemCountLines = new ToolStripMenuItem()
            {
                Text = isInPath ? "Remove from PATH (User)" : "Add to PATH (User)",
            };

            itemCountLines.Click += (sender, args) =>
            {
                if (isInPath)
                {
                    RunInSeparateThread(RemoveSelectedItemFromPath);
                }
                else
                {
                    RunInSeparateThread(AddSelectedItemToPath);
                }
            };

            menu.Items.Add(itemCountLines);

            return menu;
        }

        private void RunInSeparateThread(ThreadStart start)
        {
            var worker = new Thread(start);
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }
        private void AddSelectedItemToPath()
        {
            var oldPath = GetPath();
            var newPath = $"{oldPath};{SelectedItemPath}";
            SetPath(newPath);
        }
        
        private void RemoveSelectedItemFromPath()
        {
            var oldPath = GetPath();
            var newPath = Regex.Replace(oldPath, ItemRegexPattern, "");
            SetPath(newPath);
        }

        private bool IsInPath()
        {
            return Regex.Match(GetPath(), ItemRegexPattern).Success;
        }
        
        private static string GetPath()
        {
            return Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User);
        }

        private static void SetPath(string newPath)
        {
            Environment.SetEnvironmentVariable("Path", newPath, EnvironmentVariableTarget.User);
        }
    }
}