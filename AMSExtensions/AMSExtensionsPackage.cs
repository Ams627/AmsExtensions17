using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using System.Collections.Generic;

namespace AMS.AMSExtensions
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidAMSExPkgString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class AMSExtensionsPackage : Package
    {

        private static readonly string[] _toolWindowGuids = new[]
        {
            "{0038D243-39E6-4004-8532-13C6802D97B2}",
            "{007A3A6B-B5B2-454D-A2BD-CF929F989BE2}",
            "{0504FF91-9D61-11D0-A794-00A0C9110051}",
            "{099CA9EA-0AE4-4e31-A7E4-FE09BD1715CC}",
            "{0ad07096-bba9-4900-a651-0598d26f6d24}",
            "{0CF2BFF9-6057-4D23-A51C-A4825B06BFFE}",
            "{0F887920-C2B6-11D2-9375-0080C747D9A0}",
            "{0F887921-C2B6-11D2-9375-0080C747D9A0}",
            "{116d2292-e37d-41cd-a077-ebacac4c8cc4}",
            "{131369F2-062D-44A2-8671-91FF31EFB4F4}",
            "{13143d73-808b-4a2f-bebb-ccbe9c93fe37}",
            "{13b12e3e-c1b4-4539-9371-4fe9a0d523fc}",
            "{1651552d-26f6-4591-b1c5-654d75c10c00}",
            "{1820BAE5-C385-4492-9DE5-E35C9CF17B18}",
            "{1CBA9826-3184-4799-A184-784E41B56398}",
            "{2456bd12-ecf7-4988-a4a6-67d49173f565}",
            "{2553315d-efa1-4d33-9ed6-ffe62257ca21}",
            "{269A02DC-6AF8-11D3-BDC4-00C04F688E50}",
            "{28836128-FC2C-11D2-A433-00C04F72D18A}",
            "{2d4d7b63-2712-49a6-a309-3ff689de5696}",
            "{2D7728C2-DE0A-45B5-99AA-89B609DFDE73}",
            "{2ffc3976-88be-4e8a-8913-68245bb123f9}",
            "{31fc2115-5126-4a87-b2f7-77eaab65048b}",
            "{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}",
            "{37ABA9BE-445A-11D3-9949-00C04F68FD0A}",
            "{3822E751-EB69-4B0E-B301-595A9E4C74D5}",
            "{38610995-0ca9-460b-8482-321ac4ee2f9d}",
            "{387CB18D-6153-4156-9257-9AC3F9207BBE}",
            "{38ED9834-0C97-445b-BD1D-F78F3E08AFAC}",
            "{3addf8e2-81cc-41a0-9785-dbd2d86064bf}",
            "{3AE79031-E1BC-11D0-8F78-00A0C9110057}",
            "{3db98ce8-14f0-4b05-9b6a-b9dbe8fa94ab}",
            "{3e06c0a3-4a65-4963-8b1c-69be6043471b}",
            "{3f28efa6-37c0-4c33-95a7-c72f22064023}",
            "{402DC223-D700-4029-866F-ACEE803F3F0C}",
            "{497A7D24-3B23-4BF5-9C35-4B82E7A499FA}",
            "{4A18F9D0-B838-11D0-93EB-00A0C90F2734}",
            "{4a8466fe-5a05-4f81-ad45-7d7c1316817e}",
            "{4a9b7e51-aa16-11d0-a8c5-00a0c921a4d2}",
            "{4f8106d6-f048-4f17-b184-028a79ee938d}",
            "{4fa6ebce-130a-4282-a617-45828b337ffe}",
            "{519e8a32-1c95-4a42-956f-2cee2f28eb0f}",
            "{51C76317-9037-4CF2-A20A-6206FD30B4A1}",
            "{53544C4D-5C18-11d3-AB71-0050040AE094}",
            "{5415EA3A-D813-4948-B51E-562082CE0887}",
            "{56B32054-DE4D-4de3-8396-BCB6F98BD246}",
            "{588470CC-84F8-4A57-9AC4-86BCA0625FF4}",
            "{58875C41-862B-4d6f-B046-03E8A333907E}",
            "{5A4E9529-B6A0-46B5-BE4F-0F0B239BC0EB}",
            "{5da8d69b-058a-4294-a87a-f4f7ae2983b3}",
            "{605322a2-17ae-43f4-b60f-766556e46c87}",
            "{65DDF8C3-8F89-4077-A6C6-DBB8853AAB13}",
            "{66dba47c-61df-11d2-aa79-00c04f990343}",
            "{682e7d2a-5c60-4f00-be98-e787ff511646}",
            "{68487888-204A-11D3-87EB-00C04F7971A5}",
            "{6d3b0c34-bbce-4550-9fe8-4725d644ece5}",
            "{6d4078d1-5951-4ed1-ac0e-0a8099c1cce5}",
            "{6f0b3bcf-f1c8-49f5-a0b5-e3c29762d9a5}",
            "{6FB4A4D9-0C08-4663-AF7B-2ECBDF7A20EC}",
            "{74946810-37a0-11d2-a273-00c04f8ef4ff}",
            "{74946827-37A0-11D2-A273-00C04F8EF4FF}",
            "{778B5376-AD77-4751-ACDC-F3D18343F8DD}",
            "{7B8C4981-13EC-4c56-9F24-ABE5FAAA9440}",
            "{7f075f9f-1e1a-4ee8-9ce8-b8451fb61ec8}",
            "{80454082-9AB8-47d4-AF23-82BF6739E2A9}",
            "{809F6FF3-8092-454A-8003-6D4091F9B5BB}",
            "{83107a3e-496a-485e-b455-16ddb978e55e}",
            "{83107a3e-496a-485e-b455-16ddb978e66f}",
            "{873151D0-CF2E-48cc-B4BF-AD0394F6A3C3}",
            "{8D263989-FF4B-4a78-90C8-B2BA3FA69311}",
            "{905da7d1-18fd-4a46-8d0f-a5ff58ada9de}",
            "{92547016-2bd0-4dfe-bd4f-5b52bdce0037}",
            "{99b8fa2f-ab90-4f57-9c32-949f146f1914}",
            "{9A7CEBBB-DC5C-4986-BC49-962DA46AA506}",
            "{A0C5197D-0AC7-4B63-97CD-8872A789D233}",
            "{a2c8b774-0984-495f-8428-e6140b2d66e0}",
            "{a2eaf38f-a0ad-4503-91f8-5f004a69a040}",
            "{A975660A-81E4-4830-AA25-1B824C3BF732}",
            "{b1e99781-ab81-11d0-b683-00aa00a3ee26}",
            "{B9A151CE-EF7C-4fe1-A6AA-4777E6E518F3}",
            "{bbd3659f-2d77-3663-8b41-70dda83597cb}",
            "{bdada759-5db7-402f-96e3-402b17c8c5b4}",
            "{BE4D7042-BA3F-11D2-840E-00C04F9902C1}",
            "{c79b74ff-f1d7-4c94-aefa-4d22bfe1b1f9}",
            "{c93a910a-0fa6-4307-93a4-f2bd61ec7828}",
            "{C9C0AE26-AA77-11D2-B3F0-0000F87570EE}",
            "{CA4B8FF5-BFC7-11D2-9929-00C04F68FDAF}",
            "{cb170882-2b60-46b5-9d06-a57667186c95}",
            "{cdbdee54-b399-484b-b763-db2c3393d646}",
            "{CF2DDC32-8CAD-11D2-9302-005345000000}",
            "{CF577B8C-4134-11D2-83E5-00C04F9902C1}",
            "{d78612c7-9962-4b83-95d9-268046dad23a}",
            "{dd1ddd20-d59b-11da-a94d-0800200c9a66}",
            "{dd9b7693-1385-46a9-a054-06566904f861}",
            "{ddd3c1ac-37df-4e57-bd73-6c4c0e1cc5ba}",
            "{DE1FC918-F32E-4DD7-A915-1792A051F26B}",
            "{dee22b65-9761-4a26-8fb2-759b971d6dfc}",
            "{e1b7d1f8-9b3c-49b1-8f4f-bfc63a88835d}",
            "{E3FC08BE-3924-11DB-8AF6-B622A1EF5492}",
            "{E62CE6A0-B439-11D0-A79D-00A0C9110051}",
            "{e77209ba-064a-4625-b8ce-dfd1d7967cd1}",
            "{ECB7191A-597B-41F5-9843-03A4CF275DDE}",
            "{eefa5220-e298-11d0-8f78-00a0c9110057}",
            "{F2E84780-2AF1-11D1-A7FA-00A0C9110051}",
            "{F62AF5AD-1276-46dd-AE7B-D07AB54D1081}",
            "{fbcae063-e2c0-4ab1-a516-996ea3dafb72}"
        };

        public AMSExtensionsPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidAMSExCmdSet, (int)PkgCmdIDList.cmdidCloseCrap);
                CommandID menuCommandID2 = new CommandID(GuidList.guidAMSExCmdSet, (int)PkgCmdIDList.cmdidOpenConApp);
                CommandID menuCommandID3 = new CommandID(GuidList.guidAMSExCmdSet, (int)PkgCmdIDList.cmdidOpenCSharpConApp);
                CommandID menuCommandID4 = new CommandID(GuidList.guidAMSExCmdSet, (int)PkgCmdIDList.cmdidOpenWPFApp);
                MenuCommand mc1 = new MenuCommand(ProcessCloseCrap, menuCommandID);
                MenuCommand mc2 = new MenuCommand(ProcessOpenConApp, menuCommandID2);
                MenuCommand mc3 = new MenuCommand(ProcessOpenCSharpConApp, menuCommandID3);
                MenuCommand mc4 = new MenuCommand(ProcessOpenWPFApp, menuCommandID4);
                mcs.AddCommand(mc1);
                mcs.AddCommand(mc2);
                mcs.AddCommand(mc3);
                mcs.AddCommand(mc4);
            }
        }

        void ProcessCloseCrap(object sender, EventArgs e)
        {
            var dte = (DTE)GetService(typeof(EnvDTE.DTE));
            foreach (var guidstring in _toolWindowGuids)
            {
                try
                {
                   dte.Windows.Item(guidstring)?.Close();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(guidstring);
                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                    System.Diagnostics.Debug.WriteLine("==============================");
                }
            }
        }

        void SetOutputWindowToBuildPane()
        {
            IVsOutputWindowPane outputWindowPane;
            var outputWindow = base.GetService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            var buildPane = Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid.BuildOutputPane_guid;
            int hr = outputWindow.GetPane(buildPane, out outputWindowPane);
            outputWindowPane.Activate();
        }

        void ProcessOpenConApp(object sender, EventArgs e)
        {
            Debug.WriteLine("Open Console App");
            DTE dte = (DTE)GetService(typeof(EnvDTE.DTE));

            var prjs = dte.Solution.Projects;
            if (prjs.Count > 0)
            {
                string s = prjs.Item(1).FullName;
                s = s.Replace(".vcxproj", ".cpp");
                dte.ItemOperations.OpenFile(s);
                dte.ActiveDocument.Activate();
                dte.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr;
                dte.Find.FindWhat = "int main\\(.*\\n\\{.*\\n";
                dte.Find.Target = vsFindTarget.vsFindTargetCurrentDocument;
                dte.Find.MatchCase = false;
                dte.Find.MatchWholeWord = false;
                dte.Find.Backwards = false;
                dte.Find.MatchInHiddenText = false;
                dte.Find.Action = vsFindAction.vsFindActionFind;
                if (dte.Find.Execute() == vsFindResult.vsFindResultNotFound)
                {
                    throw new System.Exception("vsFindResultNotFound");
                }
                dte.Windows.Item("{CF2DDC32-8CAD-11D2-9302-005345000000}").Close();
                TextSelection ts = (TextSelection)dte.ActiveDocument.Selection;
                ts.CharRight();
                ts.Indent();

                SetOutputWindowToBuildPane();
                dte.ExecuteCommand("File.SaveAll", string.Empty);
                dte.ExecuteCommand("Build.BuildSolution");
                dte.ActiveDocument.Activate();
            }
        }

        void ProcessOpenCSharpConApp(object sender, EventArgs e)
        {
            Debug.WriteLine("Open CSharp Console App");
            DTE dte = (DTE)GetService(typeof(EnvDTE.DTE));

            var prjs = dte.Solution.Projects;
            if (prjs.Count > 0)
            {
                // get full path of solution (including name of solution file):
                string solutionFullPathname = dte.Solution.FullName;

                // extract the directory of the solution:
                string dir = System.IO.Path.GetDirectoryName(solutionFullPathname);

                // get the name of the first project
                string projectName = dte.Solution.Projects.Item(1).Name;

                // combine the solution dir with the project dir and "program.cs":
                string programfname = System.IO.Path.Combine(dir, projectName, "program.cs");
                //System.Windows.Forms.MessageBox.Show("hello: " + programfname);

                dte.ItemOperations.OpenFile(programfname, EnvDTE.Constants.vsViewKindCode);
                var findString = @"static void Main\(string\[\] args\)\r?\n +{";

                dte.ActiveDocument.Activate();
                dte.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr;
                dte.Find.FindWhat = findString;
                dte.Find.Target = vsFindTarget.vsFindTargetCurrentDocument;
                dte.Find.MatchCase = false;
                dte.Find.MatchWholeWord = false;
                dte.Find.Backwards = false;
                dte.Find.MatchInHiddenText = false;
                dte.Find.Action = vsFindAction.vsFindActionFind;
                if (dte.Find.Execute() == vsFindResult.vsFindResultNotFound)
                {
                    throw new System.Exception("vsFindResultNotFound");
                }
                dte.Windows.Item("{CF2DDC32-8CAD-11D2-9302-005345000000}").Close();
                TextSelection ts = (TextSelection)dte.ActiveDocument.Selection;
                ts.CharRight();
                ts.Indent();

                SetOutputWindowToBuildPane();

                dte.ExecuteCommand("File.SaveAll", string.Empty);
                dte.ExecuteCommand("Build.BuildSolution");
                dte.ActiveDocument.Activate();
            }

        }

        void ProcessOpenWPFApp(object sender, EventArgs e)
        {
            Debug.WriteLine("Open WPF App");
            DTE dte = (DTE)GetService(typeof(EnvDTE.DTE));

            var prjs = dte.Solution.Projects;
            if (prjs.Count > 0)
            {
                // get full path of solution (including name of solution file):
                string solutionFullPathname = dte.Solution.FullName;

                // extract the directory of the solution:
                string dir = System.IO.Path.GetDirectoryName(solutionFullPathname);

                // get the name of the first project
                string projectName = dte.Solution.Projects.Item(1).Name;

                // combine the solution dir with the project dir and "program.cs":
                string programfname = System.IO.Path.Combine(dir, projectName, "mainwindow.xaml");
                //System.Windows.Forms.MessageBox.Show("hello: " + programfname);

                dte.ItemOperations.OpenFile(programfname, EnvDTE.Constants.vsViewKindCode);
                //var findString = @"static void Main\(string\[\] args\)\r?\n +{";

                //dte.ActiveDocument.Activate();
                //dte.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr;
                //dte.Find.FindWhat = findString;
                //dte.Find.Target = vsFindTarget.vsFindTargetCurrentDocument;
                //dte.Find.MatchCase = false;
                //dte.Find.MatchWholeWord = false;
                //dte.Find.Backwards = false;
                //dte.Find.MatchInHiddenText = false;
                //dte.Find.Action = vsFindAction.vsFindActionFind;
                //if (dte.Find.Execute() == vsFindResult.vsFindResultNotFound)
                //{
                //    throw new System.Exception("vsFindResultNotFound");
                //}
                //dte.Windows.Item("{CF2DDC32-8CAD-11D2-9302-005345000000}").Close();
                //TextSelection ts = (TextSelection)dte.ActiveDocument.Selection;
                //ts.CharRight();
                //ts.Indent();

                SetOutputWindowToBuildPane();

                dte.ExecuteCommand("File.SaveAll", string.Empty);
                dte.ExecuteCommand("Build.BuildSolution");
                dte.ActiveDocument.Activate();
            }

        }

        void ProcessOpenWebAPICoreAPP(object sender, EventArgs e)
        {
            Debug.WriteLine("Open WPF App");
            DTE dte = (DTE)GetService(typeof(EnvDTE.DTE));

            var prjs = dte.Solution.Projects;
            if (prjs.Count > 0)
            {
                // get full path of solution (including name of solution file):
                string solutionFullPathname = dte.Solution.FullName;

                // extract the directory of the solution:
                string dir = System.IO.Path.GetDirectoryName(solutionFullPathname);

                // get the name of the first project
                string projectName = dte.Solution.Projects.Item(1).Name;

                // combine the solution dir with the project dir and "program.cs":
                string programfname = System.IO.Path.Combine(dir, projectName, "mainwindow.xaml");
                //System.Windows.Forms.MessageBox.Show("hello: " + programfname);

                dte.ItemOperations.OpenFile(programfname, EnvDTE.Constants.vsViewKindCode);
                //var findString = @"static void Main\(string\[\] args\)\r?\n +{";

                //dte.ActiveDocument.Activate();
                //dte.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr;
                //dte.Find.FindWhat = findString;
                //dte.Find.Target = vsFindTarget.vsFindTargetCurrentDocument;
                //dte.Find.MatchCase = false;
                //dte.Find.MatchWholeWord = false;
                //dte.Find.Backwards = false;
                //dte.Find.MatchInHiddenText = false;
                //dte.Find.Action = vsFindAction.vsFindActionFind;
                //if (dte.Find.Execute() == vsFindResult.vsFindResultNotFound)
                //{
                //    throw new System.Exception("vsFindResultNotFound");
                //}
                //dte.Windows.Item("{CF2DDC32-8CAD-11D2-9302-005345000000}").Close();
                //TextSelection ts = (TextSelection)dte.ActiveDocument.Selection;
                //ts.CharRight();
                //ts.Indent();

                SetOutputWindowToBuildPane();

                dte.ExecuteCommand("File.SaveAll", string.Empty);
                dte.ExecuteCommand("Build.BuildSolution");
                dte.ActiveDocument.Activate();
            }

        }

    }
}
