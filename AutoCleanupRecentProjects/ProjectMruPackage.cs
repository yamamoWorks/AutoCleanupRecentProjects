//------------------------------------------------------------------------------
// <copyright file="ProjectMruPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace AutoCleanupRecentProjects
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(ProjectMruPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly",
         Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(UIContextGuids.NoSolution)]
    public sealed class ProjectMruPackage : Package
    {
        /// <summary>
        /// ProjectMruPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "d4603dd2-b10c-4731-b45e-8ebca6631c8b";

        /// <summary>
        /// Initializes a new instance of the <see cref="ProjectMruPackage"/> class.
        /// </summary>
        public ProjectMruPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            #region other way
            // this way will be reflected at next startup.

            //var regKey = UserRegistryRoot.OpenSubKey("MRUItems")?
            //    .OpenSubKey("{a9c4a31f-f9cb-47a9-abc0-49ce82d0b3ac}")?
            //    .OpenSubKey("Items", true);

            //if (regKey == null)
            //    return;

            //var mruItems = regKey.GetValueNames()
            //    .Select(name => (string)regKey.GetValue(name))
            //    .Where(value => File.Exists(value.Split('|').First()))
            //    .ToArray();

            //foreach (var name in regKey.GetValueNames())
            //{
            //    regKey.DeleteValue(name);
            //}

            //for (var i = 0; i < mruItems.Length; i++)
            //{
            //    regKey.SetValue(i.ToString(), mruItems[i]);
            //}

            //regKey.Close();

            #endregion

            try
            {
                var dataSourceFactory = GetService(typeof(SVsDataSourceFactory)) as IVsDataSourceFactory;
                if (dataSourceFactory == null)
                    return;

                IVsUIDataSource projectMruList;
                dataSourceFactory.GetDataSource(Guid.Parse("9099ad98-3136-4aca-a9ac-7eeeaee51dca"), 1, out projectMruList);

                var projectMruListType = Type.GetType("Microsoft.VisualStudio.PlatformUI.ProjectMruList, Microsoft.VisualStudio.Shell.UI.Internal", true);
                var removeAtItemMethod = projectMruListType.GetMethod("RemoveItemAt");
                var itemsProperty = projectMruListType.GetProperty("Items");

                var fileSystemMruItemType = Type.GetType("Microsoft.VisualStudio.PlatformUI.FileSystemMruItem, Microsoft.VisualStudio.Shell.UI.Internal", true);
                var pathProperty = fileSystemMruItemType.GetProperty("Path");

                var items = (IList)itemsProperty.GetValue(projectMruList, null);
                for (var i = items.Count - 1; i > -1; i--)
                {
                    var path = (string)pathProperty.GetValue(items[i], null);
                    if (!File.Exists(path))
                    {
                        removeAtItemMethod.Invoke(projectMruList, new object[] { i });
                    }
                }
            }
            catch { }
        }

        #endregion
    }
}
