//------------------------------------------------------------------------------
// <copyright file="ForceUTF8Package.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using EnvDTE;
using EnvDTE80;
using System.IO;
using System.Text;

namespace ForceUTF8
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
    [Guid(ForceUTF8Package.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    public sealed class ForceUTF8Package : Package
    {
        /// <summary>
        /// ForceUTF8Package GUID string.
        /// </summary>
        public const string PackageGuidString = "d5ca34f2-6bde-4b84-b853-39c003dbda4d";

        /// <summary>
        /// Initializes a new instance of the <see cref="ForceUTF8Package"/> class.
        /// </summary>
        public ForceUTF8Package()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        #region Package Members

        private DocumentEvents documentEvents;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {            
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            var dte = GetService(typeof(DTE)) as DTE2;
            documentEvents = dte.Events.DocumentEvents;
            documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;
        }

        void DocumentEvents_DocumentSaved(Document document)
        {
            if (document.Kind != "{8E7B96A8-E33D-11D0-A6D5-00C04FB67F6A}")
            {
                // then it's not a text file
                return;
            }

            var path = document.FullName;

            var stream = new FileStream(path, FileMode.Open);
            var reader = new StreamReader(stream, Encoding.Default, true);
            reader.Read();

            if (reader.CurrentEncoding != Encoding.Default)
            {
                stream.Close();
                return;
            }

            string text;

            try
            {
                stream.Position = 0;
                reader = new StreamReader(stream, new UTF8Encoding(true, true));
                reader.ReadToEnd();
                stream.Close();
            }
            catch (DecoderFallbackException)
            {
                stream.Position = 0;
                reader = new StreamReader(stream, Encoding.Default);
                text = reader.ReadToEnd();
                stream.Close();
                File.WriteAllText(path, text, new UTF8Encoding(true));
            }
        }
        #endregion
    }
}
