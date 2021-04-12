using Anotar.Serilog;
using EpubSharp;
using HtmlAgilityPack;
using Microsoft.Win32;
using SuperMemoAssistant.Interop.Plugins;
using SuperMemoAssistant.Interop.SuperMemo.Content.Contents;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Builders;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Types;
using SuperMemoAssistant.Services;
using SuperMemoAssistant.Services.IO.HotKeys;
using SuperMemoAssistant.Services.IO.Keyboard;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.IO.Devices;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;

#region License & Metadata

// The MIT License (MIT)
// 
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the 
// Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.
// 
// 
// Created On:   1/24/2021 11:05:26 AM
// Modified By:  james

#endregion




namespace SuperMemoAssistant.Plugins.EpubImporter
{
  // ReSharper disable once UnusedMember.Global
  // ReSharper disable once ClassNeverInstantiated.Global
  [SuppressMessage("Microsoft.Naming", "CA1724:TypeNamesShouldNotMatchNamespaces")]
  public class EpubImporterPlugin : SMAPluginBase<EpubImporterPlugin>
  {
    #region Constructors

    /// <inheritdoc />
    public EpubImporterPlugin() { }

    #endregion




    #region Properties Impl - Public

    /// <inheritdoc />
    public override string Name => "EpubImporter";

    /// <inheritdoc />
    public override bool HasSettings => true;

    private SemaphoreSlim ImportSemaphore { get; } = new SemaphoreSlim(1, 1);

    public EpubImporterCfg Config { get; private set; }

    #endregion

    private async Task LoadConfig()
    {
      Config = await Svc.Configuration.Load<EpubImporterCfg>() ?? new EpubImporterCfg();
    }



    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {

      LoadConfig().Wait();
      
      Svc.HotKeyManager.RegisterGlobal(
        "ImportEPUB",
        "Import an Epub file",
        HotKeyScope.SMBrowser,
        new HotKey(Key.Y, KeyModifiers.CtrlShift),
        OpenFile
        );
    }

    #endregion


    #region Methods

    private void OpenFile()
    {
      if (ImportSemaphore.Wait(0) == false)
        return;

      string filePath = OpenFileDialog();

      if (filePath != null)
      {
        ImportEpub(filePath);
      }

      ImportSemaphore.Release();
    }

    public void ImportEpub(string filepath)
    {
      try
      {
        var book = EpubReader.Read(filepath);
        var parentId = CreateBookFolder(book);
        var parentEl = Svc.SM.Registry.Element[parentId];

        if (parentId < 1 || parentEl == null)
        {
          LogTo.Debug("Failed to create book folder.");
          return;
        }

        var unzipped = UnzipEpub(filepath);
        if (!Directory.Exists(unzipped))
        {
          LogTo.Debug("Failed to unzip Epub - the unzipped folder does not exist.");
          return;
        }

        var baseUrl = Path.Combine(unzipped, Path.GetDirectoryName(book.Format.Ocf.RootFilePath));
        foreach (var chapter in book.TableOfContents)
        {
          if (ImportChapter(parentEl, book, chapter, baseUrl) == -1)
          {
            LogTo.Debug("Failed to import chapter.");
            break;
          }
        }
      }
      catch (Exception e)
      {
        LogTo.Debug($"Failed to import {filepath} with exception {e}");
      }
    }

    private string UnzipEpub(string originalFilepath)
    {
      try
      {
        var outputName = Path.GetFileNameWithoutExtension(originalFilepath) + DateTime.Now.Ticks.ToString();
        var outputBase = Path.GetDirectoryName(originalFilepath);
        var outputPath = Path.Combine(outputBase, outputName);
        Directory.CreateDirectory(outputPath);
        System.IO.Compression.ZipFile.ExtractToDirectory(originalFilepath, outputPath);
        return outputPath;
      }
      catch (Exception e)
      {
        LogTo.Error($"Failed to extract epub zip file with exception {e}");
        return null;
      }
    }

    private int CreateBookFolder(EpubBook book)
    {
      return CreateSMFolder(CreateBookReference(book));
    }

    private References CreateBookReference(EpubBook book)
    {
      return new References()
        .WithAuthor(string.Join(", ", book.Authors))
        .WithTitle(book.Title);
    }

    private References CreateChapterReference(EpubBook book, EpubChapter chapter)
    {
      return new References()
        .WithAuthor(string.Join(", ", book.Authors))
        .WithTitle(book.Title + ": " + chapter.Title);
    }

    private int ImportChapter(IElement parentFolder, EpubBook book, EpubChapter chapter, string baseUrl)
    {
      var refs = CreateChapterReference(book, chapter);
      try
      {
        string html = book.Resources.Html.Where(x => x.AbsolutePath == chapter.AbsolutePath).First().TextContent;
        html = UpdateLocalLinks(html, baseUrl);
        var chapterId = CreateSMTopic(html, refs, parentFolder);
        var chapterEl = Svc.SM.Registry.Element[chapterId];
        if (chapterEl != null)
        {
          foreach (var sub in chapter.SubChapters)
            ImportChapter(chapterEl, book, sub, baseUrl);
        }

        return chapterId;
      }
      catch (Exception ex)
      {
        LogTo.Debug($"Failed to import chapter with exception {ex}");
        return -1;
      }
    }

    private string UpdateLocalLinks(string html, string baseUrl)
    {
      // Change local links to absolute links
      var doc = new HtmlDocument();
      doc.LoadHtml(html);
      foreach (var node in doc.DocumentNode.Descendants())
      {
        if (node.Name != "script")
        {
          AdjustAttributes(node, baseUrl, "href");
          AdjustAttributes(node, baseUrl, "src");
        }
      }

      return doc.DocumentNode.OuterHtml;
    }

    static void AdjustAttributes(HtmlNode root, string baseUrl, string attrName)
    {
      var query =
          from node in root.Descendants()
          let attr = node.Attributes[attrName]
          where attr != null
          select attr;
      foreach (var attr in query)
      {
        attr.Value = GetAbsoluteUrlString(baseUrl, attr.Value);
      }
    }

    static string GetAbsoluteUrlString(string baseUrl, string url)
    {
      var uri = new Uri(url, UriKind.RelativeOrAbsolute);
      if (!uri.IsAbsoluteUri)
      {
        var newUrl = Path.Combine(baseUrl, url.TrimStart('.').TrimStart('\\', '/'));
        return newUrl;
      }
      return uri.ToString();
    }

    private int CreateSMTopic(string html, References refs, IElement parent)
    {
      return CreateSMElement(new List<ContentBase> { new TextContent(true, html) }, refs, false, parent);
    }

    private int CreateSMFolder(References refs)
    {
      return CreateSMElement(new List<ContentBase>(), refs, true, Svc.SM.UI.ElementWdw.CurrentElement);
    }

    [LogToErrorOnException]
    private int CreateSMElement(List<ContentBase> contents, References refs, bool dismissed, IElement parent)
    {
      if (parent == null)
      {
        LogTo.Error("Failed to CreateSMElement because parent was null");
        return -1;
      }

      bool success = Svc.SM.Registry.Element.Add(
        out var val,
        ElemCreationFlags.ForceCreate,
        new ElementBuilder(ElementType.Topic, contents.ToArray())
          .WithPriority(Config.DefaultPriority)
          .WithParent(parent)
          .WithStatus(dismissed ? ElementStatus.Dismissed : ElementStatus.Memorized)
          .DoNotDisplay()
          .WithTitle(refs.Title)
          .WithReference((_) => refs)
      );

      if (success)
      {
        LogTo.Debug("Successfully created SM Element");
        return val[0].ElementId;
      }
      else
      {
        LogTo.Error("Failed to CreateSMElement");
        return -1;
      }
    }

    /// <summary>
    /// Show a dialog which prompts the user to pick an Epub file to import
    /// </summary>
    /// <returns>Filename or null</returns>
    public static string OpenFileDialog()
    {
      OpenFileDialog dlg = new OpenFileDialog
      {
        DefaultExt = ".epub",
        Filter = "EPUB files (*.epub)|*.epub|All files (*.*)|*.*",
        CheckFileExists = true,
      };

      bool res = (bool)dlg.GetType()
          .GetMethod("RunDialog", BindingFlags.NonPublic | BindingFlags.Instance)
          .Invoke(dlg, new object[] { Svc.SM.UI.ElementWdw.Handle });

      return res //dlg.ShowDialog(this).GetValueOrDefault(false)
          ? dlg.FileName
          : null;
    }

    // Set HasSettings to true, and uncomment this method to add your custom logic for settings
    /// <inheritdoc />
    public override void ShowSettings()
    {
      ConfigurationWindow.ShowAndActivate(HotKeyManager.Instance, Config);
    }


    #endregion
  }
}
