using Forge.Forms;
using Forge.Forms.Annotations;
using Newtonsoft.Json;
using SuperMemoAssistant.Interop.SuperMemo.Registry.Members;
using SuperMemoAssistant.Services;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.ComponentModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;

namespace SuperMemoAssistant.Plugins.EpubImporter
{
  [Form(Mode = DefaultFields.None)]
  [Title("Dictionary Settings",
   IsVisible = "{Env DialogHostContext}")]
  [DialogAction("cancel",
  "Cancel",
  IsCancel = true)]
  [DialogAction("save",
  "Save",
  IsDefault = true,
  Validates = true)]
  public class EpubImporterCfg : CfgBase<EpubImporterCfg>, INotifyPropertyChangedEx
  {
    [Title("Epub Importer Plugin")]
    [Heading("By Jamesb | Experimental Learning")]

    [Heading("Features")]
    [Text(@"- Import epub books into SuperMemo.")]

    [Heading("Support")]
    [Text("If you would like to support my projects, check out my Patreon or buy me a coffee.")]

    [Action("patreon", "Patreon", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("coffee", "Coffee", Placement = Placement.Before, LinePosition = Position.Left)]

    [Heading("Links")]
    [Action("github", "GitHub", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("feedback", "Feedback Site", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("blog", "Blog", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("youtube", "YouTube", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("twitter", "Twitter", Placement = Placement.Before, LinePosition = Position.Left)]

    [Heading("Settings")]
    [Field(Name = "Default Priority?")]
    [Value(Must.BeGreaterThanOrEqualTo,
           0,
           StrictValidation = true)]
    [Value(Must.BeLessThanOrEqualTo,
           100,
           StrictValidation = true)]
    public double DefaultPriority { get; set; } = 30;

    [Field(Name = "Default Template")]
    [SelectFrom("{Binding Templates}", DisplayPath = "Name", ValuePath = "Id", SelectionType = SelectionType.ComboBox)]
    public int DefaultTemplate { get; set; } = -1;

    [JsonIgnore]
    public IEnumerable<TemplateShim> Templates =>
      new List<TemplateShim> { new TemplateShim("(none)", -1) }
        .Concat(Svc.SM.Registry.Template.Select(t => new TemplateShim(t)))
        .ToList();

    [JsonIgnore]
    public bool IsChanged { get; set; }

    public bool HasAgreedToTOS { get; set; } = false;

    public override string ToString()
    {
      return "Epub Importer Settings";
    }

    public override void HandleAction(IActionContext actionContext)
    {

      string patreon = "https://www.patreon.com/experimental_learning";
      string coffee = "https://buymeacoffee.com/experilearning";
      string github = "https://github.com/bjsi/SuperMemoAssistant.Plugins.EpubImporter";
      string feedback = "https://feedback.experimental-learning.com/";
      string youtube = "https://www.youtube.com/channel/UCIaS9XDdQkvIjASBfgim1Uw";
      string twitter = "https://twitter.com/experilearning";
      string blog = "https://www.experimental-learning.com/";

      string action = actionContext.Action as string;
      if (action == "patreon")
        openLinkDefaultBrowser(patreon);
      else if (action == "github")
        openLinkDefaultBrowser(github);
      else if (action == "coffee")
        openLinkDefaultBrowser(coffee);
      else if (action == "feedback")
        openLinkDefaultBrowser(feedback);
      else if (action == "youtube")
        openLinkDefaultBrowser(youtube);
      else if (action == "twitter")
        openLinkDefaultBrowser(twitter);
      else if (action == "blog")
        openLinkDefaultBrowser(blog);
      else
        base.HandleAction(actionContext);
    }

    // Hack
    private DateTime LastLinkOpen { get; set; } = DateTime.MinValue;

    private void openLinkDefaultBrowser(string url)
    {
      var diffInSeconds = (DateTime.Now - LastLinkOpen).TotalSeconds;
      if (diffInSeconds > 1)
      {
        LastLinkOpen = DateTime.Now;
        Process.Start(url);
      }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public class TemplateShim
    {
      #region Constructors

      public TemplateShim(ITemplate template)
      {
        Name = template.Name;
        Id = template.Id;
      }

      public TemplateShim(string name, int id)
      {
        Name = name;
        Id = id;
      }

      #endregion




      #region Properties & Fields - Public

      public string Name { get; }
      public int Id { get; }

      #endregion
    }
  }
}
