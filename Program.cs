using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using System.Collections.Concurrent;
using System.Net;
using System.Text;
using Path = System.IO.Path;

List<string> pptExtensions = new() { "pptx", "pptm", "ppt", "potx", "potm", "pot", "ppxs", "ppsm", "pps", "ppam", "ppa" };

while (true)
{
    var filePath = string.Empty;
    Console.Write("Enter full path to file (e.g. C:\\Users\\me\\Documents\\My Presentation.pptx): ");
    do
    {
        filePath = string.Empty;
        filePath = Console.ReadLine();
        //filePath = @"C:\Users\boclifton\OneDrive - Microsoft\Documents\Presentations\Azure TUB\2022-09 - Azure-Technical Update Briefing.pptx";

        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Enter a valid path with filename and extension (e.g. C:\\Users\\me\\Documents\\My Presentation.pptx).");
        }
        if (!File.Exists(filePath))
        {
            Console.WriteLine("File not found.  Enter a valid path with filename and extension (e.g. C:\\Users\\me\\Documents\\My Presentation.pptx).");
            filePath = string.Empty;
            continue;
        }
        if (!pptExtensions.Contains(Path.GetExtension(filePath).Replace(".", "")))
        {
            Console.WriteLine($"This application only support PowerPoint files.  Please select a file with one of the following extensions: {string.Join(", ", pptExtensions)}");
            filePath = string.Empty;
            continue;
        }
    } while (string.IsNullOrEmpty(filePath));


    Console.Write("Getting links from document...");
    List<string> links = GetAllExternalLinksInPresentation(filePath);
    Console.WriteLine("done.");

    Console.WriteLine("Writing links to output file...");
    ConcurrentBag<string> outputText = new();

    Parallel.ForEach(links, l =>
    {
        var title = GetPageTitle(l);
        if (string.IsNullOrEmpty(title))
            title = "--No page title available--";
        var line = $"{title} ({l})";
        if (!outputText.Contains(line))
        {
            outputText.Add(line);
            if (outputText.Count % 10 == 0)
                Console.WriteLine($"Written {outputText.Count} links...");
        }
    });

    var fileName = Path.GetFileName(filePath);

    using StreamWriter streamWriter = new(@$"C:\Users\boclifton\OneDrive - Microsoft\Documents\Presentations\Azure TUB\{fileName} - Links.txt");
    streamWriter.WriteLine(string.Join("\n", outputText.ToArray()));
    Console.WriteLine("done. All finished.");

    Console.Write("Get links for another file? ( y / N ): ");
    var continueChoice = Console.ReadLine();
    if (continueChoice is null || string.IsNullOrEmpty(continueChoice) || continueChoice.ToLowerInvariant() == "n")
        break;
}

static List<string> GetAllExternalLinksInPresentation(string filePath)
{
    List<string> urls = new();
    using PresentationDocument document = PresentationDocument.Open(filePath, false);
    // Iterate through all the slide parts in the presentation part.
    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
    {
        IEnumerable<HyperlinkType> links = slidePart.Slide.Descendants<HyperlinkType>();

        // Iterate through all the links in the slide part.
        foreach (HyperlinkType link in links)
        {
            // Iterate through all the external relationships in the slide part. 
            foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
            {
                // If the relationship ID matches the link ID…
                if (relation.Id.Equals(link.Id))
                {
                    // Add the URI of the external relationship to the list of strings.
                    urls.Add(relation.Uri.AbsoluteUri);
                }
            }
        }
    }
    return urls;
}

static string GetPageTitle(string url)
{
    var webGet = new HtmlWeb();
    var document = webGet.Load(url);
    var title = string.Empty;

    if (webGet.StatusCode == HttpStatusCode.OK)
    {
        title = document.DocumentNode.SelectSingleNode("html/head/title")?.InnerText;

        if (title is null || title.Contains("Sign in to your account"))
            return string.Empty;

        return ReturnCleanASCII(title);
    }

    return title;
}

static string ReturnCleanASCII(string s)
{
    StringBuilder sb = new(s.Length);
    foreach (char c in s)
    {
        if ((int)c > 127) // you probably don't want 127 either
            continue;
        if ((int)c < 32)  // I bet you don't want control characters 
            continue;
        if (c == '%')
            continue;
        if (c == '?')
            continue;
        sb.Append(c);
    }

    return sb.ToString();
}