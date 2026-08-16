using System.Net;
using System.Net.Http.Headers;
using ClosedXML.Excel;
using CVIS.WorkbookUpdater.Services;
using Xunit;

namespace CVIS.WorkbookUpdater.Tests;

public sealed class WorkbookSourceProviderTests
{
    [Fact]
    public void Open_DownloadsWorkbookAndDeletesWorkingCopyOnDispose()
    {
        var handler = new RecordingHandler(_ => WorkbookResponse());
        var provider = new WorkbookSourceProvider(new HttpClient(handler));
        string localPath;

        using (var lease = provider.Open("https://files.example.test/recovery.xlsx", writable: false))
        {
            localPath = lease.LocalPath;
            Assert.True(File.Exists(localPath));
            using var workbook = new XLWorkbook(localPath);
            Assert.True(workbook.TryGetWorksheet("Import Mapping", out _));
        }

        Assert.False(File.Exists(localPath));
        Assert.Equal(HttpMethod.Get, handler.Requests.Single().Method);
    }

    [Fact]
    public void Open_RejectsHtmlSharingPage()
    {
        var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent("<html>sign in</html>", MediaTypeHeaderValue.Parse("text/html"))
        });
        var provider = new WorkbookSourceProvider(new HttpClient(handler));

        var error = Assert.Throws<InvalidOperationException>(() =>
            provider.Open("https://files.example.test/share", writable: false));

        Assert.Contains("web page", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Commit_UploadsWorkbookWithOriginalEtag()
    {
        byte[]? uploaded = null;
        EntityTagHeaderValue? ifMatch = null;
        var handler = new RecordingHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                return WorkbookResponse();
            }

            uploaded = request.Content!.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            ifMatch = request.Headers.IfMatch.SingleOrDefault();
            return new HttpResponseMessage(HttpStatusCode.NoContent);
        });
        var provider = new WorkbookSourceProvider(new HttpClient(handler));

        using (var lease = provider.Open("https://files.example.test/recovery.xlsx", writable: true))
        {
            using (var workbook = new XLWorkbook(lease.LocalPath))
            {
                workbook.Worksheet("Import Mapping").Cell("A2").Value = "Updated";
                workbook.Save();
            }

            lease.Commit();
        }

        Assert.NotNull(uploaded);
        Assert.Equal("\"version-7\"", ifMatch?.Tag);
        using var uploadedStream = new MemoryStream(uploaded!);
        using var uploadedWorkbook = new XLWorkbook(uploadedStream);
        Assert.Equal("Updated", uploadedWorkbook.Worksheet("Import Mapping").Cell("A2").GetString());
    }

    [Fact]
    public void Commit_ExplainsDownloadOnlyUrl()
    {
        var handler = new RecordingHandler(request => request.Method == HttpMethod.Get
            ? WorkbookResponse()
            : new HttpResponseMessage(HttpStatusCode.MethodNotAllowed));
        var provider = new WorkbookSourceProvider(new HttpClient(handler));

        using var lease = provider.Open("https://files.example.test/recovery.xlsx", writable: true);
        var error = Assert.Throws<InvalidOperationException>(lease.Commit);

        Assert.Contains("does not permit upload", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static HttpResponseMessage WorkbookResponse()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Import Mapping").Cell("A1").Value = "Mapping Profile";
            workbook.SaveAs(stream);
        }

        var response = new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new ByteArrayContent(stream.ToArray())
        };
        response.Content.Headers.ContentType = MediaTypeHeaderValue.Parse(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
        {
            FileName = "recovery.xlsx"
        };
        response.Headers.ETag = new EntityTagHeaderValue("\"version-7\"");
        return response;
    }

    private sealed class RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> responder) : HttpMessageHandler
    {
        public List<HttpRequestMessage> Requests { get; } = [];

        protected override HttpResponseMessage Send(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            return responder(request);
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(Send(request, cancellationToken));
    }
}
