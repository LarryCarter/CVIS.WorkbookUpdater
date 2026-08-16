using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace CVIS.WorkbookUpdater.Services;

public interface IWorkbookSourceProvider
{
    WorkbookSourceLease Open(string source, bool writable);
}

public abstract class WorkbookSourceLease : IDisposable
{
    public abstract string LocalPath { get; }
    public virtual void Commit() { }
    public virtual string PreserveBackup(string backupPath) => backupPath;
    public virtual void Dispose() { }
}

public sealed class WorkbookSourceProvider : IWorkbookSourceProvider
{
    private readonly HttpClient _httpClient;

    public WorkbookSourceProvider()
        : this(new HttpClient(new HttpClientHandler
        {
            AllowAutoRedirect = true,
            UseDefaultCredentials = true
        }))
    {
    }

    public WorkbookSourceProvider(HttpClient httpClient) => _httpClient = httpClient;

    public WorkbookSourceLease Open(string source, bool writable)
    {
        if (string.IsNullOrWhiteSpace(source))
        {
            throw new InvalidOperationException("Select the CVIS recovery workbook first.");
        }

        if (!Uri.TryCreate(source.Trim(), UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            if (!File.Exists(source))
            {
                throw new FileNotFoundException("The workbook path does not exist.", source);
            }

            return new LocalWorkbookSourceLease(Path.GetFullPath(source));
        }

        return Download(uri, writable);
    }

    private WorkbookSourceLease Download(Uri uri, bool writable)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, uri);
        using var response = _httpClient.Send(request, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        var mediaType = response.Content.Headers.ContentType?.MediaType;
        if (mediaType?.Contains("html", StringComparison.OrdinalIgnoreCase) == true)
        {
            throw new InvalidOperationException(
                "The URL returned a web page instead of an Excel workbook. Use a direct file/download URL or a locally synchronized path.");
        }

        var workingDirectory = Path.Combine(Path.GetTempPath(), "CVIS.WorkbookUpdater", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workingDirectory);
        var fileName = FileName(response.Content.Headers.ContentDisposition, response.RequestMessage?.RequestUri ?? uri);
        var localPath = Path.Combine(workingDirectory, fileName);

        try
        {
            using (var target = File.Create(localPath))
            using (var content = response.Content.ReadAsStream())
            {
                content.CopyTo(target);
            }

            EnsureExcelPackage(localPath);
            return new HttpWorkbookSourceLease(
                _httpClient,
                response.RequestMessage?.RequestUri ?? uri,
                localPath,
                workingDirectory,
                writable,
                response.Headers.ETag);
        }
        catch
        {
            Directory.Delete(workingDirectory, recursive: true);
            throw;
        }
    }

    private static string FileName(ContentDispositionHeaderValue? disposition, Uri uri)
    {
        var proposed = (disposition?.FileNameStar ?? disposition?.FileName)?.Trim('"');
        if (string.IsNullOrWhiteSpace(proposed))
        {
            proposed = Path.GetFileName(uri.LocalPath);
        }

        if (string.IsNullOrWhiteSpace(proposed) || !proposed.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            proposed = "CVIS-Recovery-Workbook.xlsx";
        }

        return string.Concat(proposed.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries));
    }

    private static void EnsureExcelPackage(string path)
    {
        using var stream = File.OpenRead(path);
        if (stream.Length < 4 || stream.ReadByte() != 'P' || stream.ReadByte() != 'K')
        {
            throw new InvalidOperationException(
                "The URL did not return an .xlsx file. Use a direct workbook URL rather than a sharing or sign-in page.");
        }
    }

    private sealed class LocalWorkbookSourceLease(string localPath) : WorkbookSourceLease
    {
        public override string LocalPath { get; } = localPath;
    }

    private sealed class HttpWorkbookSourceLease(
        HttpClient httpClient,
        Uri sourceUri,
        string localPath,
        string workingDirectory,
        bool writable,
        EntityTagHeaderValue? entityTag) : WorkbookSourceLease
    {
        public override string LocalPath { get; } = localPath;

        public override void Commit()
        {
            if (!writable)
            {
                return;
            }

            using var request = new HttpRequestMessage(HttpMethod.Put, sourceUri)
            {
                Content = new StreamContent(File.OpenRead(LocalPath))
            };
            request.Content.Headers.ContentType = new MediaTypeHeaderValue(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            if (entityTag is not null)
            {
                request.Headers.IfMatch.Add(entityTag);
            }

            using var response = httpClient.Send(request);
            if (response.StatusCode is HttpStatusCode.MethodNotAllowed or HttpStatusCode.Forbidden)
            {
                throw new InvalidOperationException(
                    "The workbook was updated locally, but this URL does not permit upload. Use a writable WebDAV/direct file URL or a locally synchronized path.");
            }

            if (response.StatusCode == HttpStatusCode.PreconditionFailed)
            {
                throw new InvalidOperationException(
                    "The remote workbook changed after it was downloaded. Reload it before applying updates so another person's changes are not overwritten.");
            }

            response.EnsureSuccessStatusCode();
        }

        public override string PreserveBackup(string backupPath)
        {
            var backupDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "CVIS.WorkbookUpdater",
                "RemoteBackups");
            Directory.CreateDirectory(backupDirectory);
            var preserved = Path.Combine(
                backupDirectory,
                $"{Path.GetFileNameWithoutExtension(LocalPath)}.{DateTime.Now:yyyyMMddHHmmssfff}.xlsx");
            File.Copy(backupPath, preserved, overwrite: false);
            return preserved;
        }

        public override void Dispose()
        {
            if (Directory.Exists(workingDirectory))
            {
                Directory.Delete(workingDirectory, recursive: true);
            }
        }
    }
}
