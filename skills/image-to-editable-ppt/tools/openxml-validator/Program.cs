using System.Reflection;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length != 1)
{
    Console.Error.WriteLine("usage: openxml-validator <presentation.pptx>");
    return 64;
}

var path = Path.GetFullPath(args[0]);
var errors = new List<object>();
string? openError = null;
try
{
    using var document = PresentationDocument.Open(path, false);
    var validator = new OpenXmlValidator();
    foreach (var error in validator.Validate(document))
    {
        errors.Add(new
        {
            error_type = error.ErrorType.ToString(),
            part = error.Part?.Uri.ToString() ?? "",
            xpath = error.Path?.XPath ?? "",
            description = error.Description,
        });
    }
}
catch (Exception exception)
{
    openError = $"{exception.GetType().FullName}: {exception.Message}";
}

var version = typeof(PresentationDocument).Assembly
    .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
    ?? typeof(PresentationDocument).Assembly.GetName().Version?.ToString()
    ?? "unknown";
var payload = new
{
    schema_version = 1,
    validator = "DocumentFormat.OpenXml.OpenXmlValidator",
    validator_version = version,
    path,
    open_error = openError,
    error_count = errors.Count,
    errors,
    passed = openError is null && errors.Count == 0,
};
Console.WriteLine(JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }));
return openError is not null ? 2 : errors.Count == 0 ? 0 : 1;
