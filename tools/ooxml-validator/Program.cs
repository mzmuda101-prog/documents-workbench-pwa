// ooxml-validator — sprawdza pliki .docx ze schematem Office (DocumentFormat.OpenXml, OpenXmlValidator).
// Łapie to, czego nie widać po samym poprawnym XML: elementy w złej kolejności (np. w:rPr),
// nieznane atrybuty, złe wartości — przez takie Word pokazuje „nieczytelna zawartość”.
//
//   dotnet run --project tools/ooxml-validator -- plik1.docx [plik2.docx …]
//   (albo zbudowany: tools/ooxml-validator/bin/Release/net10.0/OoxmlValidator plik.docx)
//
// Wynik: JSON na standardowe wyjście — [{ file, ok, error?, errors: [{ id, type, part, path, description }] }].
// Kod wyjścia: 0 — wszystkie bez błędów, 1 — są błędy schematu, 2 — pliku nie da się otworzyć.

using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var results = new List<object>();
var worst = 0;
// najnowszy zestaw schematów (w14/w15/w16… — rozszerzenia nowych Wordów nie są błędem)
var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365) { MaxNumberOfErrors = 500 };

foreach (var file in args)
{
    try
    {
        using var doc = WordprocessingDocument.Open(file, false);
        var errors = validator.Validate(doc).Select(e => new
        {
            id = e.Id,
            type = e.ErrorType.ToString(),
            part = e.Part?.Uri.ToString(),
            path = e.Path?.XPath,
            description = e.Description,
        }).ToList();
        if (errors.Count > 0) worst = Math.Max(worst, 1);
        results.Add(new { file, ok = errors.Count == 0, errors });
    }
    catch (Exception ex)
    {
        worst = 2;
        results.Add(new { file, ok = false, error = $"{ex.GetType().Name}: {ex.Message}", errors = Array.Empty<object>() });
    }
}

Console.WriteLine(JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = false }));
return worst;
