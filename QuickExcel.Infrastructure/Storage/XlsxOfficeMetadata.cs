using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

/// <summary>Validated, inert Excel metadata. Coordinate-bearing properties live in the domain model;
/// identifiers and producer metadata are copied only from the matching original workbook/sheet.</summary>
internal static class XlsxOfficeMetadata
{
    internal static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    internal static readonly XNamespace Ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private static readonly XNamespace X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
    private static readonly XNamespace X15Ac = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
    private static readonly XNamespace Revision = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private static readonly XNamespace Revision2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
    private static readonly XNamespace Revision6 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6";
    private static readonly XNamespace Revision10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
    private static readonly (string Name, IgnoredCellErrors Flag)[] ErrorFlags =
    [
        ("calculatedColumn", IgnoredCellErrors.CalculatedColumn), ("evaluationError", IgnoredCellErrors.EvaluationError),
        ("formula", IgnoredCellErrors.Formula), ("formulaRange", IgnoredCellErrors.FormulaRange),
        ("listDataValidation", IgnoredCellErrors.ListDataValidation), ("numberStoredAsText", IgnoredCellErrors.NumberStoredAsText),
        ("twoDigitTextYear", IgnoredCellErrors.TwoDigitTextYear), ("unlockedFormula", IgnoredCellErrors.UnlockedFormula),
        ("emptyCellReference", IgnoredCellErrors.EmptyCellReference)
    ];

    private static bool Attributes(XElement element, params XName[] names) =>
        element.Attributes().All(a => a.IsNamespaceDeclaration || names.Contains(a.Name));
    private static bool Scalar(XElement element, params XName[] names) => !element.HasElements &&
        string.IsNullOrWhiteSpace(element.Value) && Attributes(element, names);
    private static bool GuidValue(XAttribute attribute) => Guid.TryParse(attribute.Value, out _);
    private static XNamespace? Namespace(XElement element, string prefix)
    {
        try { System.Xml.XmlConvert.VerifyNCName(prefix); return element.GetNamespaceOfPrefix(prefix); }
        catch (Exception e) when (e is System.Xml.XmlException or ArgumentException) { return null; }
    }
    private static bool Ignorable(XElement element, string value) => value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
        .All(prefix => Namespace(element, prefix) is { NamespaceName.Length: > 0 });
    private static bool NoText(XElement element) => element.Nodes().OfType<XText>().All(t => string.IsNullOrWhiteSpace(t.Value));

    internal static bool RecognizesAttribute(XElement element, XAttribute attribute)
    {
        if (element.Name.Namespace != Main) return false;
        return element.Name.LocalName switch
        {
            "workbook" or "worksheet" when attribute.Name == Mc + "Ignorable" => Ignorable(element, attribute.Value),
            "workbookPr" when attribute.Name == "codeName" => attribute.Value.Length is > 0 and <= 1024,
            "worksheet" when attribute.Name == Revision + "uid" => GuidValue(attribute),
            "workbookView" when attribute.Name == Revision2 + "uid" => GuidValue(attribute),
            "row" or "sheetFormatPr" when attribute.Name == Ac + "dyDescent" => Descent(element) is not null,
            _ => false
        };
    }

    internal static bool RecognizesChild(XElement parent, XElement child)
    {
        // A recognized local name in another namespace must never bypass the loss guard.
        if (parent.Name == Main + "workbook")
        {
            if (parent.Elements(child.Name).Count() != 1) return false;
            if (child.Name == Main + "fileVersion") return Scalar(child, "appName", "lastEdited", "lowestEdited", "rupBuild", "codeName") &&
                (child.Attribute("codeName") is not { } code || GuidValue(code));
            if (child.Name == Mc + "AlternateContent") return AbsolutePathChoice(child);
            if (child.Name == Revision + "revisionPtr") return RevisionPointer(child);
        }
        if (parent.Name == Main + "worksheet" && parent.Elements(child.Name).Count() == 1)
        {
            if (child.Name == Main + "phoneticPr") return PhoneticShape(child);
            if (child.Name == Main + "ignoredErrors") return TryIgnoredErrors(child, out _);
        }
        return false;
    }

    private static bool AbsolutePathChoice(XElement element)
    {
        if (!Attributes(element) || !NoText(element) || element.Elements().Count() != 1) return false;
        var choice = element.Elements().Single();
        if (choice.Name != Mc + "Choice" || !NoText(choice) || !Attributes(choice, "Requires") || choice.Elements().Count() != 1) return false;
        var required = ((string?)choice.Attribute("Requires"))?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var path = choice.Elements().Single();
        return required is { Length: 1 } && Namespace(choice, required[0]) == X15 &&
            path.Name == X15Ac + "absPath" && Scalar(path, "url") && path.Attribute("url") is not null;
    }

    private static bool RevisionPointer(XElement element) => element.Attribute("revIDLastSave") is not null && element.Attribute("documentId") is not null && Scalar(element, "revIDLastSave", "documentId",
        Revision6 + "coauthVersionLast", Revision6 + "coauthVersionMax", Revision10 + "uidLastSave") &&
        element.Attributes().Where(a => !a.IsNamespaceDeclaration).All(a => a.Name == "documentId" ||
            (a.Name == Revision10 + "uidLastSave" ? GuidValue(a) : uint.TryParse(a.Value, NumberStyles.None, CultureInfo.InvariantCulture, out _)));

    internal static double? Descent(XElement? element) => element?.Attribute(Ac + "dyDescent") is { } value &&
        double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number) && number >= 0 ? number : null;

    private static bool PhoneticShape(XElement element) => Scalar(element, "fontId", "type", "alignment") &&
        uint.TryParse((string?)element.Attribute("fontId"), NumberStyles.None, CultureInfo.InvariantCulture, out _) &&
        (string?)element.Attribute("type") is null or "halfwidthKatakana" or "fullwidthKatakana" or "Hiragana" or "noConversion" &&
        (string?)element.Attribute("alignment") is null or "noControl" or "left" or "center" or "distributed";

    internal static WorksheetPhoneticProperties? ReadPhonetic(XlsxPackage package, string workbookPart, string sheetPart, XElement sheet)
    {
        if (sheet.Element(Main + "phoneticPr") is not { } phonetic || !RecognizesChild(sheet, phonetic)) return null;
        var stylesPath = package.RelatedByType(workbookPart, "styles");
        var styles = stylesPath is null ? null : package.XmlPart(stylesPath).Root;
        var fonts = styles?.Element(Main + "fonts")?.Elements(Main + "font").ToArray() ?? [];
        var index = uint.Parse(phonetic.Attribute("fontId")!.Value, CultureInfo.InvariantCulture);
        // OOXML specifies the Normal style's font when fontId is outside the font table.
        if (index >= fonts.Length)
        {
            var normal = styles?.Element(Main + "cellStyles")?.Elements(Main + "cellStyle").FirstOrDefault(e => (string?)e.Attribute("builtinId") == "0");
            var bases = styles?.Element(Main + "cellStyleXfs")?.Elements(Main + "xf").ToArray() ?? [];
            var baseIndex = (int?)normal?.Attribute("xfId") ?? 0;
            index = (uint?)bases.ElementAtOrDefault(baseIndex)?.Attribute("fontId") ?? 0;
        }
        var font = index < fonts.Length ? fonts[index] : stylesPath is null ? new XElement(Main + "font") : null;
        if (font is null || !ValidFont(font))
        {
            package.Loss(sheetPart + ": 拼音属性的字体不能安全保留。", "worksheet/phoneticPr", sheetPart);
            return null;
        }
        return new(CanonicalFont(font), (string?)phonetic.Attribute("type"), (string?)phonetic.Attribute("alignment"));
    }

    // Fonts are inert values. Reject relationships, unknown namespaces and unrecognized nested payloads.
    private static bool ValidFont(XElement font) => font.Name == Main + "font" && Attributes(font) &&
        font.Elements().Select(e => e.Name).Distinct().Count() == font.Elements().Count() && font.Elements().All(e =>
            e.Name.Namespace == Main && (e.Name.LocalName == "color" ? Scalar(e, "rgb", "indexed", "auto", "theme", "tint") :
                e.Name.LocalName is "name" or "sz" or "b" or "i" or "u" or "strike" or "outline" or "shadow" or "condense" or "extend" or "family" or "charset" or "scheme" or "vertAlign" && Scalar(e, "val")));
    private static string CanonicalFont(XElement font)
    {
        var copy = new XElement(font);
        copy.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        return copy.ToString(SaveOptions.DisableFormatting);
    }

    internal static ImmutableArray<WorksheetIgnoredError> ReadIgnoredErrors(XElement sheet) =>
        sheet.Element(Main + "ignoredErrors") is { } element && RecognizesChild(sheet, element) && TryIgnoredErrors(element, out var errors) ? errors : [];

    private static bool TryIgnoredErrors(XElement element, out ImmutableArray<WorksheetIgnoredError> result)
    {
        result = [];
        if (!Attributes(element) || !NoText(element)) return false;
        var list = new List<WorksheetIgnoredError>();
        foreach (var error in element.Elements())
        {
            if (error.Name != Main + "ignoredError" || !Scalar(error, ErrorFlags.Select(p => (XName)p.Name).Prepend("sqref").ToArray())) return false;
            var flags = IgnoredCellErrors.None;
            foreach (var (name, flag) in ErrorFlags)
            {
                var value = (string?)error.Attribute(name);
                if (value is not (null or "0" or "1" or "false" or "true")) return false;
                if (value is "1" or "true") flags |= flag;
            }
            var references = ((string?)error.Attribute("sqref"))?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (references is not { Length: > 0 } || references.Length + list.Count > 4096) return false;
            foreach (var reference in references)
            {
                try
                {
                    var endpoints = reference.Split(':');
                    if (endpoints.Length is < 1 or > 2) return false;
                    var first = TypedWorkbookXml.Position(endpoints[0]); var last = TypedWorkbookXml.Position(endpoints[^1]);
                    if (first.Row > last.Row || first.Column > last.Column) return false;
                    list.Add(new(new(first.Row, first.Column, last.Row, last.Column), flags));
                }
                catch (Exception e) when (e is FormatException or OverflowException or ArgumentException or InvalidDataException) { return false; }
            }
        }
        if (list.Count == 0) return false;
        result = list.ToImmutableArray(); return true;
    }

    internal static void WritePhonetic(WorksheetSnapshot sheet, XElement root, XElement styles)
    {
        if (sheet.Data.PhoneticProperties is not { } value) return;
        var font = Xml(System.Text.Encoding.UTF8.GetBytes(value.FontXml), new(default)).Root!;
        if (!ValidFont(font)) throw new InvalidDataException("拼音字体元数据包含未知结构。");
        var fonts = styles.Element(Main + "fonts") ?? throw new InvalidDataException("缺少字体表。");
        var candidates = fonts.Elements(Main + "font").ToArray();
        var index = Array.FindIndex(candidates, f => CanonicalFont(f) == CanonicalFont(font));
        if (index < 0)
        {
            if (candidates.Length >= 65536) throw new InvalidDataException("拼音字体表容量超限。");
            index = candidates.Length; fonts.Add(new XElement(font)); fonts.SetAttributeValue("count", index + 1);
        }
        root.Add(new XElement(Main + "phoneticPr", new XAttribute("fontId", index),
            value.Type is null ? null : new XAttribute("type", value.Type), value.Alignment is null ? null : new XAttribute("alignment", value.Alignment)));
    }

    internal static void WriteIgnoredErrors(WorksheetSnapshot sheet, XElement root)
    {
        if (sheet.Data.IgnoredErrors.IsEmpty) return;
        root.Add(new XElement(Main + "ignoredErrors", sheet.Data.IgnoredErrors.Select(error => new XElement(Main + "ignoredError",
            new XAttribute("sqref", error.Range.ToString()), ErrorFlags.Where(p => (error.Errors & p.Flag) != 0).Select(p => new XAttribute(p.Name, 1))))));
    }

    internal static void EnsureDescentNamespace(XElement root)
    {
        if (!root.DescendantsAndSelf().Attributes(Ac + "dyDescent").Any()) return;
        var prefix = BindNamespace(root, Ac, "x14ac");
        BindNamespace(root, Mc, "mc");
        var prefixes = ((string?)root.Attribute(Mc + "Ignorable") ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).ToList();
        if (!prefixes.Contains(prefix, StringComparer.Ordinal)) prefixes.Add(prefix);
        root.SetAttributeValue(Mc + "Ignorable", string.Join(" ", prefixes));
    }

    private static string BindNamespace(XElement root, XNamespace ns, string preferred)
    {
        if (root.GetPrefixOfNamespace(ns) is { Length: > 0 } existing) return existing;
        var prefix = preferred; var index = 0;
        while (root.GetNamespaceOfPrefix(prefix) is { } used && used != ns) prefix = preferred + ++index;
        root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName); return prefix;
    }

    internal sealed class Writer
    {
        private readonly WorkbookFileState? _state;
        private readonly XlsxPackage? _source;
        private readonly XElement? _workbook;
        internal Writer(WorkbookSnapshot snapshot, CancellationToken token)
        {
            _state = snapshot.FileState;
            if (_state is null) return;
            _source = new(_state.OriginalWorkbook.ToArray(), new(token));
            _workbook = _source.XmlPart(_source.RelatedByType("", "officeDocument", true)!).Root!;
        }

        // Bindings are by original stable SheetId, never by mutable name or current ordinal.
        internal void WriteSheet(WorksheetSnapshot sheet, XElement root)
        {
            if (_state is not null && _source is not null)
            {
                var index = Array.FindIndex(_state.OriginalModel.Sheets.ToArray(), s => s.Id == sheet.Id);
                if (index >= 0 && index < _state.SheetParts.Length)
                {
                    var source = _source.XmlPart(_state.SheetParts[index].Part).Root!;
                    CopyIgnorable(source, root);
                    CopyAttribute(source, root, Revision + "uid");
                }
            }
            EnsureDescentNamespace(root);
        }

        internal void WriteWorkbook(XElement target)
        {
            if (_workbook is null) return;
            CopyIgnorable(_workbook, target);
            if (_workbook.Element(Main + "workbookPr") is { } properties)
                CopyAttribute(properties, target.Element(Main + "workbookPr")!, "codeName");
            foreach (var element in _workbook.Elements().Where(e => RecognizesChild(_workbook, e)).Reverse())
            {
                var copy = new XElement(element);
                // MC Requires contains lexical prefixes, so retain the source namespace context.
                CopyNamespaces(_workbook, copy);
                if (element.Name == Main + "fileVersion") target.AddFirst(copy);
                else target.Element(Main + "workbookPr")!.AddAfterSelf(copy);
            }
            if (_workbook.Element(Main + "bookViews")?.Elements(Main + "workbookView").ToArray() is { Length: 1 } views &&
                target.Element(Main + "bookViews")?.Element(Main + "workbookView") is { } view)
                CopyAttribute(views[0], view, Revision2 + "uid");
        }

        private static void CopyAttribute(XElement source, XElement target, XName name)
        {
            if (source.Attribute(name) is not { } attribute || !RecognizesAttribute(source, attribute)) return;
            if (name.NamespaceName.Length > 0) BindNamespace(target, name.Namespace, source.GetPrefixOfNamespace(name.Namespace) ?? "ext");
            target.SetAttributeValue(name, attribute.Value);
        }
        private static void CopyIgnorable(XElement source, XElement target)
        {
            if (source.Attribute(Mc + "Ignorable") is not { } attribute || !RecognizesAttribute(source, attribute)) return;
            var prefixes = new List<string>();
            foreach (var prefix in attribute.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                prefixes.Add(BindNamespace(target, source.GetNamespaceOfPrefix(prefix)!, prefix));
            BindNamespace(target, Mc, "mc");
            target.SetAttributeValue(Mc + "Ignorable", string.Join(" ", prefixes));
        }
        private static void CopyNamespaces(XElement source, XElement target)
        {
            foreach (var attribute in source.Attributes().Where(a => a.IsNamespaceDeclaration && a.Name.LocalName != "xmlns"))
                if (target.Attribute(attribute.Name) is null) target.Add(new XAttribute(attribute));
        }
    }
}
