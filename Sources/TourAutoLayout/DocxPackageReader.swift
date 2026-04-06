import Foundation
import ZIPFoundation

#if canImport(FoundationXML)
import FoundationXML
#endif

struct DocxPackageReader {
    func loadSourceDocument(from normalizedDocument: NormalizedDocument) throws -> SourceDocument {
        let archive = try Archive(url: normalizedDocument.normalizedDocxURL, accessMode: .read)
        let documentXML = try archive.string(at: "word/document.xml")
        let relationshipsXML = try archive.optionalString(at: "word/_rels/document.xml.rels") ?? """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
        """

        let relationships = try RelationshipFileParser().parse(data: Data(relationshipsXML.utf8))
        let mediaParts = try loadMediaParts(from: archive, relationships: relationships)

        return SourceDocument(
            originalURL: normalizedDocument.originalURL,
            normalizedDocxURL: normalizedDocument.normalizedDocxURL,
            detectedFormat: normalizedDocument.detectedFormat,
            documentXML: documentXML,
            relationships: relationships,
            stylesXML: try archive.optionalString(at: "word/styles.xml"),
            numberingXML: try archive.optionalString(at: "word/numbering.xml"),
            fontTableXML: try archive.optionalString(at: "word/fontTable.xml"),
            themeXML: try archive.optionalString(at: "word/theme/theme1.xml"),
            settingsXML: try archive.optionalString(at: "word/settings.xml"),
            mediaParts: mediaParts
        )
    }

    private func loadMediaParts(from archive: Archive, relationships: [ImportedRelationship]) throws -> [ImportedMediaPart] {
        let imageTypes = Set([
            RelationshipTypes.image,
        ])

        return try relationships.compactMap { relationship in
            guard imageTypes.contains(relationship.type) else { return nil }
            guard relationship.targetMode == nil else { return nil }

            let target = relationship.target.replacingOccurrences(of: "\\", with: "/")
            guard target.hasPrefix("media/") else { return nil }
            guard let data = try archive.optionalData(at: "word/\(target)") else { return nil }

            return ImportedMediaPart(
                path: target,
                contentType: PackageContentType.contentType(forFileAt: target),
                data: data
            )
        }
    }
}

enum RelationshipTypes {
    static let header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    static let footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    static let styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    static let settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    static let numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    static let fontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    static let theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    static let image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    static let hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    static let customXML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
}

enum PackageContentType {
    static func contentType(forFileAt path: String) -> String {
        switch URL(fileURLWithPath: path).pathExtension.lowercased() {
        case "png":
            "image/png"
        case "jpg", "jpeg":
            "image/jpeg"
        case "gif":
            "image/gif"
        case "bmp":
            "image/bmp"
        case "tif", "tiff":
            "image/tiff"
        default:
            "application/octet-stream"
        }
    }
}

private final class RelationshipFileParser: NSObject, XMLParserDelegate {
    private var result = [ImportedRelationship]()

    func parse(data: Data) throws -> [ImportedRelationship] {
        let parser = XMLParser(data: data)
        parser.delegate = self
        guard parser.parse() else {
            throw parser.parserError ?? AppError.malformedDocument("关系文件解析失败")
        }
        return result
    }

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String: String] = [:]) {
        guard elementName == "Relationship" else { return }
        guard let id = attributeDict["Id"], let type = attributeDict["Type"], let target = attributeDict["Target"] else { return }

        result.append(
            ImportedRelationship(
                id: id,
                type: type,
                target: target,
                targetMode: attributeDict["TargetMode"]
            )
        )
    }
}

private extension Archive {
    func optionalData(at path: String) throws -> Data? {
        guard let entry = self[path] else { return nil }

        var data = Data()
        _ = try extract(entry) { chunk in
            data.append(chunk)
        }
        return data
    }

    func optionalString(at path: String) throws -> String? {
        guard let data = try optionalData(at: path) else { return nil }
        return String(decoding: data, as: UTF8.self)
    }

    func string(at path: String) throws -> String {
        guard let string = try optionalString(at: path) else {
            throw AppError.missingArchiveEntry(path)
        }
        return string
    }
}
