import Foundation
import ZIPFoundation

public protocol ProcessRunning: Sendable {
    func run(executableURL: URL, arguments: [String], currentDirectoryURL: URL?) throws
}

public struct ProcessRunner: ProcessRunning {
    public init() {}

    public func run(executableURL: URL, arguments: [String], currentDirectoryURL: URL? = nil) throws {
        let process = Process()
        process.executableURL = executableURL
        process.arguments = arguments
        process.currentDirectoryURL = currentDirectoryURL

        let stdout = Pipe()
        let stderr = Pipe()
        process.standardOutput = stdout
        process.standardError = stderr

        try process.run()
        process.waitUntilExit()

        guard process.terminationStatus == 0 else {
            let errorData = stderr.fileHandleForReading.readDataToEndOfFile()
            let errorMessage = String(data: errorData, encoding: .utf8)?.collapsedWhitespace ?? "未知错误"
            throw AppError.commandFailed(errorMessage)
        }
    }
}

public protocol DocumentConverting: Sendable {
    func convertDocToDocx(inputURL: URL, outputURL: URL) throws
}

public struct LibreOfficeDocumentConverter: DocumentConverting {
    private let processRunner: any ProcessRunning
    private let executableURL: URL

    public init(processRunner: any ProcessRunning = ProcessRunner(), executableURL: URL? = nil) {
        self.processRunner = processRunner
        self.executableURL = executableURL ?? Self.resolveExecutableURL()
    }

    public func convertDocToDocx(inputURL: URL, outputURL: URL) throws {
        let outputDirectory = outputURL.deletingLastPathComponent()
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)
        try processRunner.run(
            executableURL: executableURL,
            arguments: [
                "--headless",
                "--convert-to", "docx",
                "--outdir", outputDirectory.path,
                inputURL.path,
            ],
            currentDirectoryURL: outputDirectory
        )

        guard FileManager.default.fileExists(atPath: outputURL.path) else {
            throw AppError.commandFailed("LibreOffice 未生成预期的 docx 文件。")
        }
    }

    private static func resolveExecutableURL() -> URL {
        if let configured = ProcessInfo.processInfo.environment["LIBREOFFICE_PATH"], !configured.isEmpty {
            return URL(fileURLWithPath: configured)
        }

        let candidates = [
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/snap/bin/libreoffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]

        for candidate in candidates where FileManager.default.isExecutableFile(atPath: candidate) {
            return URL(fileURLWithPath: candidate)
        }

        return URL(fileURLWithPath: "/usr/bin/libreoffice")
    }
}

public struct DocumentNormalizer {
    private let converter: any DocumentConverting

    public init(converter: any DocumentConverting = LibreOfficeDocumentConverter()) {
        self.converter = converter
    }

    public func normalize(documentAt url: URL) throws -> NormalizedDocument {
        let ext = url.pathExtension.lowercased()

        switch ext {
        case "docx":
            return NormalizedDocument(
                originalURL: url,
                normalizedDocxURL: url,
                detectedFormat: .docx
            )
        case "doc":
            let tempDirectory = FileManager.default.temporaryDirectory
                .appendingPathComponent("TourAutoLayout", isDirectory: true)
                .appendingPathComponent(UUID().uuidString, isDirectory: true)
            try FileManager.default.createDirectory(at: tempDirectory, withIntermediateDirectories: true)

            let outputURL = tempDirectory
                .appendingPathComponent(url.deletingPathExtension().lastPathComponent)
                .appendingPathExtension("docx")
            try converter.convertDocToDocx(inputURL: url, outputURL: outputURL)

            return NormalizedDocument(
                originalURL: url,
                normalizedDocxURL: outputURL,
                detectedFormat: .doc
            )
        default:
            throw AppError.unsupportedSourceFormat(ext)
        }
    }
}

public struct DocxPackageReader {
    public init() {}

    public func loadSourceDocument(from normalizedDocument: NormalizedDocument) throws -> SourceDocument {
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
        let imageTypes = Set([RelationshipTypes.image])

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

public enum RelationshipTypes {
    public static let header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    public static let footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    public static let styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    public static let settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    public static let numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    public static let fontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    public static let theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    public static let image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    public static let hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    public static let customXML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
}

public enum PackageContentType {
    public static func contentType(forFileAt path: String) -> String {
        switch URL(fileURLWithPath: path).pathExtension.lowercased() {
        case "png":
            return "image/png"
        case "jpg", "jpeg":
            return "image/jpeg"
        case "gif":
            return "image/gif"
        case "bmp":
            return "image/bmp"
        case "tif", "tiff":
            return "image/tiff"
        default:
            return "application/octet-stream"
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

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName _: String?, attributes attributeDict: [String: String] = [:]) {
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

public struct TourDocumentParser: DocumentParser {
    public init() {}

    public func parse(sourceDocument: SourceDocument) throws -> ImportedDocxSnapshot {
        let documentStartTag = try extractDocumentStartTag(from: sourceDocument.documentXML)
        let rawBodyInnerXML = try extractBodyInnerXML(from: sourceDocument.documentXML)
        let extractedSectionProperties = extractSectionProperties(from: rawBodyInnerXML)
        let cleanedBodyInnerXML = removeSectionProperties(from: rawBodyInnerXML)

        var warnings = [String]()
        if sourceDocument.detectedFormat == .doc {
            warnings.append("源文件来自 .doc 转换，格式保真度可能低于原生 .docx。")
        }
        if cleanedBodyInnerXML.isEmptyTrimmed {
            warnings.append("正文为空，输出文档仅会包含模板背景。")
        }

        let relationships = keepableRelationships(from: sourceDocument.relationships, warnings: &warnings)

        return ImportedDocxSnapshot(
            sourceURL: sourceDocument.originalURL,
            detectedFormat: sourceDocument.detectedFormat,
            documentStartTag: documentStartTag,
            bodyInnerXML: cleanedBodyInnerXML,
            sectionPropertiesXML: extractedSectionProperties,
            relationships: relationships,
            stylesXML: sourceDocument.stylesXML,
            numberingXML: sourceDocument.numberingXML,
            fontTableXML: sourceDocument.fontTableXML,
            themeXML: sourceDocument.themeXML,
            settingsXML: sourceDocument.settingsXML,
            mediaParts: sourceDocument.mediaParts,
            warnings: warnings
        )
    }

    private func keepableRelationships(from relationships: [ImportedRelationship], warnings: inout [String]) -> [ImportedRelationship] {
        let supportedTypes = Set([RelationshipTypes.image, RelationshipTypes.hyperlink])
        let ignoredTypes = Set([
            RelationshipTypes.header,
            RelationshipTypes.footer,
            RelationshipTypes.styles,
            RelationshipTypes.settings,
            RelationshipTypes.numbering,
            RelationshipTypes.fontTable,
            RelationshipTypes.theme,
            RelationshipTypes.customXML,
        ])

        let unsupported = relationships.filter { !supportedTypes.contains($0.type) && !ignoredTypes.contains($0.type) }
        if !unsupported.isEmpty {
            warnings.append("部分高级引用未迁移：\(unsupported.count) 项。")
        }

        return relationships.filter { supportedTypes.contains($0.type) }
    }

    private func extractDocumentStartTag(from xml: String) throws -> String {
        if let tag = xml.firstRegexCapture(pattern: #"(?s)<w:document\b[^>]*>"#) {
            return tag
        }
        throw AppError.malformedDocument("缺少 w:document 根节点")
    }

    private func extractBodyInnerXML(from xml: String) throws -> String {
        if let body = xml.firstRegexCapture(pattern: #"(?s)<w:body\b[^>]*>(.*)</w:body>"#) {
            return body
        }
        throw AppError.malformedDocument("缺少 w:body 节点")
    }

    private func extractSectionProperties(from bodyInnerXML: String) -> String? {
        bodyInnerXML.lastRegexCapture(pattern: #"(?s)(<w:sectPr\b.*?</w:sectPr>|<w:sectPr\b[^>]*/>)"#)
    }

    private func removeSectionProperties(from bodyInnerXML: String) -> String {
        bodyInnerXML.replacingOccurrences(
            of: #"(?s)<w:sectPr\b.*?</w:sectPr>|<w:sectPr\b[^>]*/>"#,
            with: "",
            options: .regularExpression
        )
    }
}

private extension String {
    func firstRegexCapture(pattern: String) -> String? {
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let range = NSRange(startIndex..<endIndex, in: self)
        guard let match = regex.firstMatch(in: self, options: [], range: range) else { return nil }

        let captureRange = match.numberOfRanges > 1 ? match.range(at: 1) : match.range(at: 0)
        guard let resolvedRange = Range(captureRange, in: self) else { return nil }
        return String(self[resolvedRange])
    }

    func lastRegexCapture(pattern: String) -> String? {
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let range = NSRange(startIndex..<endIndex, in: self)
        let matches = regex.matches(in: self, options: [], range: range)
        guard let match = matches.last else { return nil }

        let captureRange = match.numberOfRanges > 1 ? match.range(at: 1) : match.range(at: 0)
        guard let resolvedRange = Range(captureRange, in: self) else { return nil }
        return String(self[resolvedRange])
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
        guard let value = try optionalString(at: path) else {
            throw AppError.missingArchiveEntry(path)
        }
        return value
    }
}
