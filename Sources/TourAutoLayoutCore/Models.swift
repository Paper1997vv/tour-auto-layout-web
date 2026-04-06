import Foundation

public enum SourceDocumentFormat: String, Sendable, Codable {
    case doc
    case docx
}

public enum ProcessingStatus: String, Sendable, Codable {
    case queued
    case processing
    case success
    case warning
    case failure
}

public struct PageLayout: Sendable {
    public let pageWidthTwips: Int
    public let pageHeightTwips: Int
    public let marginTopTwips: Int
    public let marginRightTwips: Int
    public let marginBottomTwips: Int
    public let marginLeftTwips: Int
    public let headerTwips: Int
    public let footerTwips: Int

    public init(
        pageWidthTwips: Int,
        pageHeightTwips: Int,
        marginTopTwips: Int,
        marginRightTwips: Int,
        marginBottomTwips: Int,
        marginLeftTwips: Int,
        headerTwips: Int,
        footerTwips: Int
    ) {
        self.pageWidthTwips = pageWidthTwips
        self.pageHeightTwips = pageHeightTwips
        self.marginTopTwips = marginTopTwips
        self.marginRightTwips = marginRightTwips
        self.marginBottomTwips = marginBottomTwips
        self.marginLeftTwips = marginLeftTwips
        self.headerTwips = headerTwips
        self.footerTwips = footerTwips
    }

    public static let goldenCruiseA4 = PageLayout(
        pageWidthTwips: 11_906,
        pageHeightTwips: 16_838,
        marginTopTwips: 2_324,
        marginRightTwips: 567,
        marginBottomTwips: 567,
        marginLeftTwips: 567,
        headerTwips: 851,
        footerTwips: 567
    )
}

public struct TemplateAsset: Sendable {
    public let data: Data
    public let originalFilename: String

    public init(data: Data, originalFilename: String) {
        self.data = data
        self.originalFilename = originalFilename
    }
}

public struct TemplateProfile: Identifiable, Sendable {
    public let id: String
    public let name: String
    public let backgroundImage: TemplateAsset
    public let pageLayout: PageLayout

    public init(id: String, name: String, backgroundImage: TemplateAsset, pageLayout: PageLayout) {
        self.id = id
        self.name = name
        self.backgroundImage = backgroundImage
        self.pageLayout = pageLayout
    }

    public static func goldenCruise(backgroundImageData: Data, originalFilename: String) -> TemplateProfile {
        TemplateProfile(
            id: "golden_cruise",
            name: "黄金邮轮",
            backgroundImage: TemplateAsset(data: backgroundImageData, originalFilename: originalFilename),
            pageLayout: .goldenCruiseA4
        )
    }
}

public struct NormalizedDocument: Sendable {
    public let originalURL: URL
    public let normalizedDocxURL: URL
    public let detectedFormat: SourceDocumentFormat

    public init(originalURL: URL, normalizedDocxURL: URL, detectedFormat: SourceDocumentFormat) {
        self.originalURL = originalURL
        self.normalizedDocxURL = normalizedDocxURL
        self.detectedFormat = detectedFormat
    }
}

public struct ImportedRelationship: Equatable, Sendable {
    public let id: String
    public let type: String
    public let target: String
    public let targetMode: String?

    public init(id: String, type: String, target: String, targetMode: String?) {
        self.id = id
        self.type = type
        self.target = target
        self.targetMode = targetMode
    }
}

public struct ImportedMediaPart: Equatable, Sendable {
    public let path: String
    public let contentType: String
    public let data: Data

    public init(path: String, contentType: String, data: Data) {
        self.path = path
        self.contentType = contentType
        self.data = data
    }
}

public struct SourceDocument: Sendable {
    public let originalURL: URL
    public let normalizedDocxURL: URL
    public let detectedFormat: SourceDocumentFormat
    public let documentXML: String
    public let relationships: [ImportedRelationship]
    public let stylesXML: String?
    public let numberingXML: String?
    public let fontTableXML: String?
    public let themeXML: String?
    public let settingsXML: String?
    public let mediaParts: [ImportedMediaPart]

    public init(
        originalURL: URL,
        normalizedDocxURL: URL,
        detectedFormat: SourceDocumentFormat,
        documentXML: String,
        relationships: [ImportedRelationship],
        stylesXML: String?,
        numberingXML: String?,
        fontTableXML: String?,
        themeXML: String?,
        settingsXML: String?,
        mediaParts: [ImportedMediaPart]
    ) {
        self.originalURL = originalURL
        self.normalizedDocxURL = normalizedDocxURL
        self.detectedFormat = detectedFormat
        self.documentXML = documentXML
        self.relationships = relationships
        self.stylesXML = stylesXML
        self.numberingXML = numberingXML
        self.fontTableXML = fontTableXML
        self.themeXML = themeXML
        self.settingsXML = settingsXML
        self.mediaParts = mediaParts
    }
}

public struct ImportedDocxSnapshot: Sendable {
    public let sourceURL: URL
    public let detectedFormat: SourceDocumentFormat
    public let documentStartTag: String
    public let bodyInnerXML: String
    public let sectionPropertiesXML: String?
    public let relationships: [ImportedRelationship]
    public let stylesXML: String?
    public let numberingXML: String?
    public let fontTableXML: String?
    public let themeXML: String?
    public let settingsXML: String?
    public let mediaParts: [ImportedMediaPart]
    public let warnings: [String]

    public init(
        sourceURL: URL,
        detectedFormat: SourceDocumentFormat,
        documentStartTag: String,
        bodyInnerXML: String,
        sectionPropertiesXML: String?,
        relationships: [ImportedRelationship],
        stylesXML: String?,
        numberingXML: String?,
        fontTableXML: String?,
        themeXML: String?,
        settingsXML: String?,
        mediaParts: [ImportedMediaPart],
        warnings: [String]
    ) {
        self.sourceURL = sourceURL
        self.detectedFormat = detectedFormat
        self.documentStartTag = documentStartTag
        self.bodyInnerXML = bodyInnerXML
        self.sectionPropertiesXML = sectionPropertiesXML
        self.relationships = relationships
        self.stylesXML = stylesXML
        self.numberingXML = numberingXML
        self.fontTableXML = fontTableXML
        self.themeXML = themeXML
        self.settingsXML = settingsXML
        self.mediaParts = mediaParts
        self.warnings = warnings
    }
}

public protocol DocumentParser: Sendable {
    func parse(sourceDocument: SourceDocument) throws -> ImportedDocxSnapshot
}

public protocol DocxRenderer: Sendable {
    func render(document: ImportedDocxSnapshot, template: TemplateProfile, outputURL: URL) throws
}

public struct ProcessedDocumentResult: Sendable {
    public let outputURL: URL
    public let warnings: [String]

    public init(outputURL: URL, warnings: [String]) {
        self.outputURL = outputURL
        self.warnings = warnings
    }
}

public struct TourProcessingPipeline {
    public let normalizer: DocumentNormalizer
    public let packageReader: DocxPackageReader
    public let parser: any DocumentParser
    public let renderer: any DocxRenderer

    public init(
        normalizer: DocumentNormalizer = DocumentNormalizer(),
        packageReader: DocxPackageReader = DocxPackageReader(),
        parser: any DocumentParser = TourDocumentParser(),
        renderer: any DocxRenderer = TourDocxRenderer()
    ) {
        self.normalizer = normalizer
        self.packageReader = packageReader
        self.parser = parser
        self.renderer = renderer
    }

    public func process(sourceURL: URL, template: TemplateProfile, outputURL: URL) throws -> ProcessedDocumentResult {
        let normalized = try normalizer.normalize(documentAt: sourceURL)
        let sourceDocument = try packageReader.loadSourceDocument(from: normalized)
        let importedDocument = try parser.parse(sourceDocument: sourceDocument)
        try renderer.render(document: importedDocument, template: template, outputURL: outputURL)
        return ProcessedDocumentResult(outputURL: outputURL, warnings: importedDocument.warnings)
    }
}

public extension String {
    var isEmptyTrimmed: Bool {
        trimmingCharacters(in: .whitespacesAndNewlines).isEmpty
    }

    var collapsedWhitespace: String {
        replacingOccurrences(of: "\\s+", with: " ", options: .regularExpression)
            .trimmingCharacters(in: .whitespacesAndNewlines)
    }
}
