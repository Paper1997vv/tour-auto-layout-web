import Foundation

enum SourceDocumentFormat: String, Sendable {
    case doc
    case docx
}

enum ProcessingStatus: String, Sendable {
    case queued
    case processing
    case success
    case warning
    case failure
}

struct PageLayout: Sendable {
    let pageWidthTwips: Int
    let pageHeightTwips: Int
    let marginTopTwips: Int
    let marginRightTwips: Int
    let marginBottomTwips: Int
    let marginLeftTwips: Int
    let headerTwips: Int
    let footerTwips: Int

    static let goldenCruiseA4 = PageLayout(
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

struct TemplateProfile: Identifiable, Sendable {
    let id: String
    let name: String
    let backgroundImageURL: URL
    let pageLayout: PageLayout

    static func goldenCruise(backgroundImageURL: URL) -> TemplateProfile {
        TemplateProfile(
            id: "golden_cruise",
            name: "黄金邮轮",
            backgroundImageURL: backgroundImageURL,
            pageLayout: .goldenCruiseA4
        )
    }
}

struct NormalizedDocument: Sendable {
    let originalURL: URL
    let normalizedDocxURL: URL
    let detectedFormat: SourceDocumentFormat
}

struct ImportedRelationship: Equatable, Sendable {
    let id: String
    let type: String
    let target: String
    let targetMode: String?
}

struct ImportedMediaPart: Equatable, Sendable {
    let path: String
    let contentType: String
    let data: Data
}

struct SourceDocument: Sendable {
    let originalURL: URL
    let normalizedDocxURL: URL
    let detectedFormat: SourceDocumentFormat
    let documentXML: String
    let relationships: [ImportedRelationship]
    let stylesXML: String?
    let numberingXML: String?
    let fontTableXML: String?
    let themeXML: String?
    let settingsXML: String?
    let mediaParts: [ImportedMediaPart]
}

struct ImportedDocxSnapshot: Sendable {
    let sourceURL: URL
    let detectedFormat: SourceDocumentFormat
    let documentStartTag: String
    let bodyInnerXML: String
    let sectionPropertiesXML: String?
    let relationships: [ImportedRelationship]
    let stylesXML: String?
    let numberingXML: String?
    let fontTableXML: String?
    let themeXML: String?
    let settingsXML: String?
    let mediaParts: [ImportedMediaPart]
    let warnings: [String]
}

struct ProcessingJob: Identifiable, Sendable {
    let id: UUID
    let sourceURL: URL
    var status: ProcessingStatus
    var warnings: [String]
    var outputURL: URL?
    var duration: TimeInterval?
    var errorMessage: String?

    init(
        id: UUID = UUID(),
        sourceURL: URL,
        status: ProcessingStatus = .queued,
        warnings: [String] = [],
        outputURL: URL? = nil,
        duration: TimeInterval? = nil,
        errorMessage: String? = nil
    ) {
        self.id = id
        self.sourceURL = sourceURL
        self.status = status
        self.warnings = warnings
        self.outputURL = outputURL
        self.duration = duration
        self.errorMessage = errorMessage
    }
}

protocol DocumentParser: Sendable {
    func parse(sourceDocument: SourceDocument) throws -> ImportedDocxSnapshot
}

protocol DocxRenderer: Sendable {
    func render(document: ImportedDocxSnapshot, template: TemplateProfile, outputURL: URL) throws
}

extension String {
    var isEmptyTrimmed: Bool {
        trimmingCharacters(in: .whitespacesAndNewlines).isEmpty
    }

    var collapsedWhitespace: String {
        replacingOccurrences(of: "\\s+", with: " ", options: .regularExpression)
            .trimmingCharacters(in: .whitespacesAndNewlines)
    }
}

extension Array where Element == ProcessingJob {
    var successCount: Int {
        filter { $0.status == .success || $0.status == .warning }.count
    }

    var failureCount: Int {
        filter { $0.status == .failure }.count
    }

    var warningCount: Int {
        filter { $0.status == .warning }.count
    }
}
