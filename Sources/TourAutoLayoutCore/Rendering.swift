import Foundation
import ZIPFoundation

public struct TourDocxRenderer: DocxRenderer {
    public init() {}

    public func render(document: ImportedDocxSnapshot, template: TemplateProfile, outputURL: URL) throws {
        let backgroundImage = try ImagePreparation.preparedImage(
            data: template.backgroundImage.data,
            originalFilename: template.backgroundImage.originalFilename
        )
        let builder = WordDocumentBuilder(template: template, document: document)
        let archiveDocument = try builder.buildArchive(backgroundImage: backgroundImage)

        if FileManager.default.fileExists(atPath: outputURL.path) {
            try FileManager.default.removeItem(at: outputURL)
        }
        try DocxArchiveWriter().write(archiveDocument, to: outputURL)
    }
}

private struct WordDocumentBuilder {
    private static let headerRelationshipID = "codexHeaderRel"
    private static let stylesRelationshipID = "codexStylesRel"
    private static let settingsRelationshipID = "codexSettingsRel"
    private static let numberingRelationshipID = "codexNumberingRel"
    private static let fontTableRelationshipID = "codexFontTableRel"
    private static let themeRelationshipID = "codexThemeRel"
    private static let headerImageRelationshipID = "codexHeaderImageRel"

    let template: TemplateProfile
    let document: ImportedDocxSnapshot

    func buildArchive(backgroundImage: PreparedImage) throws -> DocxArchiveDocument {
        let backgroundPart = ImportedMediaPart(
            path: "media/template-background.\(backgroundImage.fileExtension)",
            contentType: backgroundImage.contentType,
            data: backgroundImage.data
        )

        return DocxArchiveDocument(
            contentTypesXML: makeContentTypesXML(backgroundPart: backgroundPart),
            rootRelationshipsXML: makeRootRelationshipsXML(),
            documentXML: makeDocumentXML(bodyInnerXML: document.bodyInnerXML),
            documentRelationshipsXML: makeDocumentRelationshipsXML(),
            headerXML: makeHeaderXML(),
            headerRelationshipsXML: makeHeaderRelationshipsXML(backgroundPartPath: backgroundPart.path),
            stylesXML: document.stylesXML ?? makeFallbackStylesXML(),
            settingsXML: document.settingsXML ?? makeFallbackSettingsXML(),
            numberingXML: document.numberingXML,
            fontTableXML: document.fontTableXML,
            themeXML: document.themeXML,
            coreXML: makeCoreXML(),
            appXML: makeAppXML(),
            mediaParts: document.mediaParts + [backgroundPart]
        )
    }

    private func makeDocumentXML(bodyInnerXML: String) -> String {
        let rootTag = document.documentStartTag.isEmpty ? defaultDocumentRootTag : document.documentStartTag

        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \(rootTag)
          <w:body>
            \(bodyInnerXML)
            \(makeSectionPropertiesXML())
          </w:body>
        </w:document>
        """
    }

    private func makeSectionPropertiesXML() -> String {
        let pageLayout = template.pageLayout
        return """
        <w:sectPr>
          <w:headerReference r:id="\(Self.headerRelationshipID)" w:type="default"/>
          <w:pgSz w:w="\(pageLayout.pageWidthTwips)" w:h="\(pageLayout.pageHeightTwips)"/>
          <w:pgMar w:top="\(pageLayout.marginTopTwips)" w:right="\(pageLayout.marginRightTwips)" w:bottom="\(pageLayout.marginBottomTwips)" w:left="\(pageLayout.marginLeftTwips)" w:header="\(pageLayout.headerTwips)" w:footer="\(pageLayout.footerTwips)" w:gutter="0"/>
          <w:cols w:space="425" w:num="1"/>
          <w:docGrid w:type="lines" w:linePitch="360"/>
        </w:sectPr>
        """
    }

    private func makeHeaderXML() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
          <w:p>
            <w:r>
              <w:drawing>
                <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
                  <wp:simplePos x="0" y="0"/>
                  <wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>
                  <wp:positionV relativeFrom="page"><wp:align>top</wp:align></wp:positionV>
                  <wp:extent cx="7560000" cy="10692000"/>
                  <wp:wrapNone/>
                  <wp:docPr id="1" name="背景底图"/>
                  <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                  </wp:cNvGraphicFramePr>
                  <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:nvPicPr>
                          <pic:cNvPr id="1" name="背景底图"/>
                          <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
                        </pic:nvPicPr>
                        <pic:blipFill>
                          <a:blip r:embed="\(Self.headerImageRelationshipID)"/>
                          <a:stretch><a:fillRect/></a:stretch>
                        </pic:blipFill>
                        <pic:spPr>
                          <a:xfrm><a:off x="0" y="0"/><a:ext cx="7560000" cy="10692000"/></a:xfrm>
                          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        </pic:spPr>
                      </pic:pic>
                    </a:graphicData>
                  </a:graphic>
                </wp:anchor>
              </w:drawing>
            </w:r>
          </w:p>
        </w:hdr>
        """
    }

    private func makeHeaderRelationshipsXML(backgroundPartPath: String) -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="\(Self.headerImageRelationshipID)" Type="\(RelationshipTypes.image)" Target="\(backgroundPartPath)"/>
        </Relationships>
        """
    }

    private func makeDocumentRelationshipsXML() -> String {
        var relationships = [
            ImportedRelationship(id: Self.headerRelationshipID, type: RelationshipTypes.header, target: "header1.xml", targetMode: nil),
            ImportedRelationship(id: Self.stylesRelationshipID, type: RelationshipTypes.styles, target: "styles.xml", targetMode: nil),
            ImportedRelationship(id: Self.settingsRelationshipID, type: RelationshipTypes.settings, target: "settings.xml", targetMode: nil),
        ]

        if document.numberingXML != nil {
            relationships.append(ImportedRelationship(id: Self.numberingRelationshipID, type: RelationshipTypes.numbering, target: "numbering.xml", targetMode: nil))
        }
        if document.fontTableXML != nil {
            relationships.append(ImportedRelationship(id: Self.fontTableRelationshipID, type: RelationshipTypes.fontTable, target: "fontTable.xml", targetMode: nil))
        }
        if document.themeXML != nil {
            relationships.append(ImportedRelationship(id: Self.themeRelationshipID, type: RelationshipTypes.theme, target: "theme/theme1.xml", targetMode: nil))
        }

        relationships.append(contentsOf: document.relationships)

        let xml = relationships.map { relationship in
            var fragment = "<Relationship Id=\"\(relationship.id.xmlEscaped)\" Type=\"\(relationship.type.xmlEscaped)\" Target=\"\(relationship.target.xmlEscaped)\""
            if let targetMode = relationship.targetMode {
                fragment += " TargetMode=\"\(targetMode.xmlEscaped)\""
            }
            fragment += "/>"
            return fragment
        }.joined()

        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\(xml)</Relationships>
        """
    }

    private func makeRootRelationshipsXML() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
        """
    }

    private func makeContentTypesXML(backgroundPart: ImportedMediaPart) -> String {
        let imageDefaults = Set((document.mediaParts + [backgroundPart]).flatMap { mediaPart -> [String] in
            switch mediaPart.contentType {
            case "image/png":
                return ["<Default Extension=\"png\" ContentType=\"image/png\"/>"]
            case "image/jpeg":
                return [
                    "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>",
                    "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>",
                ]
            case "image/gif":
                return ["<Default Extension=\"gif\" ContentType=\"image/gif\"/>"]
            case "image/bmp":
                return ["<Default Extension=\"bmp\" ContentType=\"image/bmp\"/>"]
            case "image/tiff":
                return [
                    "<Default Extension=\"tif\" ContentType=\"image/tiff\"/>",
                    "<Default Extension=\"tiff\" ContentType=\"image/tiff\"/>",
                ]
            default:
                return []
            }
        }).joined()

        let numberingOverride = document.numberingXML == nil ? "" : "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>"
        let fontTableOverride = document.fontTableXML == nil ? "" : "<Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/>"
        let themeOverride = document.themeXML == nil ? "" : "<Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>"

        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          \(imageDefaults)
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
          <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
          \(numberingOverride)
          \(fontTableOverride)
          \(themeOverride)
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
        """
    }

    private func makeCoreXML() -> String {
        let formatter = ISO8601DateFormatter()
        let now = formatter.string(from: Date())
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
          <dc:title>旅游行程自动排版</dc:title>
          <dc:creator>Codex</dc:creator>
          <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
          <dcterms:created xsi:type="dcterms:W3CDTF">\(now)</dcterms:created>
          <dcterms:modified xsi:type="dcterms:W3CDTF">\(now)</dcterms:modified>
        </cp:coreProperties>
        """
    }

    private func makeAppXML() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
          <Application>TourAutoLayoutWeb</Application>
        </Properties>
        """
    }

    private func makeFallbackStylesXML() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
            <w:name w:val="Normal"/>
            <w:qFormat/>
          </w:style>
        </w:styles>
        """
    }

    private func makeFallbackSettingsXML() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:zoom w:percent="100"/>
          <w:defaultTabStop w:val="420"/>
          <w:compat/>
        </w:settings>
        """
    }

    private var defaultDocumentRootTag: String {
        """
        <w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
        """
    }
}

private enum ImagePreparation {
    static func preparedImage(data: Data, originalFilename: String) throws -> PreparedImage {
        let ext = URL(fileURLWithPath: originalFilename).pathExtension.lowercased()
        let contentType = PackageContentType.contentType(forFileAt: originalFilename)

        switch contentType {
        case "image/png":
            return PreparedImage(data: data, contentType: contentType, fileExtension: "png")
        case "image/jpeg":
            return PreparedImage(data: data, contentType: contentType, fileExtension: "jpg")
        case "image/gif":
            return PreparedImage(data: data, contentType: contentType, fileExtension: "gif")
        case "image/bmp":
            return PreparedImage(data: data, contentType: contentType, fileExtension: "bmp")
        case "image/tiff":
            return PreparedImage(data: data, contentType: contentType, fileExtension: ext == "tif" ? "tif" : "tiff")
        default:
            throw AppError.unsupportedTemplateImageFormat(ext.isEmpty ? "unknown" : ext)
        }
    }
}

private struct PreparedImage {
    let data: Data
    let contentType: String
    let fileExtension: String
}

private struct DocxArchiveDocument {
    let contentTypesXML: String
    let rootRelationshipsXML: String
    let documentXML: String
    let documentRelationshipsXML: String
    let headerXML: String
    let headerRelationshipsXML: String
    let stylesXML: String
    let settingsXML: String
    let numberingXML: String?
    let fontTableXML: String?
    let themeXML: String?
    let coreXML: String
    let appXML: String
    let mediaParts: [ImportedMediaPart]
}

private struct DocxArchiveWriter {
    func write(_ document: DocxArchiveDocument, to outputURL: URL) throws {
        let archive = try Archive(url: outputURL, accessMode: .create)

        try archive.addTextEntry(at: "[Content_Types].xml", text: document.contentTypesXML)
        try archive.addTextEntry(at: "_rels/.rels", text: document.rootRelationshipsXML)
        try archive.addTextEntry(at: "docProps/core.xml", text: document.coreXML)
        try archive.addTextEntry(at: "docProps/app.xml", text: document.appXML)
        try archive.addTextEntry(at: "word/document.xml", text: document.documentXML)
        try archive.addTextEntry(at: "word/_rels/document.xml.rels", text: document.documentRelationshipsXML)
        try archive.addTextEntry(at: "word/header1.xml", text: document.headerXML)
        try archive.addTextEntry(at: "word/_rels/header1.xml.rels", text: document.headerRelationshipsXML)
        try archive.addTextEntry(at: "word/styles.xml", text: document.stylesXML)
        try archive.addTextEntry(at: "word/settings.xml", text: document.settingsXML)

        if let numberingXML = document.numberingXML {
            try archive.addTextEntry(at: "word/numbering.xml", text: numberingXML)
        }
        if let fontTableXML = document.fontTableXML {
            try archive.addTextEntry(at: "word/fontTable.xml", text: fontTableXML)
        }
        if let themeXML = document.themeXML {
            try archive.addTextEntry(at: "word/theme/theme1.xml", text: themeXML)
        }

        for mediaPart in document.mediaParts {
            try archive.addDataEntry(at: "word/\(mediaPart.path)", data: mediaPart.data)
        }
    }
}

private extension Archive {
    func addTextEntry(at path: String, text: String) throws {
        try addDataEntry(at: path, data: Data(text.utf8))
    }

    func addDataEntry(at path: String, data: Data) throws {
        try addEntry(
            with: path,
            type: .file,
            uncompressedSize: Int64(data.count),
            compressionMethod: .deflate
        ) { position, size in
            let start = Int(position)
            return data.subdata(in: start..<(start + size))
        }
    }
}

private extension String {
    var xmlEscaped: String {
        self
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}
