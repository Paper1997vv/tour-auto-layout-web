import Foundation

struct TourDocumentParser: DocumentParser {
    func parse(sourceDocument: SourceDocument) throws -> ImportedDocxSnapshot {
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
        let supportedTypes = Set([
            RelationshipTypes.image,
            RelationshipTypes.hyperlink,
        ])

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

        let unsupported = relationships.filter { relationship in
            !supportedTypes.contains(relationship.type) && !ignoredTypes.contains(relationship.type)
        }

        if !unsupported.isEmpty {
            warnings.append("部分高级引用未迁移：\(unsupported.count) 项。")
        }

        return relationships.filter { relationship in
            supportedTypes.contains(relationship.type)
        }
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
        bodyInnerXML.lastRegexCapture(
            pattern: #"(?s)(<w:sectPr\b.*?</w:sectPr>|<w:sectPr\b[^>]*/>)"#
        )
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
        guard let range = Range(captureRange, in: self) else { return nil }
        return String(self[range])
    }

    func lastRegexCapture(pattern: String) -> String? {
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let range = NSRange(startIndex..<endIndex, in: self)
        let matches = regex.matches(in: self, options: [], range: range)
        guard let match = matches.last else { return nil }

        let captureRange = match.numberOfRanges > 1 ? match.range(at: 1) : match.range(at: 0)
        guard let range = Range(captureRange, in: self) else { return nil }
        return String(self[range])
    }
}
