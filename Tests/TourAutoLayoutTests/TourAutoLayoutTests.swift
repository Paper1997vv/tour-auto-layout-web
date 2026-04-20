import Foundation
import Testing
import Vapor
import ZIPFoundation
@testable import TourAutoLayoutCore
@testable import TourAutoLayoutWeb

struct TourAutoLayoutTests {
    @Test
    func appConfigParsesAnalyticsEnvironment() throws {
        let config = try AppConfig.fromEnvironment([
            "PORT": "8088",
            "STORAGE_ROOT": "/tmp/tour-auto-layout-tests",
            "ANALYTICS_HOST": " https://umami.example.com/ ",
            "ANALYTICS_WEBSITE_ID": " website-123 ",
        ])

        #expect(config.port == 8088)
        #expect(config.storageRoot.path == "/tmp/tour-auto-layout-tests")
        #expect(config.analyticsHost == "https://umami.example.com/")
        #expect(config.analyticsWebsiteId == "website-123")
    }

    @Test
    func appConfigLeavesAnalyticsDisabledWhenUnset() throws {
        let config = try AppConfig.fromEnvironment([
            "PORT": "8080",
            "STORAGE_ROOT": "/tmp/tour-auto-layout-tests",
            "ANALYTICS_HOST": "   ",
            "ANALYTICS_WEBSITE_ID": "",
        ])

        #expect(config.analyticsHost == nil)
        #expect(config.analyticsWebsiteId == nil)
    }

    @Test
    func normalizerConvertsDocIntoDocx() throws {
        let tempDirectory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: tempDirectory, withIntermediateDirectories: true)

        let docURL = tempDirectory.appendingPathComponent("sample.doc")
        try Data("mock-doc".utf8).write(to: docURL)

        let normalizer = DocumentNormalizer(converter: MockDocumentConverter(outputData: try makeMinimalDocxData()))
        let normalized = try normalizer.normalize(documentAt: docURL)

        #expect(normalized.detectedFormat == .doc)
        #expect(normalized.normalizedDocxURL.pathExtension.lowercased() == "docx")
        #expect(FileManager.default.fileExists(atPath: normalized.normalizedDocxURL.path))
    }

    @Test
    func parserBuildsWholeDocumentSnapshot() throws {
        let fixtureURL = try makeFixtureDocx()
        let normalized = NormalizedDocument(originalURL: fixtureURL, normalizedDocxURL: fixtureURL, detectedFormat: .docx)

        let sourceDocument = try DocxPackageReader().loadSourceDocument(from: normalized)
        let snapshot = try TourDocumentParser().parse(sourceDocument: sourceDocument)

        #expect(snapshot.bodyInnerXML.contains("原始标题"))
        #expect(snapshot.bodyInnerXML.contains("原始正文第一段"))
        #expect(snapshot.bodyInnerXML.contains("<w:tbl>"))
        #expect(snapshot.documentStartTag.contains("xmlns:w"))
        #expect(snapshot.relationships.count == 2)
        #expect(snapshot.mediaParts.count == 1)
        #expect(snapshot.stylesXML?.contains("CustomHeading") == true)
        #expect(snapshot.numberingXML?.contains("w:abstractNum") == true)
        #expect(snapshot.settingsXML?.contains("w:zoom") == true)
    }

    @Test
    func rendererPreservesBodyAndAppliesTemplateLayout() throws {
        let sourceFixtureURL = try makeFixtureDocx()
        let normalized = NormalizedDocument(originalURL: sourceFixtureURL, normalizedDocxURL: sourceFixtureURL, detectedFormat: .docx)
        let sourceDocument = try DocxPackageReader().loadSourceDocument(from: normalized)
        let snapshot = try TourDocumentParser().parse(sourceDocument: sourceDocument)

        let tempDirectory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: tempDirectory, withIntermediateDirectories: true)

        let outputURL = tempDirectory.appendingPathComponent("result.docx")
        try TourDocxRenderer().render(
            document: snapshot,
            template: .goldenCruise(backgroundImageData: pngData, originalFilename: "background.png"),
            outputURL: outputURL
        )

        let archive = try Archive(url: outputURL, accessMode: .read)
        let documentXML = try archive.readString(at: "word/document.xml")
        let relationshipsXML = try archive.readString(at: "word/_rels/document.xml.rels")
        let stylesXML = try archive.readString(at: "word/styles.xml")
        let headerXML = try archive.readString(at: "word/header1.xml")

        #expect(documentXML.contains("原始标题"))
        #expect(documentXML.contains("原始正文第一段"))
        #expect(documentXML.contains("<w:tbl>"))
        #expect(documentXML.contains("r:embed=\"rId5\""))
        #expect(documentXML.contains("w:pgSz w:w=\"11906\" w:h=\"16838\""))
        #expect(documentXML.contains("w:pgMar w:top=\"2324\" w:right=\"567\" w:bottom=\"567\" w:left=\"567\""))
        #expect(relationshipsXML.contains("Id=\"codexHeaderRel\""))
        #expect(relationshipsXML.contains("Id=\"rId5\""))
        #expect(stylesXML.contains("CustomHeading"))
        #expect(headerXML.contains("cx=\"7560000\" cy=\"10692000\""))
        #expect(try archive.readData(at: "word/media/image1.png").isEmpty == false)
    }

    @Test
    func createJobFormDecodesBracketedMultiDocuments() throws {
        let boundary = "Boundary-\(UUID().uuidString)"
        let payload = multipartBody(
            boundary: boundary,
            parts: [
                MultipartPart(
                    name: "templateImage",
                    filename: "template.png",
                    contentType: "image/png",
                    data: pngData
                ),
                MultipartPart(
                    name: "documents[]",
                    filename: "sample1.docx",
                    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    data: try makeMinimalDocxData()
                ),
                MultipartPart(
                    name: "documents[]",
                    filename: "sample2.docx",
                    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    data: try makeMinimalDocxData()
                ),
            ]
        )

        let headers = HTTPHeaders([
            ("Content-Type", "multipart/form-data; boundary=\(boundary)"),
        ])
        let application = Application(.testing)
        defer { application.shutdown() }

        let request = Request(
            application: application,
            method: .POST,
            url: URI(path: "/api/jobs"),
            version: .init(major: 1, minor: 1),
            headers: headers,
            collectedBody: ByteBuffer(data: payload),
            on: application.eventLoopGroup.next()
        )

        let form = try request.content.decode(CreateJobForm.self)

        #expect(form.documents.count == 2)
        #expect(form.documents.map { $0.filename } == ["sample1.docx", "sample2.docx"])
    }

    private func makeFixtureDocx() throws -> URL {
        let tempDirectory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: tempDirectory, withIntermediateDirectories: true)
        let outputURL = tempDirectory.appendingPathComponent("fixture.docx")

        let archive = try Archive(url: outputURL, accessMode: .create)

        try archive.addString("""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="png" ContentType="image/png"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
          <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
          <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
        </Types>
        """, at: "[Content_Types].xml")

        try archive.addString("""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """, at: "_rels/.rels")

        try archive.addString("""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
          <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
        </Relationships>
        """, at: "word/_rels/document.xml.rels")

        try archive.addString(fixtureDocumentXML, at: "word/document.xml")
        try archive.addString("""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="CustomHeading">
            <w:name w:val="CustomHeading"/>
          </w:style>
        </w:styles>
        """, at: "word/styles.xml")

        try archive.addString("""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:zoom w:percent="120"/>
        </w:settings>
        """, at: "word/settings.xml")

        try archive.addString("""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="0"/>
        </w:numbering>
        """, at: "word/numbering.xml")

        try archive.addData(at: "word/media/image1.png", data: pngData)
        return outputURL
    }

    private var fixtureDocumentXML: String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr><w:pStyle w:val="CustomHeading"/></w:pPr>
              <w:r><w:t>原始标题</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>原始正文第一段</w:t></w:r></w:p>
            <w:tbl>
              <w:tr><w:tc><w:p><w:r><w:t>单元格A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>单元格B</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            <w:p>
              <w:r>
                <w:drawing>
                  <wp:inline>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic>
                          <pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
                        </pic:pic>
                      </a:graphicData>
                    </a:graphic>
                  </wp:inline>
                </w:drawing>
              </w:r>
            </w:p>
            <w:sectPr>
              <w:pgSz w:w="12240" w:h="15840"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            </w:sectPr>
          </w:body>
        </w:document>
        """
    }

    private var pngData: Data {
        Data(base64Encoded: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wn0yt8AAAAASUVORK5CYII=")!
    }
}

private struct MultipartPart {
    let name: String
    let filename: String
    let contentType: String
    let data: Data
}

private func multipartBody(boundary: String, parts: [MultipartPart]) -> Data {
    var body = Data()
    let lineBreak = "\r\n"

    for part in parts {
        body.append(Data("--\(boundary)\(lineBreak)".utf8))
        body.append(Data("Content-Disposition: form-data; name=\"\(part.name)\"; filename=\"\(part.filename)\"\(lineBreak)".utf8))
        body.append(Data("Content-Type: \(part.contentType)\(lineBreak)\(lineBreak)".utf8))
        body.append(part.data)
        body.append(Data(lineBreak.utf8))
    }

    body.append(Data("--\(boundary)--\(lineBreak)".utf8))
    return body
}

private struct MockDocumentConverter: DocumentConverting {
    let outputData: Data

    func convertDocToDocx(inputURL _: URL, outputURL: URL) throws {
        try outputData.write(to: outputURL)
    }
}

private extension Archive {
    func addString(_ string: String, at path: String) throws {
        try addData(at: path, data: Data(string.utf8))
    }

    func addData(at path: String, data: Data) throws {
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

    func readString(at path: String) throws -> String {
        let data = try readData(at: path)
        return String(decoding: data, as: UTF8.self)
    }

    func readData(at path: String) throws -> Data {
        guard let entry = self[path] else {
            throw AppError.missingArchiveEntry(path)
        }
        var data = Data()
        _ = try extract(entry) { chunk in
            data.append(chunk)
        }
        return data
    }
}

private func makeMinimalDocxData() throws -> Data {
    let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString).appendingPathExtension("docx")
    let archive = try Archive(url: tempURL, accessMode: .create)
    try archive.addString("""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>
    """, at: "[Content_Types].xml")
    try archive.addString("""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>
    """, at: "_rels/.rels")
    try archive.addString("""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body><w:p><w:r><w:t>stub</w:t></w:r></w:p></w:body>
    </w:document>
    """, at: "word/document.xml")
    let data = try Data(contentsOf: tempURL)
    try? FileManager.default.removeItem(at: tempURL)
    return data
}
