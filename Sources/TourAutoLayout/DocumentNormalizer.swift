import Foundation

struct DocumentNormalizer {
    private let processRunner: ProcessRunning

    init(processRunner: ProcessRunning = ProcessRunner()) {
        self.processRunner = processRunner
    }

    func normalize(documentAt url: URL) throws -> NormalizedDocument {
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

            let outputURL = tempDirectory.appendingPathComponent(url.deletingPathExtension().lastPathComponent).appendingPathExtension("docx")
            try processRunner.run(
                executableURL: URL(fileURLWithPath: "/usr/bin/textutil"),
                arguments: [
                    "-convert", "docx",
                    "-output", outputURL.path,
                    url.path,
                ]
            )

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

protocol ProcessRunning {
    func run(executableURL: URL, arguments: [String]) throws
}

struct ProcessRunner: ProcessRunning {
    func run(executableURL: URL, arguments: [String]) throws {
        let process = Process()
        process.executableURL = executableURL
        process.arguments = arguments

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
