import AppKit
import Foundation
import UniformTypeIdentifiers

@MainActor
final class AppViewModel: ObservableObject {
    @Published var templateImageURL: URL?
    @Published var outputDirectoryURL: URL?
    @Published var jobs: [ProcessingJob] = []
    @Published var isProcessing = false
    @Published var progress = 0.0
    @Published var statusMessage = "请选择模板图、源文档和输出目录，然后开始生成。"

    private let parser: DocumentParser = TourDocumentParser()
    private let renderer: DocxRenderer = TourDocxRenderer()
    private let maxConcurrentJobs = 3

    func chooseTemplateImage() {
        let panel = NSOpenPanel()
        panel.title = "选择黄金邮轮模板底图"
        panel.canChooseFiles = true
        panel.canChooseDirectories = false
        panel.allowsMultipleSelection = false
        panel.allowedContentTypes = [.png, .jpeg, .gif, .tiff]

        guard panel.runModal() == .OK, let url = panel.url else { return }
        templateImageURL = url
    }

    func chooseSourceDocuments() {
        let panel = NSOpenPanel()
        panel.title = "选择旅游介绍文档"
        panel.canChooseFiles = true
        panel.canChooseDirectories = false
        panel.allowsMultipleSelection = true
        panel.allowedContentTypes = [
            UTType(filenameExtension: "doc") ?? .data,
            UTType(filenameExtension: "docx") ?? .data,
        ]

        guard panel.runModal() == .OK else { return }
        addDocuments(panel.urls)
    }

    func chooseOutputDirectory() {
        let panel = NSOpenPanel()
        panel.title = "选择输出目录"
        panel.canChooseFiles = false
        panel.canChooseDirectories = true
        panel.allowsMultipleSelection = false
        panel.canCreateDirectories = true

        guard panel.runModal() == .OK, let url = panel.url else { return }
        outputDirectoryURL = url
    }

    func addDocuments(_ urls: [URL]) {
        let filtered = urls.filter { url in
            let ext = url.pathExtension.lowercased()
            return ext == "doc" || ext == "docx"
        }

        let existing = Set(jobs.map(\.sourceURL))
        let newJobs = filtered
            .filter { !existing.contains($0) }
            .sorted { $0.lastPathComponent.localizedStandardCompare($1.lastPathComponent) == .orderedAscending }
            .map { ProcessingJob(sourceURL: $0) }

        guard !newJobs.isEmpty else { return }
        jobs.append(contentsOf: newJobs)
        statusMessage = "已加入 \(newJobs.count) 份文档，当前队列共 \(jobs.count) 份。"
    }

    func removeJob(_ job: ProcessingJob) {
        jobs.removeAll { $0.id == job.id }
    }

    func clearResults() {
        jobs.removeAll()
        progress = 0
        statusMessage = "队列已清空。"
    }

    func revealOutput(for job: ProcessingJob) {
        guard let url = job.outputURL else { return }
        NSWorkspace.shared.activateFileViewerSelecting([url])
    }

    func startProcessing() {
        guard !isProcessing else { return }
        guard let templateImageURL else {
            statusMessage = "请先选择模板底图。"
            return
        }
        guard let outputDirectoryURL else {
            statusMessage = "请先选择输出目录。"
            return
        }
        guard !jobs.isEmpty else {
            statusMessage = "请先添加待处理文档。"
            return
        }

        for index in jobs.indices {
            jobs[index].status = .queued
            jobs[index].warnings = []
            jobs[index].outputURL = nil
            jobs[index].duration = nil
            jobs[index].errorMessage = nil
        }

        isProcessing = true
        progress = 0
        statusMessage = "开始处理 \(jobs.count) 份文档。"

        let template = TemplateProfile.goldenCruise(backgroundImageURL: templateImageURL)

        do {
            let plannedJobs = try planOutputs(in: outputDirectoryURL)

            Task {
                defer { isProcessing = false }

                var completed = 0
                for chunk in plannedJobs.chunked(into: maxConcurrentJobs) {
                    for item in chunk {
                        jobs[item.index].status = .processing
                    }

                    await withTaskGroup(of: (Int, JobOutcome).self) { group in
                        for item in chunk {
                            let parser = parser
                            let renderer = renderer
                            group.addTask {
                                let outcome = Self.processJob(
                                    sourceURL: item.sourceURL,
                                    outputURL: item.outputURL,
                                    template: template,
                                    parser: parser,
                                    renderer: renderer
                                )
                                return (item.index, outcome)
                            }
                        }

                        for await (index, outcome) in group {
                            completed += 1
                            jobs[index].warnings = outcome.warnings
                            jobs[index].outputURL = outcome.outputURL
                            jobs[index].duration = outcome.duration
                            jobs[index].errorMessage = outcome.errorMessage
                            jobs[index].status = outcome.status
                            progress = Double(completed) / Double(max(plannedJobs.count, 1))

                            if let outputURL = outcome.outputURL, outcome.errorMessage == nil {
                                statusMessage = "\(outputURL.lastPathComponent) 已生成。"
                            } else if let errorMessage = outcome.errorMessage {
                                statusMessage = "\(jobs[index].sourceURL.lastPathComponent) 处理失败：\(errorMessage)"
                            }
                        }
                    }
                }

                statusMessage = "处理完成：成功 \(jobs.successCount)，警告 \(jobs.warningCount)，失败 \(jobs.failureCount)。"
            }
        } catch {
            statusMessage = error.localizedDescription
            isProcessing = false
        }
    }

    private func planOutputs(in outputDirectory: URL) throws -> [PlannedJob] {
        var reservedFilenames = Set<String>()
        return try jobs.enumerated().map { index, job in
            let outputURL = try makeOutputURL(for: job.sourceURL, in: outputDirectory, reservedFilenames: &reservedFilenames)
            return PlannedJob(index: index, sourceURL: job.sourceURL, outputURL: outputURL)
        }
    }

    private func makeOutputURL(for sourceURL: URL, in outputDirectory: URL, reservedFilenames: inout Set<String>) throws -> URL {
        let baseName = sourceURL.deletingPathExtension().lastPathComponent
        let initialURL = outputDirectory.appendingPathComponent("\(baseName)-自动排版.docx")
        if !FileManager.default.fileExists(atPath: initialURL.path), !reservedFilenames.contains(initialURL.lastPathComponent) {
            reservedFilenames.insert(initialURL.lastPathComponent)
            return initialURL
        }

        for index in 2...999 {
            let candidate = outputDirectory.appendingPathComponent("\(baseName)-自动排版-\(index).docx")
            if !FileManager.default.fileExists(atPath: candidate.path), !reservedFilenames.contains(candidate.lastPathComponent) {
                reservedFilenames.insert(candidate.lastPathComponent)
                return candidate
            }
        }

        throw AppError.outputPathUnavailable(baseName)
    }

    nonisolated private static func processJob(
        sourceURL: URL,
        outputURL: URL,
        template: TemplateProfile,
        parser: DocumentParser,
        renderer: DocxRenderer
    ) -> JobOutcome {
        let startedAt = Date()
        let normalizer = DocumentNormalizer()
        let packageReader = DocxPackageReader()

        do {
            let normalized = try normalizer.normalize(documentAt: sourceURL)
            let sourceDocument = try packageReader.loadSourceDocument(from: normalized)
            let importedDocument = try parser.parse(sourceDocument: sourceDocument)
            try renderer.render(document: importedDocument, template: template, outputURL: outputURL)

            let duration = Date().timeIntervalSince(startedAt)
            return JobOutcome(
                status: importedDocument.warnings.isEmpty ? .success : .warning,
                warnings: importedDocument.warnings,
                outputURL: outputURL,
                duration: duration,
                errorMessage: nil
            )
        } catch {
            return JobOutcome(
                status: .failure,
                warnings: [],
                outputURL: nil,
                duration: Date().timeIntervalSince(startedAt),
                errorMessage: error.localizedDescription
            )
        }
    }
}

private struct PlannedJob: Sendable {
    let index: Int
    let sourceURL: URL
    let outputURL: URL
}

private struct JobOutcome: Sendable {
    let status: ProcessingStatus
    let warnings: [String]
    let outputURL: URL?
    let duration: TimeInterval
    let errorMessage: String?
}

private extension Array {
    func chunked(into size: Int) -> [[Element]] {
        guard size > 0 else { return [self] }
        return stride(from: 0, to: count, by: size).map { start in
            Array(self[start..<Swift.min(start + size, count)])
        }
    }
}

enum AppError: LocalizedError {
    case unsupportedSourceFormat(String)
    case outputPathUnavailable(String)
    case missingArchiveEntry(String)
    case malformedDocument(String)
    case commandFailed(String)
    case imageEncodingFailed

    var errorDescription: String? {
        switch self {
        case .unsupportedSourceFormat(let ext):
            "不支持的文档格式：\(ext)"
        case .outputPathUnavailable(let name):
            "无法为 \(name) 分配输出文件名。"
        case .missingArchiveEntry(let path):
            "DOCX 缺少必要文件：\(path)"
        case .malformedDocument(let message):
            "文档内容无法解析：\(message)"
        case .commandFailed(let message):
            "系统转换命令失败：\(message)"
        case .imageEncodingFailed:
            "图片编码失败，无法写入 DOCX。"
        }
    }
}
