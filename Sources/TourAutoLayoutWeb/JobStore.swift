import Foundation
import TourAutoLayoutCore
import ZIPFoundation
import Vapor

struct AppConfig: Sendable {
    let port: Int
    let storageRoot: URL
    let maxConcurrentJobs: Int
    let maxUploadSizeMB: Int
    let publicBaseURL: String?
    let accessPassword: String?
    let accessSessionToken: String?
    let analyticsHost: String?
    let analyticsWebsiteId: String?

    var requiresPassword: Bool {
        !(accessPassword ?? "").isEmpty
    }

    static func fromEnvironment() throws -> AppConfig {
        try fromEnvironment(ProcessInfo.processInfo.environment)
    }

    static func fromEnvironment(_ environment: [String: String]) throws -> AppConfig {
        let port = Int(environment["PORT"] ?? "") ?? 8080
        let maxConcurrentJobs = max(1, Int(environment["MAX_CONCURRENT_JOBS"] ?? "") ?? 3)
        let maxUploadSizeMB = max(10, Int(environment["MAX_UPLOAD_SIZE_MB"] ?? "") ?? 80)
        let storageRoot = URL(fileURLWithPath: environment["STORAGE_ROOT"] ?? "storage", isDirectory: true)
        let publicBaseURL = environment["PUBLIC_BASE_URL"]
        let rawPassword = environment["ACCESS_PASSWORD"]?.trimmingCharacters(in: .whitespacesAndNewlines)
        let password = (rawPassword?.isEmpty == false) ? rawPassword : nil
        let analyticsHost = normalizedOptionalEnvironmentValue(environment["ANALYTICS_HOST"])
        let analyticsWebsiteId = normalizedOptionalEnvironmentValue(environment["ANALYTICS_WEBSITE_ID"])

        return AppConfig(
            port: port,
            storageRoot: storageRoot,
            maxConcurrentJobs: maxConcurrentJobs,
            maxUploadSizeMB: maxUploadSizeMB,
            publicBaseURL: publicBaseURL,
            accessPassword: password,
            accessSessionToken: password == nil ? nil : UUID().uuidString,
            analyticsHost: analyticsHost,
            analyticsWebsiteId: analyticsWebsiteId
        )
    }

    func prepareStorage() throws {
        let fileManager = FileManager.default
        try fileManager.createDirectory(at: storageRoot, withIntermediateDirectories: true)
        try fileManager.createDirectory(at: uploadsRoot, withIntermediateDirectories: true)
        try fileManager.createDirectory(at: jobsRoot, withIntermediateDirectories: true)
        try fileManager.createDirectory(at: outputsRoot, withIntermediateDirectories: true)
    }

    var uploadsRoot: URL {
        storageRoot.appendingPathComponent("uploads", isDirectory: true)
    }

    var jobsRoot: URL {
        storageRoot.appendingPathComponent("jobs", isDirectory: true)
    }

    var outputsRoot: URL {
        storageRoot.appendingPathComponent("outputs", isDirectory: true)
    }
}

struct UploadedAsset: Sendable {
    let filename: String
    let data: Data
}

struct CreateJobResponse: Content {
    let jobId: String
}

struct JobStatusResponse: Content {
    let jobId: String
    let status: String
    let progress: Double
    let createdAt: String
    let completedAt: String?
    let files: [JobFileResponse]
}

struct JobFileResponse: Content {
    let fileId: String
    let filename: String
    let status: String
    let warnings: [String]
    let errorMessage: String?
    let outputFilename: String?
    let durationSeconds: Double?
    let downloadURL: String?
}

struct AppConfigResponse: Content {
    let requiresPassword: Bool
    let authenticated: Bool
    let maxUploadSizeMB: Int
    let analyticsHost: String?
    let analyticsWebsiteId: String?
}

struct LoginRequest: Content {
    let password: String
}

private func normalizedOptionalEnvironmentValue(_ value: String?) -> String? {
    guard let trimmed = value?.trimmingCharacters(in: .whitespacesAndNewlines), !trimmed.isEmpty else {
        return nil
    }
    return trimmed
}

struct LoginResponse: Content {
    let ok: Bool
}

struct CreateJobForm: Content {
    var templateImage: File
    var documents: [File]

    enum CodingKeys: String, CodingKey {
        case templateImage
        case documents
        case documentsArray = "documents[]"
    }

    init(templateImage: File, documents: [File]) {
        self.templateImage = templateImage
        self.documents = documents
    }

    init(from decoder: Decoder) throws {
        let container = try decoder.container(keyedBy: CodingKeys.self)
        templateImage = try container.decode(File.self, forKey: .templateImage)
        if let decodedBracketArray = try? container.decode([File].self, forKey: .documentsArray), !decodedBracketArray.isEmpty {
            documents = decodedBracketArray
        } else if let decodedArray = try? container.decode([File].self, forKey: .documents), !decodedArray.isEmpty {
            documents = decodedArray
        } else if let singleBracket = try? container.decode(File.self, forKey: .documentsArray) {
            documents = [singleBracket]
        } else {
            documents = [try container.decode(File.self, forKey: .documents)]
        }
    }

    func encode(to encoder: Encoder) throws {
        var container = encoder.container(keyedBy: CodingKeys.self)
        try container.encode(templateImage, forKey: .templateImage)
        try container.encode(documents, forKey: .documents)
    }
}

actor JobStore {
    private let config: AppConfig
    private var jobs: [UUID: WebJobRecord] = [:]

    init(config: AppConfig) {
        self.config = config
    }

    func createJob(templateImage: UploadedAsset, documents: [UploadedAsset]) async throws -> UUID {
        guard !documents.isEmpty else {
            throw Abort(.badRequest, reason: "至少需要上传一份文档。")
        }

        let jobID = UUID()
        let jobRoot = config.jobsRoot.appendingPathComponent(jobID.uuidString, isDirectory: true)
        let inputsDirectory = jobRoot.appendingPathComponent("inputs", isDirectory: true)
        let outputsDirectory = jobRoot.appendingPathComponent("outputs", isDirectory: true)
        let templateDirectory = jobRoot.appendingPathComponent("template", isDirectory: true)

        try FileManager.default.createDirectory(at: inputsDirectory, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: outputsDirectory, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: templateDirectory, withIntermediateDirectories: true)

        let templateFilename = sanitizeFilename(templateImage.filename, fallbackBaseName: "template")
        let templateURL = templateDirectory.appendingPathComponent(templateFilename)
        try templateImage.data.write(to: templateURL)

        var reservedOutputNames = Set<String>()
        var files = [WebJobFileRecord]()

        for document in documents {
            let inputFilename = sanitizeFilename(document.filename, fallbackBaseName: "document")
            let inputURL = inputsDirectory.appendingPathComponent(inputFilename)
            try document.data.write(to: inputURL)

            let outputFilename = try makeOutputFilename(
                sourceFilename: inputFilename,
                reservedFilenames: &reservedOutputNames
            )
            let outputURL = outputsDirectory.appendingPathComponent(outputFilename)

            files.append(
                WebJobFileRecord(
                    id: UUID(),
                    sourceFilename: inputFilename,
                    sourceURL: inputURL,
                    status: .queued,
                    warnings: [],
                    errorMessage: nil,
                    outputFilename: outputFilename,
                    outputURL: outputURL,
                    durationSeconds: nil
                )
            )
        }

        jobs[jobID] = WebJobRecord(
            id: jobID,
            createdAt: Date(),
            completedAt: nil,
            templateFilename: templateFilename,
            templateURL: templateURL,
            files: files
        )

        Task.detached(priority: .userInitiated) { [self] in
            await processJob(id: jobID)
        }

        return jobID
    }

    func status(jobID: UUID) throws -> JobStatusResponse {
        guard let job = jobs[jobID] else {
            throw Abort(.notFound, reason: "任务不存在。")
        }
        return makeStatusResponse(for: job)
    }

    func fileDownload(jobID: UUID, fileID: UUID) throws -> DownloadPayload {
        guard let job = jobs[jobID] else {
            throw Abort(.notFound, reason: "任务不存在。")
        }
        guard let file = job.files.first(where: { $0.id == fileID }) else {
            throw Abort(.notFound, reason: "文件不存在。")
        }
        guard let outputFilename = file.outputFilename else {
            throw Abort(.notFound, reason: "输出文件不存在。")
        }
        guard FileManager.default.fileExists(atPath: file.outputURL.path) else {
            throw Abort(.notFound, reason: "结果尚未生成。")
        }

        let data = try Data(contentsOf: file.outputURL)
        return DownloadPayload(filename: outputFilename, contentType: docxContentType, data: data)
    }

    func archiveDownload(jobID: UUID) throws -> DownloadPayload {
        guard let job = jobs[jobID] else {
            throw Abort(.notFound, reason: "任务不存在。")
        }

        let readyFiles = job.files.filter {
            $0.status == .success || $0.status == .warning
        }
        guard !readyFiles.isEmpty else {
            throw Abort(.notFound, reason: "当前没有可下载的结果文件。")
        }

        let zipURL = config.outputsRoot.appendingPathComponent("\(job.id.uuidString).zip")
        if FileManager.default.fileExists(atPath: zipURL.path) {
            try FileManager.default.removeItem(at: zipURL)
        }

        let archive = try Archive(url: zipURL, accessMode: .create)
        for file in readyFiles {
            let data = try Data(contentsOf: file.outputURL)
            try archive.addEntry(
                with: file.outputFilename ?? "result.docx",
                type: .file,
                uncompressedSize: Int64(data.count),
                compressionMethod: .deflate
            ) { position, size in
                let start = Int(position)
                return data.subdata(in: start..<(start + size))
            }
        }

        let data = try Data(contentsOf: zipURL)
        return DownloadPayload(filename: "job-\(job.id.uuidString).zip", contentType: "application/zip", data: data)
    }

    private func processJob(id: UUID) async {
        guard let job = jobs[id] else { return }
        let workItems = job.files
        let templateData: Data
        do {
            templateData = try Data(contentsOf: job.templateURL)
        } catch {
            markWholeJobFailed(jobID: id, message: error.localizedDescription)
            return
        }

        for chunk in workItems.chunked(into: config.maxConcurrentJobs) {
            markProcessing(jobID: id, fileIDs: chunk.map(\.id))

            await withTaskGroup(of: FileOutcome.self) { group in
                for item in chunk {
                    let templateFilename = job.templateFilename
                    group.addTask {
                        let startedAt = Date()
                        do {
                            let template = TemplateProfile.goldenCruise(
                                backgroundImageData: templateData,
                                originalFilename: templateFilename
                            )
                            let pipeline = TourProcessingPipeline(
                                normalizer: DocumentNormalizer(converter: LibreOfficeDocumentConverter())
                            )
                            let result = try pipeline.process(
                                sourceURL: item.sourceURL,
                                template: template,
                                outputURL: item.outputURL
                            )
                            return FileOutcome(
                                fileID: item.id,
                                status: result.warnings.isEmpty ? .success : .warning,
                                warnings: result.warnings,
                                errorMessage: nil,
                                durationSeconds: Date().timeIntervalSince(startedAt)
                            )
                        } catch {
                            return FileOutcome(
                                fileID: item.id,
                                status: .failure,
                                warnings: [],
                                errorMessage: error.localizedDescription,
                                durationSeconds: Date().timeIntervalSince(startedAt)
                            )
                        }
                    }
                }

                for await outcome in group {
                    apply(outcome: outcome, to: id)
                }
            }
        }

        finish(jobID: id)
    }

    private func markProcessing(jobID: UUID, fileIDs: [UUID]) {
        guard var job = jobs[jobID] else { return }
        for index in job.files.indices where fileIDs.contains(job.files[index].id) {
            job.files[index].status = .processing
        }
        jobs[jobID] = job
    }

    private func apply(outcome: FileOutcome, to jobID: UUID) {
        guard var job = jobs[jobID] else { return }
        guard let index = job.files.firstIndex(where: { $0.id == outcome.fileID }) else { return }
        job.files[index].status = outcome.status
        job.files[index].warnings = outcome.warnings
        job.files[index].errorMessage = outcome.errorMessage
        job.files[index].durationSeconds = outcome.durationSeconds
        jobs[jobID] = job
    }

    private func finish(jobID: UUID) {
        guard var job = jobs[jobID] else { return }
        job.completedAt = Date()
        jobs[jobID] = job
    }

    private func markWholeJobFailed(jobID: UUID, message: String) {
        guard var job = jobs[jobID] else { return }
        for index in job.files.indices {
            job.files[index].status = .failure
            job.files[index].errorMessage = message
        }
        job.completedAt = Date()
        jobs[jobID] = job
    }

    private func makeStatusResponse(for job: WebJobRecord) -> JobStatusResponse {
        let completedCount = job.files.filter { status in
            switch status.status {
            case .success, .warning, .failure:
                return true
            case .queued, .processing:
                return false
            }
        }.count

        let progress = job.files.isEmpty ? 0 : Double(completedCount) / Double(job.files.count)
        let status: String
        if job.files.contains(where: { $0.status == .processing }) {
            status = "processing"
        } else if job.completedAt != nil {
            status = "completed"
        } else {
            status = "queued"
        }

        return JobStatusResponse(
            jobId: job.id.uuidString,
            status: status,
            progress: progress,
            createdAt: iso8601(job.createdAt),
            completedAt: job.completedAt.map(iso8601(_:)),
            files: job.files.map { file in
                JobFileResponse(
                    fileId: file.id.uuidString,
                    filename: file.sourceFilename,
                    status: file.status.rawValue,
                    warnings: file.warnings,
                    errorMessage: file.errorMessage,
                    outputFilename: file.outputFilename,
                    durationSeconds: file.durationSeconds,
                    downloadURL: (file.status == .success || file.status == .warning)
                        ? "/api/jobs/\(job.id.uuidString)/download/\(file.id.uuidString)"
                        : nil
                )
            }
        )
    }

    private func makeOutputFilename(sourceFilename: String, reservedFilenames: inout Set<String>) throws -> String {
        let baseName = URL(fileURLWithPath: sourceFilename).deletingPathExtension().lastPathComponent
        let initialName = "\(baseName)-自动排版.docx"

        if !reservedFilenames.contains(initialName) {
            reservedFilenames.insert(initialName)
            return initialName
        }

        for index in 2...999 {
            let candidate = "\(baseName)-自动排版-\(index).docx"
            if !reservedFilenames.contains(candidate) {
                reservedFilenames.insert(candidate)
                return candidate
            }
        }

        throw AppError.outputPathUnavailable(baseName)
    }

    private func sanitizeFilename(_ filename: String, fallbackBaseName: String) -> String {
        let candidate = filename.isEmpty ? fallbackBaseName : filename
        let cleaned = candidate.replacingOccurrences(
            of: #"[\\/:*?"<>|]+"#,
            with: "-",
            options: .regularExpression
        )
        return cleaned.isEmpty ? fallbackBaseName : cleaned
    }

    private func iso8601(_ date: Date) -> String {
        let formatter = ISO8601DateFormatter()
        return formatter.string(from: date)
    }

    private var docxContentType: String {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }
}

private struct WebJobRecord: Sendable {
    let id: UUID
    let createdAt: Date
    var completedAt: Date?
    let templateFilename: String
    let templateURL: URL
    var files: [WebJobFileRecord]
}

private struct WebJobFileRecord: Sendable {
    let id: UUID
    let sourceFilename: String
    let sourceURL: URL
    var status: ProcessingStatus
    var warnings: [String]
    var errorMessage: String?
    let outputFilename: String?
    let outputURL: URL
    var durationSeconds: Double?
}

private struct FileOutcome: Sendable {
    let fileID: UUID
    let status: ProcessingStatus
    let warnings: [String]
    let errorMessage: String?
    let durationSeconds: Double
}

struct DownloadPayload {
    let filename: String
    let contentType: String
    let data: Data
}

extension Array {
    fileprivate func chunked(into size: Int) -> [[Element]] {
        guard size > 0 else { return [self] }
        return stride(from: 0, to: count, by: size).map { start in
            Array(self[start..<Swift.min(start + size, count)])
        }
    }
}
