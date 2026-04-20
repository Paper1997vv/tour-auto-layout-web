import Foundation
import MultipartKit
import Vapor

struct AccessPasswordMiddleware: AsyncMiddleware {
    let config: AppConfig

    func respond(to request: Request, chainingTo next: AsyncResponder) async throws -> Response {
        guard config.requiresPassword else {
            return try await next.respond(to: request)
        }

        guard request.cookies["tour_access"]?.string == config.accessSessionToken else {
            throw Abort(.unauthorized, reason: "需要登录后才能继续使用。")
        }

        return try await next.respond(to: request)
    }
}

func routes(_ app: Application) throws {
    app.get { request in
        request.redirect(to: "/index.html")
    }

    app.get("api", "config") { request async throws -> AppConfigResponse in
        AppConfigResponse(
            requiresPassword: request.application.appConfig.requiresPassword,
            authenticated: request.application.appConfig.accessSessionToken == nil
                || request.cookies["tour_access"]?.string == request.application.appConfig.accessSessionToken,
            maxUploadSizeMB: request.application.appConfig.maxUploadSizeMB
        )
    }

    app.on(.POST, "api", "auth", "login", body: .collect(maxSize: "10kb")) { request async throws -> Response in
        let payload = try request.content.decode(LoginRequest.self)
        let config = request.application.appConfig

        guard let password = config.accessPassword, !password.isEmpty else {
            return Response(status: .ok, body: .init(string: #"{"ok":true}"#))
        }

        guard payload.password == password, let token = config.accessSessionToken else {
            throw Abort(.unauthorized, reason: "密码错误。")
        }

        let response = Response(status: .ok, body: .init(string: #"{"ok":true}"#))
        response.headers.replaceOrAdd(
            name: .setCookie,
            value: "tour_access=\(token); Path=/; HttpOnly; SameSite=Lax"
        )
        response.headers.replaceOrAdd(name: .contentType, value: "application/json; charset=utf-8")
        return response
    }

    let protected = app.grouped(AccessPasswordMiddleware(config: app.appConfig))

    protected.on(.POST, "api", "jobs", body: .collect(maxSize: "200mb")) { request async throws -> CreateJobResponse in
        let payload = try decodeCreateJobForm(from: request)
        try validate(templateFile: payload.templateImage, documents: payload.documents, config: request.application.appConfig)

        let jobID = try await request.application.jobStore.createJob(
            templateImage: UploadedAsset(filename: payload.templateImage.filename, data: Data(payload.templateImage.data.readableBytesView)),
            documents: payload.documents.map { UploadedAsset(filename: $0.filename, data: Data($0.data.readableBytesView)) }
        )
        return CreateJobResponse(jobId: jobID.uuidString)
    }

    protected.get("api", "jobs", ":jobId") { request async throws -> JobStatusResponse in
        let jobID = try request.requireUUID(named: "jobId")
        return try await request.application.jobStore.status(jobID: jobID)
    }

    protected.get("api", "jobs", ":jobId", "download.zip") { request async throws -> Response in
        let jobID = try request.requireUUID(named: "jobId")
        let payload = try await request.application.jobStore.archiveDownload(jobID: jobID)
        return makeDownloadResponse(for: payload)
    }

    protected.get("api", "jobs", ":jobId", "download", ":fileId") { request async throws -> Response in
        let jobID = try request.requireUUID(named: "jobId")
        let fileID = try request.requireUUID(named: "fileId")
        let payload = try await request.application.jobStore.fileDownload(jobID: jobID, fileID: fileID)
        return makeDownloadResponse(for: payload)
    }
}

private func decodeCreateJobForm(from request: Request) throws -> CreateJobForm {
    guard let boundary = request.headers.contentType?.parameters["boundary"] else {
        throw Abort(.unsupportedMediaType, reason: "需要使用 multipart/form-data 上传文件。")
    }

    guard let body = request.body.data else {
        throw Abort(.badRequest, reason: "请求体为空。")
    }

    let parser = MultipartParser(boundary: boundary)
    var parts: [MultipartPart] = []
    var currentHeaders: HTTPHeaders = .init()
    var currentBody = ByteBuffer()

    parser.onHeader = { field, value in
        currentHeaders.replaceOrAdd(name: field, value: value)
    }
    parser.onBody = { chunk in
        var chunk = chunk
        currentBody.writeBuffer(&chunk)
    }
    parser.onPartComplete = {
        parts.append(MultipartPart(headers: currentHeaders, body: currentBody))
        currentHeaders = .init()
        currentBody = ByteBuffer()
    }

    try parser.execute(body)

    guard let templatePart = parts.firstPart(named: "templateImage"), let templateFile = File(multipart: templatePart) else {
        throw Abort(.badRequest, reason: "缺少模板图。")
    }

    let documentParts = parts.filter { part in
        part.name == "documents" || part.name == "documents[]"
    }
    let documents = documentParts.compactMap(File.init(multipart:))

    guard documents.count == documentParts.count else {
        throw Abort(.badRequest, reason: "文档上传数据无效。")
    }

    return CreateJobForm(templateImage: templateFile, documents: documents)
}

private func validate(templateFile: File, documents: [File], config: AppConfig) throws {
    let templateExtension = URL(fileURLWithPath: templateFile.filename).pathExtension.lowercased()
    let allowedTemplateExtensions = Set(["png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff"])

    guard allowedTemplateExtensions.contains(templateExtension) else {
        throw Abort(.badRequest, reason: "模板图仅支持 png、jpg、gif、bmp、tif。")
    }

    guard !documents.isEmpty else {
        throw Abort(.badRequest, reason: "至少需要上传一份文档。")
    }

    let allowedDocumentExtensions = Set(["doc", "docx"])
    let totalBytes = templateFile.data.readableBytes + documents.reduce(0) { partialResult, file in
        partialResult + file.data.readableBytes
    }
    let limitBytes = config.maxUploadSizeMB * 1024 * 1024
    guard totalBytes <= limitBytes else {
        throw Abort(.badRequest, reason: "上传体积超过限制（\(config.maxUploadSizeMB)MB）。")
    }

    for file in documents {
        let ext = URL(fileURLWithPath: file.filename).pathExtension.lowercased()
        guard allowedDocumentExtensions.contains(ext) else {
            throw Abort(.badRequest, reason: "文档仅支持 .doc 和 .docx。")
        }
    }
}

private func makeDownloadResponse(for payload: DownloadPayload) -> Response {
    var headers = HTTPHeaders()
    headers.replaceOrAdd(name: .contentType, value: payload.contentType)
    headers.replaceOrAdd(
        name: .contentDisposition,
        value: #"attachment; filename="\#(payload.filename)""#
    )
    return Response(status: .ok, headers: headers, body: .init(data: payload.data))
}

private extension Request {
    func requireUUID(named parameter: String) throws -> UUID {
        guard let value = parameters.get(parameter), let uuid = UUID(uuidString: value) else {
            throw Abort(.badRequest, reason: "参数 \(parameter) 无效。")
        }
        return uuid
    }
}
