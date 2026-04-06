import Foundation

public enum AppError: LocalizedError {
    case unsupportedSourceFormat(String)
    case unsupportedTemplateImageFormat(String)
    case outputPathUnavailable(String)
    case missingArchiveEntry(String)
    case malformedDocument(String)
    case commandFailed(String)
    case invalidInput(String)

    public var errorDescription: String? {
        switch self {
        case .unsupportedSourceFormat(let ext):
            return "不支持的文档格式：\(ext)"
        case .unsupportedTemplateImageFormat(let ext):
            return "不支持的模板图片格式：\(ext)"
        case .outputPathUnavailable(let name):
            return "无法为 \(name) 分配输出文件名。"
        case .missingArchiveEntry(let path):
            return "DOCX 缺少必要文件：\(path)"
        case .malformedDocument(let message):
            return "文档内容无法解析：\(message)"
        case .commandFailed(let message):
            return "系统转换命令失败：\(message)"
        case .invalidInput(let message):
            return message
        }
    }
}
