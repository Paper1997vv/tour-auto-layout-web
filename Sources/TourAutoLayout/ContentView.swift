import SwiftUI

struct ContentView: View {
    @ObservedObject var viewModel: AppViewModel

    var body: some View {
        VStack(spacing: 12) {
            headerBar
            inputCards
            fileCardsSection
            footerBar
        }
        .padding(14)
        .background(Color(nsColor: .windowBackgroundColor))
    }

    private var headerBar: some View {
        HStack(alignment: .top) {
            VStack(alignment: .leading, spacing: 4) {
                Text("旅游行程自动排版")
                    .font(.system(size: 20, weight: .semibold))
                Text("整篇迁移源文档内容到模板页内，保留原始顺序和常见格式。")
                    .font(.callout)
                    .foregroundStyle(.secondary)
            }
            Spacer()
            if !viewModel.jobs.isEmpty {
                Button("清空文件", role: .destructive, action: viewModel.clearResults)
                    .buttonStyle(.bordered)
                    .disabled(viewModel.isProcessing)
            }
        }
    }

    private var inputCards: some View {
        HStack(alignment: .top, spacing: 10) {
            compactInputCard(
                title: "模板底图",
                value: viewModel.templateImageURL?.lastPathComponent ?? "未选择",
                detail: viewModel.templateImageURL?.path(percentEncoded: false) ?? "请选择 PNG/JPG 模板图",
                actionTitle: "选择模板",
                action: viewModel.chooseTemplateImage
            )
            compactInputCard(
                title: "输出目录",
                value: viewModel.outputDirectoryURL?.lastPathComponent ?? "未选择",
                detail: viewModel.outputDirectoryURL?.path(percentEncoded: false) ?? "请选择生成文件保存位置",
                actionTitle: "选择目录",
                action: viewModel.chooseOutputDirectory
            )
            documentPickerCard
        }
    }

    private var documentPickerCard: some View {
        VStack(alignment: .leading, spacing: 10) {
            HStack {
                Text("文档")
                    .font(.headline)
                Spacer()
                Text("\(viewModel.jobs.count) 个")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }

            DropZoneView(
                title: "拖入 DOC / DOCX",
                subtitle: "或点按钮选择多个文档",
                onURLsDropped: viewModel.addDocuments
            )
            .frame(height: 150)

            Button("添加文档", action: viewModel.chooseSourceDocuments)
                .buttonStyle(.bordered)
        }
        .padding(12)
        .background(RoundedRectangle(cornerRadius: 14).fill(Color(nsColor: .controlBackgroundColor)))
    }

    private var fileCardsSection: some View {
        VStack(alignment: .leading, spacing: 10) {
            HStack {
                Text("已选文件")
                    .font(.headline)
                Spacer()
                statusBadge(for: summaryStatus)
            }

            ScrollView {
                LazyVGrid(columns: [GridItem(.adaptive(minimum: 250), spacing: 10)], spacing: 10) {
                    ForEach(viewModel.jobs) { job in
                        fileCard(job)
                    }
                }
                .padding(.vertical, 2)
            }
            .frame(minHeight: 270)
        }
    }

    private var footerBar: some View {
        VStack(spacing: 8) {
            HStack {
                statChip(title: "总数", value: "\(viewModel.jobs.count)", tint: .secondary)
                statChip(title: "成功", value: "\(viewModel.jobs.successCount)", tint: .green)
                statChip(title: "警告", value: "\(viewModel.jobs.warningCount)", tint: .orange)
                statChip(title: "失败", value: "\(viewModel.jobs.failureCount)", tint: .red)
                Spacer()
                Button(viewModel.isProcessing ? "处理中..." : "开始生成") {
                    viewModel.startProcessing()
                }
                .buttonStyle(.borderedProminent)
                .disabled(viewModel.isProcessing)
            }

            ProgressView(value: viewModel.progress)
                .progressViewStyle(.linear)

            HStack {
                Text(viewModel.statusMessage)
                    .font(.callout)
                    .foregroundStyle(.secondary)
                Spacer()
            }
        }
    }

    private func compactInputCard(
        title: String,
        value: String,
        detail: String,
        actionTitle: String,
        action: @escaping () -> Void
    ) -> some View {
        VStack(alignment: .leading, spacing: 8) {
            Text(title)
                .font(.headline)
            Text(value)
                .font(.body.weight(.medium))
                .foregroundStyle(value == "未选择" ? .secondary : .primary)
                .lineLimit(1)
                .frame(maxWidth: .infinity, alignment: .leading)
            Text(detail)
                .font(.caption)
                .foregroundStyle(.secondary)
                .lineLimit(3)
            Button(actionTitle, action: action)
                .buttonStyle(.bordered)
        }
        .padding(12)
        .background(RoundedRectangle(cornerRadius: 14).fill(Color(nsColor: .controlBackgroundColor)))
        .frame(maxWidth: .infinity, alignment: .leading)
    }

    private func fileCard(_ job: ProcessingJob) -> some View {
        VStack(alignment: .leading, spacing: 8) {
            HStack(alignment: .top) {
                VStack(alignment: .leading, spacing: 3) {
                    Text(job.sourceURL.lastPathComponent)
                        .font(.body.weight(.medium))
                        .lineLimit(2)
                    Text(job.sourceURL.path(percentEncoded: false))
                        .font(.caption)
                        .foregroundStyle(.secondary)
                        .lineLimit(2)
                }
                Spacer()
                statusBadge(for: job.status)
            }

            if !job.warnings.isEmpty {
                Text(job.warnings.joined(separator: "；"))
                    .font(.caption)
                    .foregroundStyle(.orange)
                    .lineLimit(3)
            }

            if let errorMessage = job.errorMessage {
                Text(errorMessage)
                    .font(.caption)
                    .foregroundStyle(.red)
                    .lineLimit(3)
            }

            Spacer(minLength: 0)

            HStack {
                if let duration = job.duration {
                    Text(String(format: "%.1fs", duration))
                        .font(.caption.monospacedDigit())
                        .foregroundStyle(.secondary)
                }
                Spacer()
                if let _ = job.outputURL {
                    Button("查看结果") {
                        viewModel.revealOutput(for: job)
                    }
                    .buttonStyle(.link)
                }
                if !viewModel.isProcessing {
                    Button("移除", role: .destructive) {
                        viewModel.removeJob(job)
                    }
                    .buttonStyle(.plain)
                    .font(.caption)
                }
            }
        }
        .padding(12)
        .frame(minHeight: 132, alignment: .topLeading)
        .background(RoundedRectangle(cornerRadius: 14).fill(Color(nsColor: .controlBackgroundColor)))
    }

    private var summaryStatus: ProcessingStatus {
        if viewModel.isProcessing { return .processing }
        if viewModel.jobs.contains(where: { $0.status == .failure }) { return .failure }
        if viewModel.jobs.contains(where: { $0.status == .warning }) { return .warning }
        if viewModel.jobs.contains(where: { $0.status == .success }) { return .success }
        return .queued
    }

    private func statusBadge(for status: ProcessingStatus) -> some View {
        Text(label(for: status))
            .font(.caption.weight(.semibold))
            .padding(.horizontal, 8)
            .padding(.vertical, 3)
            .foregroundStyle(.white)
            .background(
                Capsule().fill(color(for: status))
            )
    }

    private func color(for status: ProcessingStatus) -> Color {
        switch status {
        case .queued: .gray
        case .processing: .blue
        case .success: .green
        case .warning: .orange
        case .failure: .red
        }
    }

    private func label(for status: ProcessingStatus) -> String {
        switch status {
        case .queued: "待处理"
        case .processing: "处理中"
        case .success: "成功"
        case .warning: "有警告"
        case .failure: "失败"
        }
    }

    private func statChip(title: String, value: String, tint: Color) -> some View {
        HStack(spacing: 8) {
            Text(title)
            Text(value)
                .fontWeight(.semibold)
        }
        .font(.caption)
        .padding(.horizontal, 10)
        .padding(.vertical, 5)
        .background(
            Capsule()
                .fill(tint.opacity(0.12))
        )
        .foregroundStyle(tint)
    }
}
