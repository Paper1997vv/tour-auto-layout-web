import SwiftUI
import UniformTypeIdentifiers

struct DropZoneView: View {
    let title: String
    let subtitle: String
    let onURLsDropped: ([URL]) -> Void

    @State private var isTargeted = false

    var body: some View {
        VStack(alignment: .leading, spacing: 16) {
            Text(title)
                .font(.headline)
            Text(subtitle)
                .foregroundStyle(.secondary)

            Spacer()

            VStack(spacing: 10) {
                Image(systemName: "doc.on.doc")
                    .font(.system(size: 42, weight: .medium))
                Text("把 Word 文档拖到这里")
                    .font(.title3.weight(.semibold))
                Text("仅接收 .doc / .docx")
                    .foregroundStyle(.secondary)
            }
            .frame(maxWidth: .infinity, maxHeight: .infinity)
            .padding(24)
            .background(
                RoundedRectangle(cornerRadius: 18)
                    .strokeBorder(isTargeted ? Color.accentColor : Color.secondary.opacity(0.25), style: StrokeStyle(lineWidth: 2, dash: [8, 6]))
                    .background(
                        RoundedRectangle(cornerRadius: 18)
                            .fill(isTargeted ? Color.accentColor.opacity(0.08) : Color(nsColor: .controlBackgroundColor))
                    )
            )
        }
        .padding(18)
        .background(RoundedRectangle(cornerRadius: 18).fill(Color(nsColor: .controlBackgroundColor)))
        .onDrop(of: [UTType.fileURL.identifier], isTargeted: $isTargeted, perform: handleDrop)
    }

    private func handleDrop(_ providers: [NSItemProvider]) -> Bool {
        guard !providers.isEmpty else { return false }

        let group = DispatchGroup()
        let collector = URLCollector()

        for provider in providers where provider.hasItemConformingToTypeIdentifier(UTType.fileURL.identifier) {
            group.enter()
            provider.loadItem(forTypeIdentifier: UTType.fileURL.identifier, options: nil) { item, _ in
                defer { group.leave() }

                if let data = item as? Data, let url = URL(dataRepresentation: data, relativeTo: nil) {
                    collector.append(url)
                } else if let url = item as? URL {
                    collector.append(url)
                }
            }
        }

        group.notify(queue: .main) {
            onURLsDropped(collector.urls)
        }

        return true
    }
}

private final class URLCollector: @unchecked Sendable {
    private let lock = NSLock()
    private var storage = [URL]()

    func append(_ url: URL) {
        lock.lock()
        storage.append(url)
        lock.unlock()
    }

    var urls: [URL] {
        lock.lock()
        defer { lock.unlock() }
        return storage
    }
}
