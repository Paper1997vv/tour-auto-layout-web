import SwiftUI

@main
struct TourAutoLayoutApp: App {
    @StateObject private var viewModel = AppViewModel()

    var body: some Scene {
        WindowGroup("旅游行程自动排版") {
            ContentView(viewModel: viewModel)
                .frame(minWidth: 900, minHeight: 640)
        }
        .windowResizability(.contentSize)
    }
}
