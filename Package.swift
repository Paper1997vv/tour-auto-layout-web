// swift-tools-version: 6.2

import PackageDescription

let package = Package(
    name: "TourAutoLayout",
    platforms: [
        .macOS(.v14),
    ],
    products: [
        .library(name: "TourAutoLayoutCore", targets: ["TourAutoLayoutCore"]),
        .executable(name: "TourAutoLayoutWeb", targets: ["TourAutoLayoutWeb"]),
    ],
    dependencies: [
        .package(url: "https://github.com/vapor/vapor.git", from: "4.110.1"),
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.19"),
    ],
    targets: [
        .target(
            name: "TourAutoLayoutCore",
            dependencies: [
                "ZIPFoundation",
            ],
            path: "Sources/TourAutoLayoutCore"
        ),
        .executableTarget(
            name: "TourAutoLayoutWeb",
            dependencies: [
                "TourAutoLayoutCore",
                .product(name: "Vapor", package: "vapor"),
            ],
            path: "Sources/TourAutoLayoutWeb"
        ),
        .testTarget(
            name: "TourAutoLayoutTests",
            dependencies: [
                "TourAutoLayoutCore",
                "ZIPFoundation",
            ],
            path: "Tests/TourAutoLayoutTests"
        ),
    ]
)
