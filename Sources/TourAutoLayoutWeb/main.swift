import Foundation
import TourAutoLayoutCore
import Vapor

enum JobStoreKey: StorageKey {
    typealias Value = JobStore
}

enum AppConfigKey: StorageKey {
    typealias Value = AppConfig
}

extension Application {
    var appConfig: AppConfig {
        get {
            guard let value = storage[AppConfigKey.self] else {
                fatalError("AppConfig not configured")
            }
            return value
        }
        set {
            storage[AppConfigKey.self] = newValue
        }
    }

    var jobStore: JobStore {
        get {
            guard let value = storage[JobStoreKey.self] else {
                fatalError("JobStore not configured")
            }
            return value
        }
        set {
            storage[JobStoreKey.self] = newValue
        }
    }
}

var environment = try Environment.detect()
try LoggingSystem.bootstrap(from: &environment)

let app = Application(environment)
defer { app.shutdown() }

let config = try AppConfig.fromEnvironment()
try config.prepareStorage()

app.http.server.configuration.hostname = "0.0.0.0"
app.http.server.configuration.port = config.port
app.routes.defaultMaxBodySize = "200mb"
app.middleware.use(FileMiddleware(publicDirectory: app.directory.publicDirectory))

app.appConfig = config
app.jobStore = JobStore(config: config)

try routes(app)
try app.run()
