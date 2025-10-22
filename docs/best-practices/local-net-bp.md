### .NET Local Device Projects — Best Practices

---

### Overview
Clean, secure, testable, and maintainable code for local devices focuses on small deployable footprints, offline-first data, resilient local storage, low-latency UX, secure on-device secrets, and reproducible developer workflows that match the device environment.

---

### Targets and project setup
- **Choose the right runtime**
  - .NET MAUI for cross-platform mobile/desktop; WinUI/WPF for Windows-native; .NET for IoT on ARM; Blazor Hybrid for web-like local UIs.
- **Project layout**
  - Organize by feature/vertical slice with a shared domain library for business rules; separate projects for UI, device integrations, and platform shims.
- **Multi-targeting**
  - Multi-target libraries to share code across Windows, Android, iOS, and Linux; keep platform-specific code behind clear abstractions.
- **Packaging and publishing**
  - Prefer self-contained single-file publishes for constrained devices; choose runtime identifiers (RIDs) matching target architectures.

---

### Architecture, storage, and sync
- **Offline-first model**
  - Treat the device as the authoritative local source; design deterministic conflict-resolution and local-first UX.
- **Local storage choices**
  - Use SQLite (Microsoft.Data.Sqlite or EF Core) for structured data; file system or isolated storage for blobs; consider Realm or LiteDB for simpler models or built-in sync.
- **Sync and reconciliation**
  - Implement incremental sync, change-tracking (timestamps/version tokens), idempotent operations, optimistic concurrency, and merge rules that preserve user intent.
- **Transport abstraction**
  - Expose a transport interface for syncing so HTTP, gRPC, or custom relays can be swapped without touching domain code.
- **Power and connectivity awareness**
  - Detect network and battery status; adapt sync frequency, defer heavy work when low power, and use background tasks sensibly.

---

### Security and secrets on-device
- **Platform secure storage**
  - Use Windows Credential Locker, iOS Keychain, Android Keystore; wrap access behind an interface for testing and desktop variants.
- **Least privilege**
  - Request minimal permissions and surface intent/explanation in the UX before prompting.
- **Encrypt sensitive local data**
  - Protect databases and files with device-backed keys; rely on platform crypto primitives; do not roll your own cryptography.
- **Authentication and token handling**
  - Use short-lived tokens, refresh securely, and purge cached credentials on sign-out or suspected compromise.
- **Tamper and integrity checks**
  - Validate integrity of critical assets and detect debug/tamper states where appropriate for your threat model.

---

### Developer workflows, testing, and CI for device projects
- **Emulators and hardware parity**
  - Automate tests on emulators and a representative set of physical devices; include at least one physical-device run in CI for platform-specific issues.
- **Unit and integration tests**
  - Keep domain logic testable using DI and abstractions; run storage integration tests against disposable SQLite files or in-memory providers.
- **UI and E2E automation**
  - Use Appium, WinAppDriver, or MAUI UITest for smoke flows; create reproducible test data sets for manual QA.
- **CI for device builds**
  - Build per-target RID, sign artifacts when required, produce self-contained bundles, and run linters, analyzers, and tests in pipeline.
- **Developer ergonomics**
  - Provide developer-mode toggles, verbose telemetry that’s build-time disableable, and scripts to reproduce production-like launches locally.

---

### Observability, performance, and reliability
- **On-device telemetry**
  - Collect lightweight, privacy-aware metrics and logs; batch uploads and respect user opt-in/consent requirements.
- **Low-overhead diagnostics**
  - Use EventCounters, DiagnosticSource, or compact traces that can be enabled for troubleshooting without major overhead.
- **Resource constraints**
  - Measure memory, CPU, and battery on representative hardware; avoid large allocations, prefer Span<T> and pooling in hot paths.
- **Resilience patterns**
  - Use retries with exponential backoff, circuit breakers, and graceful degradation; design for intermittent connectivity.
- **Update and rollout strategy**
  - Support incremental updates and safe rollback; for intermittent connectivity, allow staged rollouts and remote feature flags.

---

### Coding standards and patterns
- **Small, focused components**
  - Keep methods and classes single-responsibility to make resource usage predictable and testing easier.
- **Dependency injection and abstraction**
  - Use DI for platform services (storage, network, sensors) so you can stub them for tests.
- **Async everywhere for I/O**
  - Use async/await for network and storage APIs to keep UI responsive and avoid blocking threads.
- **Fail-fast and observable errors**
  - Surface recoverable vs unrecoverable errors clearly; log actionable context to aid debugging on devices.

---

### Quick actionable checklist
- Target the appropriate .NET runtime and multi-target when needed.  
- Isolate platform code behind interfaces; keep domain logic pure.  
- Use SQLite or a small local DB and adopt an offline-first sync model.  
- Store secrets in platform secure storage and encrypt sensitive files.  
- Produce self-contained bundles per device RID and sign where required.  
- Run unit tests, DB integration tests, and UI automation on emulators plus at least one physical device.  
- Measure and optimize memory, CPU, and battery on representative hardware.  
- Implement incremental sync, exponential backoff, and graceful offline degradation.  
- Ship privacy-respecting telemetry that can be toggled remotely.  
- Provide developer scripts and sample datasets to reproduce device scenarios locally.

---

### Notes for teams
- Maintain ADRs for device-specific architectural choices and keep them near