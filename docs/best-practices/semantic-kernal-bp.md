# Semantic Kernel Agents Connecting to Local LLMs — Best Practices

---

## Overview
Design Semantic Kernel agents for local LLMs with privacy, reliability, and low-latency in mind. Prioritize clear capability boundaries, robust prompt engineering, safe execution of tools and plugins, reproducible developer workflows, and lightweight observability suitable for offline or partially connected environments.

---

## Goals and constraints
- **Primary goals**
  - Keep sensitive data on-device; avoid sending private artifacts to external services.
  - Enable interactive, low-latency agent behavior that runs reliably on constrained hardware.
  - Make agent behavior auditable, testable, and reproducible across devices.
- **Common constraints**
  - Limited CPU, memory, and storage; variable GPU availability.
  - Intermittent or no network connectivity.
  - Device-specific LLM runtimes and model formats.

---

## Agent architecture and separation of concerns
- **Core separation**
  - Keep the Semantic Kernel core (planning, memory, skill orchestration) separate from model adapters, transport, and platform shims.
- **Skill surface design**
  - Implement small, well-tested skills with explicit inputs, outputs, and side-effect descriptions.
  - Limit skill privileges; require explicit allow-lists for skills that perform sensitive operations.
- **Transport and runtime adapters**
  - Wrap local LLM runtimes behind an adapter interface to allow switching between runtimes (e.g., llama.cpp, ONNX runtime, vendor SDK) without changing kernel logic.
- **Tooling and tool invocation**
  - Model-based tool selection must be deterministic and testable; use structured tool invocation formats (JSON or function-calling) rather than free-form text.

---

## Prompting, context, and retrieval
- **Prompt templates**
  - Store versioned prompt templates and document their intended use; avoid embedding secrets directly in templates.
- **Context window management**
  - Use retrieval-augmented generation (RAG) to fit large context into smaller model windows; prioritize recent and high-signal context.
- **Chunking and embeddings**
  - Preprocess long documents into consistent chunks with metadata, generate embeddings locally, and use a local vector store optimized for the device (lightweight SQLite+FAISS/Annoy/Weaviate-lite).
- **Prompt safety and instruction guarding**
  - Sanitize and validate instructions that could cause risky behavior; use guardrails in prompts and post-response validators.

---

## Security, privacy, and secrets
- **Keep data local**
  - Default to local-only operation; require explicit opt-in for telemetry or external services.
- **Secrets and keys**
  - Store any keys or credentials in platform secure storage (Keychain/Keystore/Credential Locker); never bake secrets into models or prompt templates.
- **Execution safety**
  - Run any code or shell-invoking skills in isolated, limited-permission sandboxes; enforce timeouts and resource caps.
- **Model integrity and provenance**
  - Verify model artifacts using signatures/checksums at load time; log model version and source for audits.

---

## Testing, reproducibility, and CI
- **Deterministic testing**
  - Pin seeds, model versions, and prompt templates for unit tests of agent behavior.
- **Behavioral unit tests**
  - Test skills and orchestration with canned model responses using mocked local LLM adapters.
- **Integration tests**
  - Run integration tests against representative local runtimes in CI (headless or containerized) and include a small set of physical-device smoke tests where necessary.
- **Regression tests for prompts**
  - Add golden-response or invariants checks for critical prompts to detect prompt drift after template or model updates.

---

## Observability and debugging
- **Lightweight telemetry**
  - Capture structured logs of agent decisions, invoked skills, prompt templates used, model versions, and confidence scores; batch and encrypt telemetry for optional upload when allowed.
- **Reproducible traces**
  - Record the minimal reproducer for unusual behavior: model version, seed, prompt, retrieved