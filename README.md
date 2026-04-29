# T3Lab Revit API

[![Revit Version](https://img.shields.io/badge/Revit-2020%2B-blue.svg)](https://www.autodesk.com/products/revit/overview)
[![pyRevit](https://img.shields.io/badge/pyRevit-4.8%2B-orange.svg)](https://github.com/eirannejad/pyRevit)
[![Architecture](https://img.shields.io/badge/Architecture-T3Lab%20MAS%203.0-green.svg)](#architecture)

T3Lab Revit API is an advanced BIM Automation and Intelligence framework for Autodesk Revit. Built upon the **T3Lab Master Architecture System (MAS)**, it bridges the gap between traditional BIM workflows and modern Artificial Intelligence.

---

## Architecture: T3Lab MAS 3.0

The framework is organized into three distinct layers, creating a self-sustaining ecosystem for architectural intelligence:

### 1. Intelligence Layer (Agent-Core)
*   **T3Lab Assistant**: A natural language bridge allowing users to query and command Revit using Vietnamese or English.
*   **Local Intelligence**: Support for offline LLMs (via Ollama) ensuring data privacy and high-speed local inference.

### 2. Execution Layer (Tool-Stack)
*   **Production Tools**: Specialized panels for Annotation, Project Management, and Data Export.
*   **Discipline Modules**: Advanced logic for Architecture, Structure, and Coordination.
*   **Automation Wizards**: Multi-step wizards (like BatchOut) that handle complex task sequences with minimal user input.

### 3. Data Fabric Layer (Cloud-Connect)
*   **Cloud API**: Vercel-hosted backend for distributed family loading and metadata management.
*   **Hybrid Storage**: Seamless switching between local library files and cloud-hosted BIM content.

---

## Key Modules

### AI Connection & Control
*   **T3Lab Assistant**: Context-aware AI that understands Revit terminology and executes complex commands.

### BatchOut: The Export Master
*   A unified export engine for PDF, DWG, NWD, and IFC.
*   Features custom naming patterns, revision-aware filtering, and automatic sheet set generation.

### Project & Annotation Intelligence
*   **Workset Management**: Automated isolation and coordination views.
*   **SmartAlign**: Geometry-aware alignment and distribution logic.
*   **Family Synthesis**: Generate parametric families from structured JSON data.

---

## The Evolution Loop

T3Lab Revit API is designed for **Self-Evolution**. By utilizing specialized agents (QA-Agent, Tool-Builder, UI-Agent), the framework allows for:
1.  **Automated Error Detection**: Identifying BIM inconsistencies via the `checks` module.
2.  **Rapid Tool Iteration**: Building new pushbutton scripts through the `tool-builder` agent.
3.  **Knowledge Accumulation**: Storing project-specific wisdom in a local knowledge graph.

---

## Project Structure

```bash
t3lab-revit-api/
├── T3Lab.extension/          # Main pyRevit Extension
│   ├── T3Lab.tab/            # Ribbon Panels (AI, Annotation, Project, Export)
│   ├── lib/                  # Framework Core (GUI, NLU, RAG, Utils)
│   ├── checks/               # Quality Assurance Scripts
│   └── commands/             # Standalone Automation Commands
├── api/                      # Vercel Serverless Functions
├── scripts/                  # Maintenance & Environment Setup
├── requirements.txt          # Framework Dependencies
└── vercel.json               # Cloud Infrastructure Config
```

---

## Setup & Installation

1.  Clone this repository to: `%APPDATA%\pyRevit\Extensions\T3Lab.extension`
2.  Ensure **pyRevit 4.8+** is installed and linked.
3.  Reload pyRevit to initialize the **T3Lab** tab.
4.  Configure AI settings in the **AI Connection > Settings** panel.

---

## Author
**Tran Tien Thanh**
Architect & BIM Developer
- [trantienthanh909@gmail.com](mailto:trantienthanh909@gmail.com)
- [T3Lab.Space](https://t3lab.space)

---
*Empowering BIM with Intelligence.*
