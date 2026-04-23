# Excel VBA Runtime Framework

> A modular framework for building structured and scalable Excel VBA solutions.

---

<img width="1536" height="1024" alt="umbrella" src="https://github.com/user-attachments/assets/6cbf4f4b-eefa-43e8-8c34-723e7d14d23f" />


## Overview

This repository provides a **cohesive runtime layer for Excel VBA**, designed to move beyond ad-hoc macros toward **engineered, maintainable applications**.

It introduces a structured approach to:

- execution control
- UI management
- event-driven interaction

The goal is to make Excel VBA more suitable for **complex, professional-grade use cases**, while keeping components modular, reusable, and independently adoptable.

---

## Architecture

The framework extends Excel with a runtime layer composed of three main domains:

- **Execution Engine**  
  Performance control, timing, and runtime environment management

- **UI Controller**  
  Centralized control of Excel interface elements and workbook presentation

- **Interaction Layer**  
  Event-driven components integrated with Excel workflows

---

## Components

### Execution Engine

#### [VBA Performance Manager](https://github.com/danielep71/VBA-Performance_Manager)

High-precision timing and execution-control component for Excel VBA on Windows.

It provides the foundation for:

- performance instrumentation
- runtime control
- repeatable benchmarking
- Excel environment optimization

Key capabilities include:

- multiple timing backends behind one interface
- session-bound timing model
- benchmark-overhead measurement
- shared Excel “time-waster” suppression
- performance improvement support even when no elapsed-time measurement is being taken

---

### UI Controller

#### [VBA-EXCEL_UI](https://github.com/danielep71/VBA-EXCEL_UI)

Centralized and structured control of Excel UI elements and application interface behavior.

It provides the foundation for:

- centralized Excel interface control
- application-like workbook presentation
- consistent visual environments
- reusable UI-management patterns in VBA projects

Key capabilities include:

- tri-state UI control (`show` / `hide` / `leave unchanged`)
- structured-result diagnostics
- explicit snapshot / reset lifecycle
- workbook demo support
- regression validation support
- WinAPI-based title-bar visibility management

---

### Interaction Layer

#### Date-Time Picker *(coming soon)*

Event-driven date and time selection integrated directly into Excel workflows.

Planned focus areas include:

- contextual UI activation
- event-driven interaction with worksheet selection
- guided date and time entry
- reusable workbook-level interaction patterns

---

## Design principles

### Structured, not ad-hoc

Promotes modular design, separation of concerns, and reusable components rather than scattered procedural macros.

---

### Deterministic behavior

Emphasizes clear error policies, predictable execution paths, and explicit control semantics.

---

### Performance-first

Execution control is treated as a primary concern, not an afterthought.

---

### Event-driven architecture

Components react to Excel context rather than relying only on manual triggers.

---

### Reusable component model

Each component is designed to be usable on its own, while also fitting into a broader framework architecture.

---

## Example use cases

- high-performance processing of large datasets
- Excel-based applications with controlled UI
- interactive data entry systems
- consistent execution environments for complex workflows
- presentation-oriented or kiosk-like workbook shells

---

## Ecosystem direction

This framework is designed to evolve into a broader ecosystem, including:

- additional UI components
- extended runtime utilities
- richer interaction layers
- integration with external data sources
- reusable patterns for advanced Excel application design

---

## Related work

### KPR Financial Pricing Framework *(planned)*

A complementary VBA framework focused on financial analytics, pricing logic, date conventions, curve construction, and instrument-level calculations.

This is intended to sit alongside the runtime framework rather than inside it.

---

## Author

Daniele Penza
