# Changelog

All notable changes to the CustomButton framework will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

---

## [1.5] - 2026-10-04

### Added

- Added the read-only `.ParentHwnd` property.
  - Stores the native window handle (`hWnd`) of the parent UserForm for advanced inter-window communication and tracking.

- Added support for new button configuration properties:
  - `.ControlTipText` – Enables native hover tooltips for custom button labels.
  - `.Locked` – Provides state locking compatibility.
  - `.HelpContextID` – Supports legacy Windows help system integration.
  - `.MousePointer` – Allows custom cursor selection via standard `fmMousePointer` enumerations.
  - `.MouseIcon` – Supports custom `.cur` file loading for tailored hover cursors.

### Changed

- Restructured the architectural design of the `CustomButton` class.
  - Introduced a unified button-top pattern to greatly simplify underlying code and message handling.
  - Enables complex features that require a single centralized control layer to function correctly.

### Fixed

- Resolved visual rendering glitches where rapid mouse movement could occasionally leave a button drawn in an incorrect hover or active state.

### Performance

- Implemented multiple internal code refinements to streamline event handling and reduce execution overhead.

### Documentation

- Updated and expanded the documentation to reflect all new button properties, architecture patterns, and features supported in version 1.5.

---

## [1.4] - 2026-05-17

### Added

- Initial Release of the `CustomButton` framework.
  - Core architecture established to replace standard MSForms `CommandButtons` with transparent `MSForms.Label` objects.
  - Features modern flat/ghost aesthetics, custom hover states, and click animations.