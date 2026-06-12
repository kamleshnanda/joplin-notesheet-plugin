# Requirements Document

## Introduction

This document specifies the requirements for a Joplin plugin that integrates Univer spreadsheet functionality. The plugin will replace the deprecated LuckySheets-based JSheets plugin with a modern, open-source spreadsheet solution that provides Excel-like features within Joplin notes. The plugin must support formula calculations, data persistence across Dropbox syncs, and seamless integration with Joplin's note editing experience.

## Glossary

- **Joplin**: An open-source note-taking application that supports Markdown and plugin extensions
- **Univer**: An open-source office document embedding solution providing spreadsheet functionality
- **Plugin**: A Joplin extension that adds new functionality to the application
- **Spreadsheet_Component**: The embedded Univer spreadsheet interface within a Joplin note
- **Note_Sync**: Joplin's mechanism for synchronizing notes across devices via Dropbox
- **Formula_Engine**: The calculation system that processes spreadsheet formulas
- **Plugin_API**: Joplin's interface for plugin development and integration

## Requirements

### Requirement 1: Plugin Project Structure

**User Story:** As a plugin developer, I want a properly structured Joplin plugin project, so that I can develop and build the plugin according to Joplin's standards.

#### Acceptance Criteria

1. THE Plugin SHALL conform to Joplin's plugin manifest specification
2. THE Plugin SHALL include all required configuration files for building and packaging
3. THE Plugin SHALL organize source code following Joplin plugin best practices
4. THE Plugin SHALL declare all necessary dependencies in the package configuration
5. THE Plugin SHALL provide build scripts for development and production packaging

### Requirement 2: Univer SDK Integration

**User Story:** As a plugin developer, I want to integrate the Univer SDK, so that spreadsheet functionality is available within the plugin.

#### Acceptance Criteria

1. THE Plugin SHALL initialize the Univer SDK using the free open-source version
2. THE Plugin SHALL configure Univer with spreadsheet capabilities enabled
3. THE Plugin SHALL load required Univer modules and dependencies
4. WHEN the plugin initializes, THE Plugin SHALL verify Univer SDK availability
5. IF Univer initialization fails, THEN THE Plugin SHALL log an error and degrade gracefully

### Requirement 3: Spreadsheet Embedding

**User Story:** As a Joplin user, I want to embed spreadsheets within my notes, so that I can include structured data alongside my text content.

#### Acceptance Criteria

1. WHEN a user invokes the spreadsheet creation command, THE Plugin SHALL insert a spreadsheet placeholder in the note
2. THE Spreadsheet_Component SHALL render within the note editor interface
3. THE Spreadsheet_Component SHALL display with appropriate dimensions and styling
4. WHEN a note contains multiple spreadsheets, THE Plugin SHALL render each independently
5. THE Plugin SHALL distinguish spreadsheet content from regular Markdown content
6. THE Plugin SHALL support inline spreadsheet initialization within the Markdown note body
7. THE Spreadsheet_Component SHALL render at the exact position where the fence block appears in the Markdown content

### Requirement 4: Spreadsheet Creation Interface

**User Story:** As a Joplin user, I want an intuitive interface for creating new spreadsheets, so that I can quickly add spreadsheets to my notes.

#### Acceptance Criteria

1. THE Plugin SHALL provide a toolbar button for creating new spreadsheets
2. THE Plugin SHALL provide a keyboard shortcut for creating new spreadsheets
3. THE Plugin SHALL provide a menu command for creating new spreadsheets
4. WHEN a user creates a spreadsheet, THE Plugin SHALL insert it at the current cursor position
5. WHEN a spreadsheet is created, THE Plugin SHALL initialize it with default dimensions

### Requirement 5: Spreadsheet Editing Interface

**User Story:** As a Joplin user, I want to edit spreadsheet content directly within my notes, so that I can modify data without leaving the note editor.

#### Acceptance Criteria

1. WHEN a user clicks on a spreadsheet, THE Spreadsheet_Component SHALL enter edit mode
2. THE Spreadsheet_Component SHALL support cell selection and navigation
3. THE Spreadsheet_Component SHALL support text input in selected cells
4. THE Spreadsheet_Component SHALL support copy, cut, and paste operations
5. WHEN a user clicks outside the spreadsheet, THE Spreadsheet_Component SHALL exit edit mode

### Requirement 6: Formula Support

**User Story:** As a Joplin user, I want to use formulas in spreadsheet cells, so that I can perform calculations on my data.

#### Acceptance Criteria

1. WHEN a user enters a formula starting with "=", THE Formula_Engine SHALL evaluate the formula
2. THE Formula_Engine SHALL support basic arithmetic operations (+, -, \*, /)
3. THE Formula_Engine SHALL support common functions (SUM, AVERAGE, COUNT, MIN, MAX)
4. WHEN a formula references other cells, THE Formula_Engine SHALL update when referenced cells change
5. IF a formula contains an error, THEN THE Spreadsheet_Component SHALL display an error indicator

### Requirement 7: Data Persistence

**User Story:** As a Joplin user, I want my spreadsheet data to persist across sessions, so that my work is not lost when I close and reopen notes.

#### Acceptance Criteria

1. WHEN a user modifies spreadsheet content, THE Plugin SHALL serialize the data to a storage format
2. THE Plugin SHALL store spreadsheet data within the note's content structure
3. WHEN a note is opened, THE Plugin SHALL deserialize and restore spreadsheet data
4. THE Plugin SHALL preserve all cell values, formulas, and formatting
5. THE Plugin SHALL use a format compatible with Joplin's Markdown storage

### Requirement 8: Sync Compatibility

**User Story:** As a Joplin user, I want my spreadsheets to sync across devices via cloud services, so that I can access my data on desktop and mobile.

#### Acceptance Criteria

1. THE Plugin SHALL store spreadsheet data in a format that survives Note_Sync
2. WHEN a note syncs to another device, THE Plugin SHALL correctly restore spreadsheet content
3. THE Plugin SHALL handle sync conflicts gracefully without data corruption
4. THE Plugin SHALL maintain data integrity across different Joplin versions
5. THE Plugin SHALL use text-based storage to ensure compatibility with Dropbox, Joplin Cloud, and other cloud sync services
6. THE Plugin SHALL store spreadsheet data entirely within the note's Markdown content to ensure cloud backup compatibility

### Requirement 9: Plugin Build and Packaging

**User Story:** As a plugin developer, I want to build and package the plugin, so that it can be installed in a local Joplin instance.

#### Acceptance Criteria

1. THE Plugin SHALL provide a build command that compiles all source code
2. THE Plugin SHALL bundle all dependencies into the plugin package
3. THE Plugin SHALL generate a .jpl file for Joplin installation
4. THE Plugin SHALL include version information in the package manifest
5. WHEN the build completes, THE Plugin SHALL output the package to a known location

### Requirement 10: Local Installation and Testing

**User Story:** As a plugin developer, I want to install and test the plugin locally, so that I can verify functionality before distribution.

#### Acceptance Criteria

1. THE Plugin SHALL be installable via Joplin's plugin installation interface
2. WHEN installed, THE Plugin SHALL appear in Joplin's plugin list
3. THE Plugin SHALL activate without errors in a local Joplin instance
4. THE Plugin SHALL provide visible UI elements (toolbar buttons, menu items)
5. WHEN tested, THE Plugin SHALL successfully create and edit spreadsheets in notes

### Requirement 11: Error Handling

**User Story:** As a Joplin user, I want the plugin to handle errors gracefully, so that my note-taking experience is not disrupted by plugin failures.

#### Acceptance Criteria

1. IF the Univer SDK fails to load, THEN THE Plugin SHALL display an error message to the user
2. IF spreadsheet data is corrupted, THEN THE Plugin SHALL attempt recovery or display a warning
3. WHEN an error occurs, THE Plugin SHALL log diagnostic information for debugging
4. THE Plugin SHALL prevent errors from crashing the Joplin application
5. IF a formula evaluation fails, THEN THE Spreadsheet_Component SHALL display an error in the cell

### Requirement 12: User Experience

**User Story:** As a Joplin user, I want a seamless spreadsheet editing experience, so that working with spreadsheets feels natural within my notes.

#### Acceptance Criteria

1. THE Spreadsheet_Component SHALL render with minimal visual delay
2. THE Spreadsheet_Component SHALL respond to user input without noticeable lag
3. THE Plugin SHALL maintain consistent styling with Joplin's interface
4. THE Spreadsheet_Component SHALL support keyboard navigation for accessibility
5. WHEN switching between notes, THE Plugin SHALL preserve spreadsheet state correctly

### Requirement 13: Software License

**User Story:** As a Joplin user, I want to pay attention to the opensource requirements so that I can package the plugin and publish on github and eventually to Joplin plugin repo for free usage

#### Acceptance Criteria

1. THE Plugin SHALL contain code that can be used in opensource without license.
2. THE Plugin SHALL provide installation setup that is minimal and works out of the box after installation.
3. THE Plugin code SHALL following github structure with appropriate project level documentation for other users to be able to contribute.
4. THE Plugin should be buildable and produce packages that can be published on Joplin plugin repository.
5. The Plugin once published on Joplin plugin repository should be discoverable with the following keywords "Excel", "Spreadsheet", "Univer", "Kamlesh"
