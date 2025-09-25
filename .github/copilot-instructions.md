# Copilot Instructions for Dashboard Project

## Special Triggers
- **"Review project"**: Perform full context loading and project state assessment:
  1. Read DEVELOPMENT_LOG.md, copilot-instructions.md, and Workbook_Metadata.txt
  2. Check Git status and recent changes
  3. Analyze current workspace structure
  4. Provide summary of project state and pending work
  5. Report any discrepancies or areas needing attention
- **"Always" keyword**: If the user includes the word "Always" in an instruction, treat it as a persistent rule and update this instruction file to reflect the concept automatically.
- **Always push after every change**: Commit and push to remote repository after every file modification, creation, or significant change to maintain continuous backup and synchronization.

## Project Overview
This project is a modular, high-performance Excel dashboard for industrial equipment search and display. It uses advanced VBA, dynamic configuration tables, and mode-driven search logic. The architecture is designed for maintainability, scalability, and easy transfer between environments.

## Project Documentation
- **DEVELOPMENT_LOG.md**: Complete history of project decisions and changes
- **ai_instructions**: Core agent behavior and workflow patterns
- **Workbook_Metadata.txt**: Current Excel structure and configuration

## Architecture & Major Components
- **Dashboard.cls**: Handles worksheet events, triggers refreshes, and manages pulse logic for dynamic updates.
- **mod_ModeDrivenSearch.bas**: Implements mode-driven search, filter formula evaluation, and output routines using the ModeConfigTable.
- **temp_mod_ConfigTableTools.bas**: Maintains and verifies ModeConfigTable entries, ensuring search modes are correctly defined and updated.
- **ThisWorkbook.cls**: Manages workbook open events, clears filters, and provides config access.
- **ConfigSheet & ModeConfigTable**: Central tables for configuration and search mode definitions.
- **Dev Notes worksheet**: Tracks project goals, features, and future enhancements.

## Key Patterns & Conventions

## Module Consolidation & Future Segregation
- **Development practice**: Keep code in a consolidated module (or as few as possible) during development for easier debugging and rapid iteration.
- **Modularization ready**: Write code with clear separation of concerns, using well-defined functions and comments, so it can be split into individual modules later with minimal changes.
- **Isolation pattern**: Encapsulate logic for search, config, and output routines so each can be moved to its own module when scaling up.

## Documentation & Code Organization
- **Extensive inline documentation**: Always include clear, well-organized notes and comments within your code. Document the purpose, dependencies, and integration points of each function, subroutine, and module.
- **Clarity and maintainability**: Structure comments to explain not just what the code does, but why it exists, how it interacts with other components, and any assumptions or requirements.
- **Dependency tracking**: Note any external references, required named ranges, tables, or config keys directly in the code for easy auditing and future refactoring.

## Developer Workflows
- **Metadata export**: Use the macro in `export.bas` to export workbook metadata (worksheets, named ranges, tables, config) for review or AI context.
- **ModeConfigTable management**: Update or add search modes using helper routines in `temp_mod_ConfigTableTools.bas`.
- **Debugging**: Use diagnostic logging patterns (see Dev Notes) and modular error handling. Toggle logging via config table entries.
- **Startup routines**: Workbook open events clear filters, reset inputs, and show the dashboard at top-left (see `ThisWorkbook.cls`).
- **Documentation Sync**: When updating instruction files:
  * Keep ai_instructions and copilot-instructions.md in sync
  * Update both files for shared concepts
  * Maintain distinct focus (ai_instructions for agent behavior, copilot-instructions.md for project structure)

## Integration Points
- **Excel structured tables**: All config and data flows use ListObjects and named ranges.
- **No external dependencies**: All logic is contained within VBA modules and Excel tables.

## Examples
- To add a new search mode, create a row in `ModeConfigTable` and use a helper like `Ensure_ModeConfigEntry_SootblowerLocation`.
- To update output columns, modify the `ConfigTable` and reference via config-driven routines.
- To refresh results, trigger the appropriate event in `Dashboard.cls` (e.g., pulse cell change).

## References
- `Dashboard.cls`, `mod_ModeDrivenSearch.bas`, `temp_mod_ConfigTableTools.bas`, `ThisWorkbook.cls`, `ConfigSheet`, `ModeConfigTable`, `Dev Notes worksheet`

## Special Triggers
- **"Always" keyword**: If the user includes the word "Always" in an instruction, treat it as a persistent rule and update this instruction file to reflect the concept automatically.

---

If any section is unclear or missing, please provide feedback so instructions can be improved for your workflow.
