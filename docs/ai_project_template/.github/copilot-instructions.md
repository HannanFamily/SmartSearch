# Copilot Instructions - Project Template

## Special Triggers
- **"Review project"**: Perform full context loading and project state assessment:
  1. Read DEVELOPMENT_LOG.md, copilot-instructions.md, and Workbook_Metadata.txt
  2. Check Git status and recent changes
  3. Analyze current workspace structure
  4. Provide summary of project state and pending work
  5. Report any discrepancies or areas needing attention
- **"Always" keyword**: If the user includes the word "Always" in an instruction, treat it as a persistent rule and update this instruction file to reflect the concept automatically.

## Project Documentation
- **DEVELOPMENT_LOG.md**: Complete history of project decisions and changes
- **ai_instructions**: Core agent behavior and workflow patterns
- **Workbook_Metadata.txt**: Current data structure and configuration

## Module Consolidation & Future Segregation
- **Development practice**: Keep code in a consolidated module (or as few as possible) during development for easier debugging and rapid iteration.
- **Modularization ready**: Write code with clear separation of concerns, using well-defined functions and comments, so it can be split into individual modules later with minimal changes.
- **Isolation pattern**: Encapsulate logic for search, config, and output routines so each can be moved to its own module when scaling up.

## Documentation & Code Organization
- **Extensive inline documentation**: Always include clear, well-organized notes and comments within your code. Document the purpose, dependencies, and integration points of each function, subroutine, and module.
- **Clarity and maintainability**: Structure comments to explain not just what the code does, but why it exists, how it interacts with other components, and any assumptions or requirements.
- **Dependency tracking**: Note any external references, required named ranges, tables, or config keys directly in the code for easy auditing and future refactoring.

## Developer Workflows
- **Metadata export**: Use macros or scripts to export workbook or data structure metadata for review or AI context.
- **Config table management**: Update or add search modes using helper routines.
- **Debugging**: Use diagnostic logging patterns and modular error handling. Toggle logging via config table entries.
- **Startup routines**: Ensure startup events clear filters, reset inputs, and show the dashboard at top-left.
- **Documentation Sync**: When updating instruction files:
  * Keep ai_instructions and copilot-instructions.md in sync
  * Update both files for shared concepts
  * Maintain distinct focus (ai_instructions for agent behavior, copilot-instructions.md for project structure)
- **Git Integration**: Always push changes after making edits to maintain synchronization.

## Integration Points
- **Structured tables**: All config and data flows use structured tables and named ranges.
- **No external dependencies**: All logic is contained within code modules and config tables.
- **Git Repository**: Project is maintained in a version-controlled repository.

## Examples
- To add a new search mode, create a row in the config table and use a helper routine.
- To update output columns, modify the config table and reference via config-driven routines.
- To refresh results, trigger the appropriate event in the dashboard module.

## References
- `ai_instructions`, `.github/DEVELOPMENT_LOG.md`, `Workbook_Metadata.txt`

---

If any section is unclear or missing, please provide feedback so instructions can be improved for your workflow.

# [END OF TEMPLATE]
