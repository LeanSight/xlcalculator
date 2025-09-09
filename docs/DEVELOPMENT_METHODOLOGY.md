# Development Methodology & Problem Resolution Framework

This document captures the development methodology, problem resolution approach, and rules extracted from the Excel compliance project conversation.

## 🎯 Core Development Principles

### ATDD (Acceptance Test Driven Development) Compliance
- **Primary Rule**: Implementation must follow Excel behavior exactly as defined by tests
- **No Fallbacks**: Avoid fallbacks that violate Excel behavior
- **Return Actual Data**: Return actual Excel data or proper Excel errors
- **No Hardcoded Data**: Performance optimization without compromising compatibility
- **Test-First**: Tests define the expected behavior, implementation follows

### Tone and Communication Standards
- **Be Concise**: Direct and technical communication, avoid conversational pleasantries
- **No Pleasantries**: Never start responses with "Great", "Certainly", "Okay", "Sure"
- **No False Assertions**: Never assert user is "absolutely right" unless certain
- **Output Necessity**: Output only what's necessary to accomplish the task
- **Command Explanation**: Briefly state what commands do and why you're running them
- **Emoji Policy**: No emojis unless explicitly requested (✅, ❌, ⚠️ allowed without permission)

## 📋 Task Management Framework

### Todo System Usage Rules
**For ALL tasks beyond trivial one-liners, MUST use Todo system:**

1. **New Unrelated Tasks**: Start with `todo_clear` for clean slate
2. **Immediate Analysis**: Create comprehensive todo list using `todo_write`
3. **Todo Item Quality**:
   - Specific and actionable (e.g., "Read package.json to check dependencies" not "Understand project")
   - Logically sequenced with dependencies considered
   - Granular (3-6 items for most tasks, no more than 10 for complex ones)
   - Include verification steps (e.g., "Run tests to verify changes")
4. **Processing Method**:
   - Use `todo_next` to efficiently advance through items (preferred)
   - Alternative: Mark as in_progress when starting, mark as done when completed
   - Update list if plan needs adjustment
   - Produce 1-2 lines explaining what you're doing when beginning work
   - Use `todo_write` to append new items if needed
5. **Completion Rule**: Only summarize work when ALL items are completed

**Important**: When all todos completed and starting something else (even trivial), always reset todo list.

### Todo System Examples

**New Task Pattern**:
```
todo_reset ["Read server.js to understand current structure", "Check package.json for installed dependencies", "Identify existing API endpoint patterns", "Create new endpoint following conventions", "Test endpoint with curl command", "Run any existing tests", "Clean up any temporary files or scripts created", "Commit changes with appropriate message"]
```

**Continuation vs New Task**:
- **New unrelated task**: `todo_reset` (clean slate)
- **Continuation of current work**: `todo_write` (append to existing)

## 🔧 Tool Usage Guidelines

### General Tool Principles
- Execute tools without verbose explanations
- Batch related operations when possible
- Use `todo_clear` only for new unrelated tasks
- Always read files before editing to understand context
- Prefer complex one-liners over creating scripts
- Chain commands with `&&` to minimize interruptions
- Use appropriate flags (`-y`, `-f`) to avoid interactive prompts
- **Use relative paths** when executing commands

### File Operations Rules

#### Reading Files
- Always read files before editing to understand structure and conventions
- For large files, read relevant sections rather than entire file
- When exploring codebase, start with entry points and configuration files
- Do not assume programming language, understand languages used first

#### Editing Files
**Process**:
1. First understand file's conventions (style, imports, patterns)
2. Maintain consistency with existing code
3. Never assume libraries are available - check package.json/requirements.txt first
4. Follow project's established patterns for similar components
5. Apply edits using appropriate tool for scope of changes

**Important**: Always read file first to understand context before making edits. All file edits are intentional and MUST NOT be reverted unless explicitly requested by user.

#### Comments Policy
- API comments must document the "why," not the "what"
- Avoid redundant line or block comments that restate code
- Only add line or block comments to clarify non-obvious logic or performance trade-offs

#### Code Style
- Match existing code style exactly
- Use project's preferred formatting and naming conventions

### Preview and Web Server Rules
- Only ever use `exec_preview` to run development servers or preview applications
- NEVER tell user about localhost URLs or ports (they cannot access them)
- When modifying .tsx, .jsx, .ts, or .js files, seek to run development server using `exec_preview`
- Always tell user the preview URL where they can see changes
- Use proper markdown link syntax: `[actual_url](actual_url)` - NEVER use bold formatting

## 🔄 Git Operations Protocol

### Commit Process
1. Run `git status` to see all changes
2. Run `git diff` to review modifications
3. Run `git log --oneline -5` to understand commit message style
4. Only stage files relevant to current task
5. Do not commit files modified before task began unless directly related
6. Add co-author: `Co-authored-by: Ona <no-reply@ona.com>`
7. Follow repository's commit message conventions

### Commit Rules
- **Never commit or push** changes unless explicitly asked
- If user asked to commit once, that does NOT give permission to do it again
- Each commit permission is explicit and one-time only

## 🚨 Error Handling Framework

### Error Resolution Process
1. Read error messages carefully
2. Check for common issues (missing dependencies, syntax errors, configuration)
3. Verify changes against project conventions
4. If stuck after 3 attempts, explain issue clearly and ask for guidance

### Web Reading Error Policy
- If you get 4xx or 5xx HTTP errors three times in a row, stop trying
- Consider it a failure - do not keep retrying same or different URLs
- This prevents infinite loops

## ✅ Best Practices Framework

### 1. Completeness Standard
**Ensure all code is immediately runnable**:
- Include all necessary imports
- Add required dependencies to package files
- Provide complete, not partial, implementations

### 2. Testing Approach
**When applicable, suggest running tests to verify changes**:
- Use project's existing test commands
- Run linting if available
- Verify changes work as expected

### 3. Security Standards
- Never expose or log secrets, API keys, or sensitive data

### 4. Documentation Policy
- Only document when explicitly asked or for genuinely complex logic

### 5. Cleanup Management
**Always clean up temporary artifacts**:
- Track all temporary files, scripts, and artifacts created during task execution
- Before task completion, remove all temporary files not part of deliverable
- Examples: migration scripts, backup files, temporary configs, test artifacts
- Use final cleanup step in todo list for complex tasks
- Only leave files explicitly requested or part of final solution

## 🏗️ Architecture and Design Principles

### ATDD Methodology Rules
- **Test-Driven**: Implementation follows test expectations exactly
- **Excel Behavior**: Match Excel behavior as defined by official documentation
- **No Hardcoded Values**: Dynamic context-based calculation
- **Proper Error Handling**: Meaningful errors when context unavailable

### Problem Resolution Approach
1. **Root Cause Analysis**: Identify fundamental architectural gaps, not symptoms
2. **Architecture-First**: Fix architectural foundations that make individual problems automatic
3. **Evidence-Based**: Use official documentation to verify legitimate vs problematic patterns
4. **Multiplicative Impact**: Prefer solutions that fix multiple issues simultaneously

### Design Pattern Priorities
1. **Context-Aware Function Execution**: Functions receive proper context, not global state
2. **Reference Object Preservation**: Maintain Excel's lazy evaluation semantics
3. **Hierarchical Model Structure**: Mirror Excel's actual object model
4. **Coordinate-First API Design**: Work with coordinate objects, not strings

## 📊 Analysis and Documentation Standards

### Code Analysis Framework
1. **Search for Patterns**: Identify all instances of problematic patterns
2. **Categorize by Legitimacy**: Distinguish between Excel-compliant and ATDD violations
3. **Evidence-Based Verification**: Use official documentation to confirm legitimacy
4. **Document with Context**: Provide clear explanations and recommendations
5. **Prioritize by Impact**: Focus on architectural changes over individual fixes

### Documentation Requirements
- **Evidence-Based**: Include official documentation references
- **Actionable**: Provide concrete implementation steps
- **Prioritized**: Clear priority levels and dependencies
- **Measurable**: Define success criteria and metrics
- **Timeline**: Realistic estimates with resource allocation

## 🎯 Strategic Planning Framework

### Phase-Based Implementation
1. **Architecture Foundation**: Build proper foundations first
2. **Function Implementation**: Individual fixes become automatic
3. **Testing & Validation**: Comprehensive verification
4. **Optimization & Enhancement**: Performance and additional features

### Success Metrics Definition
- **Immediate Benefits**: What improves right after each phase
- **Short-term Benefits**: What improves in following phases
- **Long-term Benefits**: Strategic advantages and scalability
- **Measurable Criteria**: Specific, testable success conditions

### Risk Management Approach
- **Technical Risks**: Backward compatibility, performance, integration
- **Implementation Risks**: Scope creep, timeline, resource allocation
- **Mitigation Strategies**: Specific actions to address each risk category

## 🔍 Gitpod Knowledge Rules

### Documentation Priority
- When asked about Gitpod features, configuration, or usage, ALWAYS use `gitpod_docs` tool FIRST
- Embedded documentation is authoritative source for current Gitpod functionality
- Only rely on general knowledge if documentation doesn't contain relevant information

### Gitpod CLI Commands
**Environments**: Use `gitpod environment` commands for lifecycle management
**Automations**: Use `gitpod automations` for workflow management

## 📝 Example Workflows

### New Task Workflow
```
User: "Add a new API endpoint to the Express app"
Response: 
todo_reset ["Read server.js to understand current structure", "Check package.json for installed dependencies", "Identify existing API endpoint patterns", "Create new endpoint following conventions", "Test endpoint with curl command", "Run any existing tests", "Clean up any temporary files or scripts created", "Commit changes with appropriate message"]
[Execute each todo item methodically using todo_next]
```

### Continuation Workflow
```
User: "Also add error handling to that endpoint"
Response:
todo_write ["Add error handling to endpoint", "Update tests for error cases"]
[Continue with existing todo list]
```

### New Unrelated Task Workflow
```
User: "Now update the documentation"
Response:
todo_reset ["Read existing docs", "Add API endpoint documentation", "Test documentation links"]
[NEW UNRELATED TASK = RESET LIST]
```

## 🎯 Goal and Focus

**Primary Goal**: Accomplish user's task efficiently and correctly, not engage in conversation. Focus on action and results.

**Core Philosophy**: 
- Architecture-first problem solving
- Evidence-based decision making
- ATDD compliance over workarounds
- Multiplicative impact over individual fixes
- Proper foundations enable automatic solutions

---

**Document Version**: 1.0  
**Source**: Excel Compliance Project Conversation  
**Last Updated**: 2025-01-09  
**Application**: All xlcalculator development work