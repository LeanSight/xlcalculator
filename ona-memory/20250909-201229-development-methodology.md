# Development Methodology & Problem Resolution Framework

**Document Version**: 3.1  
**Last Updated**: 2025-01-09  
**Application**: All software development work following ATDD methodology

---

## 1. ATDD Core Methodology

### 1.1 Double Nested Cycle Approach

#### **🔄 Outer Cycle (ATDD) - Outside-In**
- **Primary Rule**: Implementation must follow expected behavior exactly as defined by acceptance tests
- **Test-First**: Acceptance tests define business behavior, implementation follows
- **No Test Bypassing**: Never implement functionality that circumvents acceptance test expectations

#### **🔄 Inner Cycle (TDD) - Inside-Out**
- **Unit-Level TDD**: For each acceptance test failure, decompose into unit tests following Red-Green-Refactor
- **Granular Development**: Build components incrementally through unit test cycles
- **Integration Focus**: Unit tests support acceptance test fulfillment

### 1.2 Implementation Phases

#### **🔴 Red Phase (Failing Tests)**
- **Acceptance Level**: Write failing acceptance test based on business requirements
- **Unit Level**: Write failing unit tests for required components
- **No Implementation**: Code only after test fails

#### **🟢 Green Phase (Passing Tests)**
- **Minimal Implementation**: Write simplest code to make current test pass
- **Return Actual Data**: Return real data or proper error responses, no hardcoded fallbacks
- **No Premature Optimization**: Focus on making test pass first
- **📝 Immediate Commit**: Every green test triggers immediate commit and push to repository

#### **🔵 Refactor Phase (Code Improvement)**
- **Maintain Behavior**: Improve code structure without changing test outcomes
- **Eliminate ALL Duplicate Logic**: Remove every instance of code duplication by extracting common functionality
- **Functional Style Priority**: Prefer functional style over object-oriented when it results in simpler, self-documenting, and maintainable code
- **Idiomatic Code**: Use most idiomatic code for the language and version being used, following recommended best practices
- **Single Responsibility with Pareto**: If a module becomes too large, apply Single Responsibility Principle following Pareto criterion (80/20 rule) to avoid creating too many files
- **Performance Optimization**: Enhance performance without compromising test compatibility
- **Clean Code**: Apply SOLID principles and design patterns when they improve simplicity
- **📝 Immediate Commit**: Every refactor completion triggers immediate commit and push to repository

### 1.3 Quality Gates
- **All Tests Pass**: No implementation complete until both acceptance and unit tests pass
- **Living Documentation**: Tests serve as executable specifications
- **Regression Prevention**: New features cannot break existing test suite
- **Traceability**: Every commit linked to specific test phase (Red/Green/Refactor)

### 1.4 ATDD Compliance Rules
- **No Fallbacks**: Avoid fallbacks that violate expected behavior
- **No Hardcoded Values**: Dynamic context-based calculation
- **Proper Error Handling**: Meaningful errors when context unavailable
- **Specification Compliance**: Match expected behavior as defined by official documentation

---

## 2. Communication Standards

### 2.1 Tone Requirements
- **Be Concise**: Direct and technical communication, avoid conversational pleasantries
- **No Pleasantries**: Never start responses with "Great", "Certainly", "Okay", "Sure"
- **No False Assertions**: Never assert user is "absolutely right" unless certain
- **Output Necessity**: Output only what's necessary to accomplish the task
- **Command Explanation**: Briefly state what commands do and why you're running them

### 2.2 Content Policy
- **Emoji Policy**: No emojis unless explicitly requested (✅, ❌, ⚠️ allowed without permission)
- **No Emotes**: Avoid actions inside asterisks unless specifically requested
- **Evidence-Based**: Include official documentation references when applicable
- **Actionable Information**: Provide concrete implementation steps

---

## 3. Task Management Framework

### 3.1 Todo System Usage Rules
**For ALL tasks beyond trivial one-liners, MUST use Todo system:**

1. **New Unrelated Tasks**: Start with `todo_clear` for clean slate
2. **Immediate Analysis**: Create comprehensive todo list using `todo_write`
3. **Todo Item Quality**:
   - Specific and actionable (e.g., "Read package.json to check dependencies" not "Understand project")
   - Logically sequenced with dependencies considered
   - Granular (3-6 items for most tasks, no more than 10 for complex ones)
   - Include verification steps (e.g., "Run tests to verify changes")
4. **Processing Method**: Use `todo_next` to efficiently advance through items (preferred)
5. **Completion Rule**: Only summarize work when ALL items are completed

### 3.2 Task Classification
- **New unrelated task**: `todo_reset` (clean slate)
- **Continuation of current work**: `todo_write` (append to existing)
- **Complex tasks**: Always include cleanup step in todo list

### 3.3 Todo System Examples

**New Task Pattern**:
```
todo_reset ["Read existing code to understand current structure", "Check dependencies and requirements", "Identify existing patterns for similar functionality", "Create new functionality following conventions", "Test functionality", "Run any existing tests", "Save analysis to ona-memory/[timestamp]-task-analysis.md", "Clean up tmp/ directory of temporary files", "Commit changes with appropriate message"]
```

**Documentation and Cleanup Pattern**:
```
todo_write ["Document findings in ona-memory/[timestamp]-analysis-results.md", "Move temporary scripts from tmp/ to final location if needed", "Clean up all files in tmp/ directory", "Verify no temporary artifacts remain"]
```

---

## 4. File Operations & Code Standards

### 4.1 File Reading Rules
- **Always read files before editing** to understand structure and conventions
- **For large files**: Read relevant sections rather than entire file
- **Explore codebase**: Start with entry points and configuration files
- **No assumptions**: Don't assume programming language, understand languages used first

### 4.2 File Editing Process
1. **Understand Context**: Read file's conventions (style, imports, patterns)
2. **Maintain Consistency**: Match existing code style exactly
3. **Check Dependencies**: Never assume libraries are available - check package.json/requirements.txt first
4. **Follow Patterns**: Apply project's established patterns for similar components
5. **Intentional Edits**: All file edits are intentional and MUST NOT be reverted unless explicitly requested

### 4.3 Code Quality Standards

#### Comments Policy
- **API Comments**: Document the "why," not the "what"
- **Avoid Redundant Comments**: No line or block comments that restate code
- **Clarification Only**: Only add comments to clarify non-obvious logic or performance trade-offs

#### Code Style
- **Match Existing Style**: Use project's preferred formatting and naming conventions
- **Self-Documenting**: Use descriptive variable names that explain intent
- **Explicit Logic**: Prefer explicit conditionals over clever shortcuts
- **Minimal Comments**: Make code self-explanatory to minimize comments
- **Functional Style Priority**: Prefer functional style over object-oriented when it results in simpler, self-documenting, and maintainable code
- **Idiomatic Code**: Use most idiomatic code for the specific language and version being used
- **Best Practices**: Follow recommended best practices for the language and framework
- **Zero Duplication**: Eliminate all duplicate logic through extraction of common functions

### 4.4 Completeness Standard
**Ensure all code is immediately runnable**:
- Include all necessary imports
- Add required dependencies to package files
- Provide complete, not partial, implementations

---

## 5. Git Operations & Version Control

### 5.1 Commit Process
1. Run `git status` to see all changes
2. Run `git diff` to review modifications
3. Run `git log --oneline -5` to understand commit message style
4. Only stage files relevant to current task
5. Do not commit files modified before task began unless directly related
6. Add co-author: `Co-authored-by: Ona <no-reply@ona.com>`
7. Follow repository's commit message conventions

### 5.2 ATDD Git Flow Integration
```bash
# After every Green phase
git add . && git commit -m "🟢 Make [test description] pass" && git push

# After every Refactor phase  
git add . && git commit -m "🔵 Refactor [component/feature] - [improvement description]" && git push
```

### 5.3 Commit Rules
- **Never commit or push** changes unless explicitly asked
- **One-time Permission**: Each commit permission is explicit and one-time only
- **Phase Tracking**: Every commit must be linked to specific ATDD phase

---

## 6. Testing & Quality Assurance

### 6.1 Testing Approach
- **ATDD Driven**: Tests define expected behavior, implementation follows
- **Test-First**: Write failing tests before any implementation
- **Complete Coverage**: Both acceptance and unit tests required
- **Continuous Verification**: Run tests frequently during development

### 6.2 Test Structure Requirements

#### Red-Green-Refactor Cycle
1. **Red**: Write failing test that defines expected behavior
2. **Green**: Implement minimal code to make test pass
3. **Refactor**: Improve code quality while maintaining test passage

#### Test Organization
- **Acceptance Tests**: Define business behavior in `tests/acceptance/`
- **Unit Tests**: Support acceptance test fulfillment in `tests/unit/`
- **Integration Tests**: Validate component interactions in `tests/integration/`

### 6.3 Quality Verification
- **Run Project Tests**: Use project's existing test commands
- **Linting**: Run linting if available
- **Coverage**: Verify test coverage meets standards
- **No Regression**: New features cannot break existing test suite

---

## 7. Error Handling & Resolution

### 7.1 Error Resolution Process
1. **Read Error Messages**: Carefully analyze error messages
2. **Common Issues Check**: Missing dependencies, syntax errors, configuration
3. **Convention Verification**: Verify changes against project conventions
4. **Escalation**: If stuck after 3 attempts, explain issue clearly and ask for guidance

### 7.2 Web Reading Error Policy
- **Retry Limit**: If you get 4xx or 5xx HTTP errors three times in a row, stop trying
- **Failure Recognition**: Consider it a failure - do not keep retrying same or different URLs
- **Loop Prevention**: Prevents infinite loops

### 7.3 Security Standards
- **No Exposure**: Never expose or log secrets, API keys, or sensitive data
- **Safe Handling**: Secure handling of all sensitive information

---

## 8. Architecture & Design Principles

### 8.1 Problem Resolution Approach
1. **Root Cause Analysis**: Identify fundamental architectural gaps, not symptoms
2. **Architecture-First**: Fix architectural foundations that make individual problems automatic
3. **Evidence-Based**: Use official documentation to verify legitimate vs problematic patterns
4. **Multiplicative Impact**: Prefer solutions that fix multiple issues simultaneously
5. **Gap Analysis**: Distinguish between function bugs and evaluator architecture limitations

### 8.2 Solution Selection Criteria
**Primary Criteria**:
1. **Cleanliness**: Minimal, focused changes that address core issue
2. **Self-Documentation**: Code that clearly expresses intent without extensive comments
3. **Low Risk**: Changes that minimize chance of introducing new bugs
4. **Immediate Impact**: Solutions that directly fix identified problems
5. **Maintainability**: Code that is easy to understand and modify in future
6. **Zero Duplication**: Solutions that eliminate all instances of duplicate logic
7. **Functional Simplicity**: Prefer functional style over OOP when it results in simpler, more maintainable code
8. **Idiomatic Implementation**: Use most idiomatic code for the specific language and version being used

### 8.3 Design Pattern Priorities
1. **Context-Aware Function Execution**: Functions receive proper context, not global state
2. **Reference Object Preservation**: Maintain lazy evaluation semantics where appropriate
3. **Hierarchical Model Structure**: Mirror reference implementation's object model
4. **Coordinate-First API Design**: Work with structured objects, not string parsing
5. **Error Propagation Consistency**: Maintain error types through evaluation chain

### 8.4 Implementation Strategy Types
1. **Hybrid Targeted Fixes**: Combine targeted fixes without major architectural changes
2. **Architecture-First Approach**: Implement foundations that make function fixes automatic
3. **Incremental Migration**: Gradual transition with backward compatibility
4. **Collaborative Integration**: Work with upstream dependencies

---

## 9. Tool Usage & Environment

### 9.1 General Tool Principles
- **Execute Without Verbose Explanations**: Run tools directly
- **Batch Operations**: Batch related operations when possible
- **Relative Paths**: Use relative paths when executing commands
- **Minimize Interruptions**: Chain commands with `&&` 
- **Avoid Interactive Prompts**: Use appropriate flags (`-y`, `-f`)

### 9.2 Preview and Web Server Rules
- **Development Servers**: Only use `exec_preview` to run development servers or preview applications
- **No localhost URLs**: NEVER tell user about localhost URLs or ports (they cannot access them)
- **Auto-run for Changes**: When modifying .tsx, .jsx, .ts, or .js files, seek to run development server
- **Provide Preview URL**: Always tell user the preview URL where they can see changes
- **Proper Links**: Use proper markdown link syntax: `[actual_url](actual_url)` - NEVER use bold formatting

### 9.3 Environment-Specific Knowledge

#### Gitpod/Ona Rules
- **Documentation Priority**: When asked about Gitpod features, configuration, or usage, ALWAYS use `gitpod_docs` tool FIRST
- **Authoritative Source**: Embedded documentation is authoritative source for current functionality
- **Fallback Knowledge**: Only rely on general knowledge if documentation doesn't contain relevant information

#### CLI Commands
- **Environments**: Use `gitpod environment` commands for lifecycle management
- **Automations**: Use `gitpod automations` for workflow management

---

## 10. Ona/Gitpod File Management

### 10.1 Document Storage in ona-memory/

#### **Naming Convention**
- **Format**: `[timestamp]-[descriptive-name].md`
- **Timestamp**: YYYYMMDD-HHMMSS (ISO format without separators)
- **Descriptive Name**: Clear, concise description in English
- **Extension**: Always `.md` for markdown format

#### **Content Guidelines**
- **Structured Format**: Use markdown headers for organization
- **ATDD Context**: Include phase information (RED/GREEN/REFACTOR)
- **Actionable Insights**: Document recommendations and next steps
- **Evidence-Based**: Include data, metrics, and concrete observations
- **Permanent Record**: Treat as persistent documentation for future reference

### 10.2 Temporary Files in tmp/

#### **Usage Rules**
- **Session Scope**: Files exist only for current development session
- **Auto-cleanup**: May be automatically removed between sessions
- **No Permanent Data**: Never store important information only in tmp/
- **Transient Nature**: Use for debugging, testing, temporary configurations

### 10.3 File Management Workflow

#### **During Development**
1. **Create temporary files** in `tmp/` for debugging and testing
2. **Document findings** in `ona-memory/[timestamp]-[description].md`
3. **Move final scripts** from `tmp/` to appropriate project locations
4. **Clean up `tmp/`** before task completion

#### **Documentation Process**
1. **Start Analysis**: Create working files in `tmp/`
2. **Gather Data**: Use temporary scripts for analysis
3. **Document Results**: Create final report in `ona-memory/[timestamp]-analysis.md`
4. **Clean Temporary**: Remove all temporary files after documentation

#### **ATDD Integration**
- **RED Phase**: Document failing tests analysis in `ona-memory/[timestamp]-red-phase-analysis.md`
- **GREEN Phase**: Save implementation notes in `ona-memory/[timestamp]-green-implementation.md`
- **REFACTOR Phase**: Document improvements in `ona-memory/[timestamp]-refactor-summary.md`

#### **Cleanup Management**
**Always clean up temporary artifacts**:
- **Track Creation**: Track all temporary files, scripts, and artifacts created in `tmp/` during task execution
- **Pre-completion Cleanup**: Before task completion, remove all temporary files from `tmp/` not part of deliverable
- **Preserve Documentation**: Keep analysis documents in `ona-memory/` as permanent records
- **Final Step**: Use final cleanup step in todo list for complex tasks: "Clean up tmp/ directory"
- **Preserve Deliverables**: Only leave files explicitly requested or part of final solution outside `tmp/`

---

## 🎯 Success Metrics

**Primary Goal**: Accomplish user's task efficiently and correctly following ATDD methodology

**Core Philosophy**: 
- Architecture-first problem solving
- Evidence-based decision making
- ATDD compliance over workarounds
- Multiplicative impact over individual fixes
- Proper foundations enable automatic solutions

**Quality Indicators**:
- Every piece of code justified by tests
- Every commit linked to specific ATDD phase
- No implementation without corresponding tests
- Continuous integration maintained
- Clean, self-documenting code that expresses intent clearly