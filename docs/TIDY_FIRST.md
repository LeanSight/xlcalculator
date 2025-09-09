# Tidy First - Coding Agent Guide

## Core Principles

### What is Tidying?
Tidying involves making small, incremental changes to the **structure** (not behavior) of code to improve:
- **Readability** - Making code easier to understand
- **Maintainability** - Reducing future maintenance costs
- **Flexibility** - Enabling easier future changes

### Key Rules
1. **Separate Structure from Behavior** - Structural changes should be distinct from behavioral changes
2. **Incremental Improvement** - Small steps are safer and more manageable
3. **Human-Centered Design** - Code is written for humans to read
4. **Time Limit** - Maximum 1 hour of tidying before making behavioral changes

---

## The 15 Tidying Techniques

### 1. Guard Clauses
**Purpose**: Reduce nesting and improve readability by handling edge cases early.

**Before:**
```javascript
function processUser(user) {
    if (user) {
        if (user.isActive) {
            if (user.hasPermission) {
                return doSomething(user);
            }
        }
    }
    return null;
}
```

**After:**
```javascript
function processUser(user) {
    if (!user) return null;
    if (!user.isActive) return null;
    if (!user.hasPermission) return null;
    
    return doSomething(user);
}
```

### 2. Dead Code
**Purpose**: Remove unused code that adds confusion and maintenance burden.

**Actions:**
- Delete unused functions, variables, imports
- Remove commented-out code
- Eliminate unreachable code paths

### 3. Normalize Symmetries
**Purpose**: Make similar operations follow the same pattern.

**Before:**
```python
if status == "active": return True
elif status == "inactive": return False
else: raise Exception("Invalid status")
```

**After:**
```python
status_map = {"active": True, "inactive": False}
if status not in status_map:
    raise Exception("Invalid status")
return status_map[status]
```

### 4. New Interface, Old Implementation
**Purpose**: Create the interface you wish you had, delegating to existing implementation.

```python
# Create the interface you want
def calculate_user_score(user):
    return legacy_score_calculation(user.data, user.preferences, user.history)

# Use the new, cleaner interface
score = calculate_user_score(current_user)
```

### 5. Reading Order
**Purpose**: Arrange code so it reads naturally from top to bottom.

**Guidelines:**
- Put main logic first
- Helper functions follow
- Constants and configuration at the top or bottom

### 6. Cohesion Order
**Purpose**: Group related elements together.

```python
# Group related user operations
class UserManager:
    def create_user(self): pass
    def update_user(self): pass
    def delete_user(self): pass
    
    def send_welcome_email(self): pass
    def send_notification(self): pass
```

### 7. Move Declaration and Initialization Together
**Purpose**: Reduce cognitive load by keeping variable declaration close to usage.

**Before:**
```javascript
let result;
let count;
let data;

// 20 lines of other code

result = processData();
count = result.length;
data = result.items;
```

**After:**
```javascript
// 20 lines of other code

const result = processData();
const count = result.length;
const data = result.items;
```

### 8. Explaining Variables
**Purpose**: Extract complex expressions into well-named variables.

**Before:**
```python
if user.age >= 18 and user.country in ['US', 'CA', 'UK'] and user.verified:
    allow_access()
```

**After:**
```python
is_adult = user.age >= 18
is_supported_country = user.country in ['US', 'CA', 'UK']
is_verified_user = user.verified

if is_adult and is_supported_country and is_verified_user:
    allow_access()
```

### 9. Explaining Constants
**Purpose**: Replace magic numbers and strings with named constants.

**Before:**
```python
if response.status_code == 404:
    handle_not_found()
```

**After:**
```python
HTTP_NOT_FOUND = 404

if response.status_code == HTTP_NOT_FOUND:
    handle_not_found()
```

### 10. Explicit Parameters
**Purpose**: Make implicit dependencies explicit through parameters.

**Before:**
```python
def calculate_tax():
    # Implicitly uses global TAX_RATE
    return price * TAX_RATE
```

**After:**
```python
def calculate_tax(price, tax_rate):
    return price * tax_rate
```

### 11. Chunk Statements
**Purpose**: Use blank lines to group related statements.

```python
# Input validation
if not user_id:
    raise ValueError("User ID required")
if not isinstance(user_id, int):
    raise ValueError("User ID must be integer")

# Database operation
user = database.get_user(user_id)
if not user:
    raise NotFoundError("User not found")

# Business logic
if user.is_premium():
    apply_premium_features(user)
else:
    apply_standard_features(user)
```

### 12. Extract Helper
**Purpose**: Create helper functions to name and isolate logical operations.

**Before:**
```python
def process_order(order):
    # Validate order
    if not order.items or len(order.items) == 0:
        raise ValueError("Empty order")
    
    total = sum(item.price * item.quantity for item in order.items)
    if total < 0:
        raise ValueError("Invalid total")
    
    # Apply discounts
    if order.customer.is_premium:
        total *= 0.9
    if total > 100:
        total -= 10
    
    return total
```

**After:**
```python
def process_order(order):
    validate_order(order)
    total = calculate_total(order)
    total = apply_discounts(total, order.customer)
    return total

def validate_order(order):
    if not order.items or len(order.items) == 0:
        raise ValueError("Empty order")

def calculate_total(order):
    total = sum(item.price * item.quantity for item in order.items)
    if total < 0:
        raise ValueError("Invalid total")
    return total

def apply_discounts(total, customer):
    if customer.is_premium:
        total *= 0.9
    if total > 100:
        total -= 10
    return total
```

### 13. One Pile
**Purpose**: Sometimes bringing scattered code together helps understanding.

**When to use:**
- Code is split across multiple files/functions but separation hinders understanding
- Temporarily consolidate to see the full picture, then re-organize appropriately

### 14. Explaining Comments
**Purpose**: Add comments when code logic is complex or non-obvious.

```python
# Using binary search because dataset is pre-sorted and can be large (>10k items)
def find_user_index(users, target_id):
    left, right = 0, len(users) - 1
    # ... binary search implementation
```

### 15. Delete Redundant Comments
**Purpose**: Remove comments that don't add value.

**Remove:**
```python
# Increment counter by 1
counter += 1

# Return the result
return result
```

**Keep:**
```python
# Compensate for leap year calculation quirk in legacy system
adjusted_date = date + timedelta(days=1)
```

---

## When to Apply Tidying

### Tidy First When:
- You need to understand code before changing it
- The change will be easier after structural improvements
- The tidying takes less time than the benefit it provides

### Tidy After When:
- You're going to change the same area again soon
- You want to leave the code better than you found it
- You have time after completing the main work

### Don't Tidy When:
- You'll never touch the code again
- Time pressure is extreme
- The code works and changes are risky

---

## Application Guidelines

### Tidying Decision Process:
1. **Identify the need** - Does the code structure hinder understanding or change?
2. **Choose technique** - Which tidying technique addresses the issue?
3. **Apply incrementally** - Make small, safe changes
4. **Verify behavior unchanged** - Run tests to ensure no behavioral changes
5. **Continue or stop** - Based on time investment vs. value

### Integration with ATDD:
- **RED Phase**: Apply tidying to understand existing code before writing tests
- **GREEN Phase**: Minimal tidying only if it helps implementation
- **REFACTOR Phase**: Primary time for applying tidying techniques

### Quality Indicators:
- Code reads more naturally
- Complex expressions are explained
- Related code is grouped together
- Magic numbers/strings are named
- Duplicate logic is eliminated
- Functions have single, clear purposes

---

## Quick Reference

**Most Common Tidying Techniques:**
1. **Guard Clauses** - Reduce nesting
2. **Explaining Variables** - Name complex expressions
3. **Extract Helper** - Create focused functions
4. **Dead Code** - Remove unused code
5. **Explaining Constants** - Name magic values

**Red Flags Requiring Tidying:**
- Deep nesting (>3 levels)
- Long functions (>20 lines)
- Magic numbers/strings
- Duplicate code blocks
- Complex conditional expressions
- Unclear variable names
- Mixed levels of abstraction

**Remember:** The goal is code that's easier to understand and change, not perfect code.