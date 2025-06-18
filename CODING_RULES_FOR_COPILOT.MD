# Unified Coding Rules

This document combines and enforces all code quality, structure, security, and process requirements for this project.

---

## 1. Code Quality & Structure

- **Clarity & Simplicity:**  
  Write clean, self-commenting code using clear and descriptive variable and function names. Prioritize readability and maintainability.

- **DRY (Don't Repeat Yourself):**  
  Avoid code duplication. Reuse existing code where possible and create reusable components/functions.

- **Modularity:**  
  Break large files into small, single-responsibility modules. Avoid files over 200–300 lines. Use FastAPI Routers to organize endpoints logically.

- **Prefer Existing Patterns:**  
  Always look for existing code and patterns before creating new ones. Don’t introduce new patterns unless necessary and approved.

- **Avoid Unnecessary Scripts:**  
  Don’t add standalone or one-off scripts unless they’re absolutely required. If a script will only be run once, it shouldn't be committed.

---

## 2. Python Standards & Practices

- **PEP 8 Compliance:**  
  All Python code must strictly follow the PEP 8 style guide.

- **Type Hinting:**  
  Use Python type hints for all function signatures and relevant variables.

- **Environment Awareness:**  
  Write code with clear separation for development, testing, and production environments. Never mock or fake data in dev or prod.

---

## 3. Error Handling & Logging

- **Robust Error Handling:**  
  Implement clear and structured error handling. Return meaningful HTTP status codes and error messages.

- **Structured Logging:**  
  Use the `logging` module with JSON formatting. Always log to standard output.

- **Maintain Pattern Consistency:**  
  Do not introduce new technology or paradigms when fixing issues unless all other options are exhausted. If new logic is adopted, remove legacy code to avoid duplication.

---

## 4. Security & Validation

- **Sanitize All Input:**  
  Validate and sanitize user input on the backend.

- **Prevent SQL Injection:**  
  Use parameterized queries (SQLAlchemy manages this by default).

- **Implement Auth Checks:**  
  Ensure proper authentication and authorization checks are in place for all protected endpoints.

---

## 5. API Development

- **OpenAPI Documentation:**  
  Use `tags` and `description` in FastAPI route decorators to produce high-quality auto-generated API documentation.

- **Limit Scope of Changes:**  
  Only modify code related to the requested task. Avoid touching unrelated areas of the codebase.

---

## 6. Testing & Mocking

- **Write Thorough Tests:**  
  IMPORTANT: Provide full unit and integration test coverage, aim for 100% coverage. Use `pytest` or the project-standard testing framework. Work is only complete when all test are written and run withput any errors!

- **Mock Only in Tests:**  
  Never include stubbed, mocked, or fake data outside of the testing environment.

---

## 7. Deployment & Execution

- **Don’t Overwrite .env Files:**  
  Never modify or overwrite the `.env` file without explicit permission.

---

## Final Notes

- **Keep the Codebase Clean:**  
  Maintain a tidy, organized structure throughout the project.

- **Be Conservative with Changes:**  
  Avoid major architectural changes to well-functioning code unless explicitly instructed.

- **Consider Ripple Effects:**  
  Think through how your changes may impact other methods, modules, and workflows.
