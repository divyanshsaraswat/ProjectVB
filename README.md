# VB6 Project Overview

## Entry Point

This section outlines the flow of the application starting from the login to the main interface and subsequent forms.

### 1. `frmLogin.frm`
- The initial form displayed when the application starts.
- Handles user authentication and login validation.

### 2. `MDIForm1.frm`
- Acts as the **Main Application Window**.
- Contains the **Menu Bar** which provides access to all major features.
- Serves as the entry point for opening other child forms within the MDI interface.

### 3. [Next Form Placeholder]
- *(Add description of the next form or module in the application here)*

---

## Structure

The project follows a typical VB6 pattern:

- `.frm` files represent **Forms**
- `.bas` files are **Modules** (shared functions, global variables)
- `.cls` files are **Class Modules**
- `.vbp` is the **Project File**
- `.frx` files hold **binary form data** (like images, layout presets, etc.)

---

## Migration Plan (Optional Section)

If you're migrating this to **C# .NET Framework**, consider:

- Mapping each `.frm` to a corresponding `Form` class
- Replacing `MSFlexGrid` with `DataGridView`
- Using `MenuStrip` for menu bars
- Managing database connections with `OleDbConnection`

---

> ⚠️ You can expand this README as you uncover new forms or modules and start planning the migration.
