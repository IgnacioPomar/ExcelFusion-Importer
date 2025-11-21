# ExcelFusion Importer — Functional Specification

## Overview
ExcelFusion Importer is a standalone Java 17+ desktop application built using SWT. Its purpose is to load, preview, validate, and import multiple Excel files into MariaDB or PostgreSQL.  
The application provides a multi-step wizard interface guiding users through file selection, structure validation, schema inference, database configuration, and data import.

---

## 1. General Requirements

- **Programming language:** Java 21+
- **UI framework:** SWT (standalone, included with the application)
- **Excel parsing:** Apache POI
- **Database integration:** JDBC for MariaDB and PostgreSQL
- **Packaging:** Single deployable artifact with SWT libraries
- **Project name:** ExcelFusion Importer

---

## 2. Wizard Steps

### Step 1 — File Selection

#### Features
- User selects a **directory** containing Excel `.xls` or `.xlsx` files.
- Two modes:
  1. **Single-file mode**
  2. **Auto-selection mode** with pattern filtering

#### Pattern Matching Rules
- The pattern is split into **space-separated words**.
- Matching is **case-insensitive**.
- A file matches **only if it contains all the words** in its filename.

#### Imported File Tracking
- A file named `traspasados_a_BBDD.txt` exists in the same directory.
- Contains **one filename per line**.
- These files must appear in the list:
  - With **strikethrough**
  - In **light gray**
  - **Not selectable**

---

### Step 2 — Excel Preview and Row Configuration

#### Preview
- Display a grid preview of the first rows of the first sheet.
- Provide a sheet selector:
  - **All sheets** (bold)
  - Each sheet by name

#### Header and Data Rows
- User selects:
  - **Header row number** (e.g., `10`)
  - **Data start row number** (e.g., `15`)
- Validation rule:
  - `dataStartRow > headerRow`

#### Additional Tools
- Column highlight tool by column name.
- Options:
  - **Add auto-increment column**
  - **Fill empty cells using the value from the previous row**

---

### Step 3 — Cross-file Structure Validation

#### When Header Row is Defined
- Validate that all selected **files and sheets** share the same structure:
  - Same number of columns
  - Same order
  - Same text in header row
- Display a validation report:
  - **Green** → matching `file@sheet`
  - **Red** → mismatched `file@sheet`

#### Options
- If mismatches exist, enable:
  - **Import only matching files/sheets** (default ON)

#### When Header Row is NOT Defined
- Show warning message.
- Skip structure validation.
- Use **generic column names**: `A`, `B`, `C`, …

---

### Step 4 — Column Type Inference

#### Automatic Inference
- For each column, analyze **up to 50 rows** across all selected files.
- Detect type as:
  - **Date**
  - **Currency**
  - **Integer**
  - **Text**

#### User Validation Screen
- Present a table with:
  - Column name
  - Inferred type (editable)
  - Example value
- User may override any type before continuing.

---

### Step 5 — Database Configuration

#### Supported Databases
- **MariaDB**
- **PostgreSQL**

#### Configuration Fields
- Server host
- Port
- Database name
- Username
- Password
- Checkbox: **Create database if not existing**

#### Additional Feature
- **Test Connection** button

---

### Step 6 — Import Execution

#### Progress Display
- Real-time log, e.g.:
  ```
  Processing file.xlsx@Sheet1...
  => completed
  ```
- Textual progress indicator:
  ```
  File X of N / Sheet Y of M
  ```

#### Table Creation Rules
- If the table already exists → **abort import with an error**
- Column name normalization:
  - Lowercase
  - Underscores instead of spaces
  - Remove accents/special characters
  - Auto-resolve duplicates (e.g., `_2`, `_3`)
- Column types come from the inference step
- Optional auto-increment ID column

#### Data Insertion Rules
- Insert data using JDBC batch operations.
- Fill empty cells from previous row when option is enabled.
- **Single global transaction**:
  - On any failure → rollback **everything**

#### Completion
- Append imported filenames to `traspasados_a_BBDD.txt`.

---

## 3. Configuration Persistence

- A configuration file is saved in the **same directory as the Excel files**.
- Stores:
  - Last used database settings
  - Last used directory path
- Settings are **per directory**, not global.

---

## 4. Project Structure 

```
excel-fusion-importer/
  ├── src/
  │   ├── main/java/
  │   │   ├── ui/wizard/
  │   │   ├── excel/
  │   │   ├── db/
  │   │   ├── service/
  │   │   ├── model/
  │   │   └── config/
  │   └── main/resources/
  ├── pom.xml
  ├── README.md
  ├── specification.md
  └── LICENSE
```

---

## 5. Licensing

The project uses **The Unlicense**, placing the code in the public domain.

---

## 6. Notes

This specification is designed for clarity, extensibility, and compatibility with automatic generation tools such as GitHub Copilot.  
It defines all required UI behaviors, data-processing rules, and system interactions needed to start implementing the project.
