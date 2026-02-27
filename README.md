Below is a clean, technical **README.md** in English, suitable for a professional GitHub repository.

---

# 📦 SolidWorks – Automatic Custom Property Extraction (.SLDPRT)

## 📌 Description

VBA macro for **SolidWorks** that:

* Iterates through all `.sldprt` files inside a selected folder
* Reads properties from:

  * **Custom (Document level)**
  * **Configuration-specific**
* Consolidates all data into a single `.csv` file
* Generates one row per file
* Automatically saves the output to the user's Desktop

---

## 🎯 Purpose

Automate metadata extraction for:

* Engineering control
* Data auditing
* Standardization validation
* Technical reporting
* External system integration

---

## 📊 Extracted Properties

The macro extracts the following properties:

* MACHINE
* SUBASSEMBLY NAME
* ASSEMBLY NAME
* DESIGNATION
* POSITION
* QUANTITY
* CREATION DATE
* DRAFTSMAN
* DESIGNER
* DESCRIPTION
* MATERIAL
* CODE
* UNIT WEIGHT
* DRAWING NUMBER
* REVISION
* APPROVED BY
* REPLACED BY

Property retrieval includes fallback logic:

1. Document level (`CustomPropertyManager("")`)
2. Active configuration (`CustomPropertyManager(configName)`)

---

## ⚙️ How to Use

### 1️⃣ Create the Macro

Inside SolidWorks:

```
Tools → Macro → New
```

Save as:

```
Export_Properties_Final.swp
```

Paste the provided VBA code into the editor.

---

### 2️⃣ Run the Macro

* Execute the macro
* Select the folder containing `.sldprt` files
* Wait for processing to complete

---

## 📁 Output

Generated file:

```
Properties_Exported.csv
```

Location:

```
User Desktop
```

Format:

```
FileName;Machine;SubassemblyName;...
```

Delimiter used:

```
;
```

---

## 🧠 Technical Overview

For each file, the macro:

1. Opens the part silently
2. Accesses properties via:

   ```
   ModelDoc2.Extension.CustomPropertyManager
   ```
3. Handles:

   * `valOut`
   * `resolvedValOut`
4. Applies fallback logic (Document → Configuration)
5. Closes the file
6. Writes structured output to CSV

---

## 🛡 Edge Case Handling

✔ Document-level properties
✔ Configuration-level properties
✔ Expression-based properties
✔ Plain text properties
✔ Empty fields

---

## 📌 Requirements

* SolidWorks installed
* Macro execution enabled
* `.sldprt` files with standardized custom properties

---

## ⚠ Limitations

* Does not iterate subfolders (can be extended)
* Reads only the active configuration
* Does not process assemblies (`.sldasm`)
* Outputs `.csv` only (no native `.xlsx`)

---

## 🚀 Possible Enhancements

* Recursive folder traversal
* Multi-configuration extraction
* Native Excel output (.xlsx)
* GUI-based interface
* Data validation & inconsistency reporting
* Conversion into a permanent SolidWorks Add-in

---

## 👤 Author

Developed for engineering data standardization and automation workflows.

---

If you want, I can also provide:

* A shorter public-facing README
* A more enterprise-oriented version
* Badges and project structure suggestions
* A full GitHub repository structure template
