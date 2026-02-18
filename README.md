# 🛠️ System Inventory Tool (Automation Script)

### 💡 Overview
**System Inventory Tool** is a robust Python-based automation script designed to streamline hardware auditing and asset management. This project was inspired by my real-world experience managing **150+ workstations** during my IT cooperative training at **Arab Open University**, where I identified the need to replace manual data entry with an automated, error-free solution.

### 🚀 Key Features
* **Smart "Update or Add" Logic:** Automatically detects if a device is already in the registry based on the Device Name. It updates technical specs while preserving manual entries like "Physical Location" and "Notes".
* **Automated Data Collection:** Extracts critical hardware information including:
    * Manufacturer & Model (e.g., HP EliteBook, Dell Latitude)
    * IP Address
    * Total RAM (Rounded to the nearest GB)
    * Disk Space (C: Drive capacity)
* **Professional Reporting:** Generates a beautifully formatted Excel file (`Master_IT_Inventory.xlsx`) with:
    * Color-coded headers (Professional Blue).
    * Auto-adjusted column widths for better readability.
    * A "Status" column to track whether a device was "Added" or "Updated".
* **Precise Audit Trail:** Includes a "Last Scan" timestamp (Date & Time) for every entry.
* **Stand-alone Executable:** Can be converted to an `.exe` file to run on any Windows machine without requiring a Python installation.

### 🛠️ Built With
* **Python 3.x**
* **Pandas & Openpyxl:** For advanced data manipulation and Excel formatting.
* **Psutil & Subprocess:** For deep system interfacing and hardware data retrieval.

### 📥 How to Run

#### Option 1: For Users (Standalone App)
1.  Navigate to the **Releases** section on the right.
2.  Download the latest `Inventory_Tool.exe`.
3.  Run the file on any Windows workstation.
4.  The `Master_IT_Inventory.xlsx` will be generated/updated in the same directory.

#### Option 2: For Developers (Source Code)
1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/Waad-CS/System-Inventory-Tool.git](https://github.com/Waad-CS/System-Inventory-Tool.git)
    ```
2.  **Install dependencies:**
    ```bash
    pip install pandas psutil openpyxl
    ```
3.  **Run the script:**
    ```bash
    python main.py
    ```

### 📝 Author
**Waad B.** *Computer Science Graduate | IT Support Specialist* Jeddah, Saudi Arabia
