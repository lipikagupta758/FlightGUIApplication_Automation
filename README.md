# 🛫 FlightGUI Application Automation

This project contains automation scripts for testing the **FlightGUI Application** using **OpenText UFT (Unified Functional Testing)**. 
The automation verifies functionality of the GUI through test scripts that can be executed manually.

---

## 📌 Features

- Automated GUI testing for the Flight GUI application
- Scripted test execution using UFT
- Data-driven testing supported through Excel
- Test result reports generated in HTML format
- Results saved with timestamps in the `Results/` folder

---

## ✅ Test Cases Implemented

The automation suite currently includes the following test scenarios:

1. **Launch Application**
2. **Login to the Application**
3. **Book a Ticket**
4. **Search for an Order**
5. **Close the Application**

> Test data and configuration for each test case are provided in the `TestCasesSheet.xlsx` file located in the `TestData/` folder.

---

## 🖥️ Prerequisites

Before running the test scripts, ensure the following:

- ✅ Windows OS
- ✅ [OpenText UFT](https://www.microfocus.com/en-us/products/uft-one/overview) installed
- ✅ Microsoft Excel installed
- ✅ The Flight GUI sample application available (part of UFT installation)
- ✅ Proper permissions to run VBScript files

---

You can run the automation **in two ways**:

### ▶️ Option 1: Using the AutoRunDriverScript.vbs

1. Open **Command Prompt**.
2. Navigate to the project directory:
cd path\to\FlightGUIApplication_Automation
3. Run the script:
cscript AutoRunDriverScript.vbs

### ▶️ Option 2: Manually in UFT
Open UFT.

Use File > Open > Test and browse to the MainTest folder.

Click Run to execute the test.

After the test completes, check the Results/ folder for the HTML report.

**📄 Test Reports**
Located in the Results/ directory.

Reports are generated as HTMLReport.html.

Each run creates a new report file.

**🛠️ Customization**
You can update test data by editing the Excel file TestCasesSheet.xlsx in the TestData/ folder.

The test logic can be modified within the MainTest test actions inside UFT.
