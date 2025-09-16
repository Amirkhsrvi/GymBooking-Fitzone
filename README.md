# Fit Zone Gym Management Spreadsheet

A complete Excel-based management system designed for a gym business called **Fit Zone**.  
The application supports **bookings, invoices, member and staff management, wage calculations, and revenue monitoring**.  
It uses a combination of **formulas, data validation, macros, and VBA** to provide a user-friendly and efficient experience.

---

## 📌 Features

### 🏠 Home Sheet
- Branded with a logo and current date (`=TODAY()` and `TEXT()`).
- Navigation buttons (macros) to all key worksheets.

### 📅 Booking System
- Book sports activities with **member ID, time, activity, and instructor option**.  
- Automatic validation: dropdowns for activities & times, digits-only for IDs.  
- Prevents overbooking with **conditional formatting**.  
- Maintenance table shows unavailable activities.  
- VBA macros handle **submit, clear, and navigation buttons**.

### 🧾 Invoices
- Auto-generated invoice numbers (`=MAX()+1`).  
- Member details retrieved via `VLOOKUP`.  
- Discounts applied:
  - **Off-peak discount** (5% between 9–12).  
  - **OPA discount** (35% for senior members).  
- Instructor costs added if required.  
- Printable invoice with buttons to **print, clear, and navigate**.  

### 👥 Member Management
- Add new members with auto-generated IDs.  
- Calculate **youngest, oldest, and average age** (`=MAX()`, `=MIN()`, `=AVERAGE()`).  
- Export records to PDF.  
- Filtering and navigation buttons included.

### 🏋️ Activities
- List of activities with hourly costs.  
- Maintenance schedule with `COUNTIFS` formula.  
- Hidden from general users, admin-only access.  

### 👨‍💼 Staff & Wages
- Staff details (personal info, roles, rates).  
- Wage calculator:
  - `VLOOKUP` retrieves staff info.  
  - `IF` formulas calculate weekly wages and overtime.  
- Data validation ensures correct inputs (ID, hours, etc.).  
- Wages pivot table summarizes costs by role.  
- Export to PDF with macros.

### 📊 Revenue & Statistics
- Monthly income/expenses tracked with `SUMIFS`.  
- Automatic profit/loss detection with conditional formatting.  
- Summary charts: clustered columns for revenue, expenses, profit.  

---

## 🛠️ Tech & Tools
- **Microsoft Excel**  
- **Formulas**: `VLOOKUP`, `IF`, `SUMIFS`, `COUNTIFS`, `TEXT`, `CHOOSE`, `MAX`, `MIN`, `AVERAGE`.  
- **Macros & VBA**: navigation, data submission, invoice generation, PDF export.  
- **Pivot Tables**: wage summaries.  
- **Conditional Formatting**: overbooking prevention, financial status alerts.  
- **UI Design**: color scheme (#E7E6E6 and #67A97B), buttons, clear layout.

---

## 🎯 Learning Outcomes
- Built a **fully functional booking and management system** in Excel.  
- Applied advanced formulas, data validation, and error prevention.  
- Designed **event-driven macros** for navigation and automation.  
- Implemented a **discount system and wage calculator**.  
- Practiced professional UI/UX design within Excel.  

---

## 🔮 Future Improvements
- Add a **user guide sheet** for first-time users.  
- More detailed wage system (holidays, days off, bonus types).  
- Improved staff data entry form.  
- Enhanced reporting features.  

---

## 📖 Evaluation
This project demonstrates how Excel can be used as a **comprehensive business application**.  
It improves booking efficiency, prevents errors, records financial data, and helps management with decision-making.  
The system balances **user-friendliness, functionality, and business needs** effectively. 
