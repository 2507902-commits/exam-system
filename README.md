# Online Examination System (VB.NET, Windows Forms, MS Access)

Open **`OnlineExamSystem.sln`** in Visual Studio (2019/2022, “.NET desktop development”
workload) and press **F5**.

* The Access database (`Data\ExamSystem.accdb`) is created automatically on first run,
  together with the demo accounts and 10 sample questions.
* If Visual Studio asks about the .NET Framework 4.8 targeting pack, install it.
* If you see “provider 'Microsoft.ACE.OLEDB.12.0' is not registered”, install the
  **Microsoft Access Database Engine 2016 Redistributable** (same bitness as the app), or
  switch *Project → Properties → Compile → Target CPU* to **x86**.

## Demo accounts

| Username | Password | Role |
|---|---|---|
| admin | admin123 | Administrator |
| student1 | student123 | Student |
| student2 | student123 | Student |

## Student flow

Log in → student dashboard → **START EXAM** → answer the multiple-choice questions with the
radio buttons → **Submit** (or let the 30-minute timer run out) → the exam is marked
automatically and the score, percentage and grade are saved to Access.

## Administrator flow

Log in as `admin` → manage the question bank, add students / reset passwords, and view,
search or export every result to CSV.

Full explanation of the design, database schema, marking algorithm and troubleshooting:
see **`DOCUMENTATION.md`**.
