# Online Examination System — VB.NET (Windows Forms) + Microsoft Access

A complete, working multiple-choice examination system written in **VB.NET** with a
**Windows Forms GUI** and a **Microsoft Access database backend**.

Students log in, sit a timed exam with **radio-button answers**, and are **marked
automatically** the moment they submit (or when the timer reaches `00:00`).
Administrators manage questions, students and results from a separate dashboard.

---

## 1. What is included

| Feature | Where it lives |
|---|---|
| Login with username/password and two roles (Administrator / Student) | `UI\LoginForm.vb` |
| Student dashboard: exam rules, question count, personal results | `UI\StudentDashboardForm.vb` |
| Exam screen: one question at a time, 4 radio buttons, Previous/Next, Submit | `UI\ExamForm.vb` |
| Live countdown timer (turns red in the last 5 minutes, auto-submits at 00:00) | `UI\ExamForm.vb` (`examTimer`, `UpdateClockLabel`) |
| Automatic marking, percentage and grade | `Logic\MarkingEngine.vb` |
| Result screen with score, grade, time used, question-by-question breakdown | `UI\ExamSummaryForm.vb` |
| Administrator dashboard (stats, questions, students, results) | `UI\AdminDashboardForm.vb` |
| Add / edit / delete / activate / deactivate questions | `UI\QuestionEditorForm.vb` |
| Add students, reset passwords, delete students | `UI\AdminDashboardForm.vb` |
| Search results and export them to CSV | `UI\AdminDashboardForm.vb` |
| Automatic creation of the Access database, tables and demo data | `Data\DatabaseHelper.vb` |
| All SQL (parameterised) | `Data\DatabaseHelper.vb` |

The automatic marking engine has been unit-tested: all-correct, partially-correct,
unanswered, grade bands and case-insensitive answers all produce the expected marks.

---

## 2. Requirements

* **Visual Studio 2019 or 2022** with the **“.NET desktop development”** workload.
* **.NET Framework 4.8** targeting pack (usually already installed; otherwise the
  Visual Studio Installer offers it).
* A **Microsoft Access database engine** on the PC:
  * If Microsoft Office (32-bit or 64-bit) is installed, you already have it.
  * If not, install the free **Microsoft Access Database Engine 2016 Redistributable**
    from Microsoft’s download centre — choose the same bitness as the app
    (64-bit for the default **AnyCPU** build, 32-bit if you switch the project to **x86**).

> The app creates its database automatically — you never have to design tables by hand.

---

## 3. Quick start in Visual Studio

1. **Unzip** `OnlineExamSystem.zip` anywhere (e.g. `Documents\OnlineExamSystem`).
2. Double-click **`OnlineExamSystem.sln`** (or *File → Open → Project/Solution*).
3. In **Solution Explorer** you should see the project **OnlineExamSystem** with folders
   `Data`, `Logic`, `UI`.
4. Press **F5** (Start). The first run:
   * creates a folder `bin\Debug\Data\`,
   * creates `ExamSystem.accdb`, its four tables, three user accounts and 10 demo questions,
   * opens the login window.
5. Log in with one of these demo accounts:

| Username | Password | Role |
|---|---|---|
| `admin` | `admin123` | Administrator |
| `student1` | `student123` | Student (John Kamau) |
| `student2` | `student123` | Student (Mary Wanjiku) |

6. As a student: click **START EXAM**, answer with the radio buttons, press **Submit**
   — the score, percentage and grade appear immediately and are stored in Access.
7. As the admin: manage questions/students and view or export all results.

> **Tip:** you can open the forms visually — in Solution Explorer double-click
> `LoginForm.vb`, `ExamForm.vb`, etc. The Visual Studio **form designer** shows the real layout
> because the `.Designer.vb` files are the standard designer format.

---

## 4. The database

### 4.1 Where is it?

```
...\OnlineExamSystem\bin\Debug\Data\ExamSystem.accdb      (64-bit build / Access installed)
...\OnlineExamSystem\bin\Debug\Data\ExamSystem.mdb        (fallback: 32-bit build, Jet only)
```

The programme chooses `.accdb` with the **ACE** provider, and falls back to a `.mdb` file
with the built-in **Jet** provider when ACE is not available (32-bit builds).

* The folder and file are created on first run.
* **Reset the demo database:** close the app and delete the `Data` folder — it is
  rebuilt with the demo users and questions on the next start.
* **Keep your results:** copy `ExamSystem.accdb` somewhere safe before using
  *Clean Solution* or deleting `bin`.
* You can open the file directly in **Microsoft Access** to browse the tables.

### 4.2 Tables

**tblUsers** — everyone who can log in

| Field | Type | Notes |
|---|---|---|
| UserID | AutoNumber | primary key |
| Username | Text(50) | unique index `idxUsers_Username` |
| Password | Text(50) | stored as typed (see §9 for hashing) |
| FullName | Text(100) | shown on the dashboards |
| Role | Text(20) | `Administrator` or `Student` |

**tblQuestions** — the question bank

| Field | Type | Notes |
|---|---|---|
| QuestionID | AutoNumber | primary key |
| QuestionText | Memo | the question |
| OptionA … OptionD | Text(255) | up to four choices |
| CorrectOption | Text(1) | `A`, `B`, `C` or `D` |
| Marks | Integer | marks for this question |
| IsActive | Yes/No | inactive questions are not given to students |

**tblResults** — one row per exam attempt

| Field | Type | Notes |
|---|---|---|
| ResultID | AutoNumber | primary key |
| UserID | Integer | link to `tblUsers` |
| StudentName | Text(100) | copied at the time of the exam |
| Score | Integer | marks earned |
| TotalMarks | Integer | marks available |
| Percentage | Double | e.g. 70.0 |
| Grade | Text(2) | A–F |
| DateTaken | Date/Time | when the exam was submitted |
| TimeTakenSeconds | Integer | seconds actually used |

**tblAnswers** — one row per question per attempt (the answer sheet)

| Field | Type | Notes |
|---|---|---|
| AnswerID | AutoNumber | primary key |
| ResultID | Integer | link to `tblResults` |
| QuestionID | Integer | link to `tblQuestions` |
| QuestionText | Memo | copy of the question text (keeps old reports readable) |
| SelectedOption | Text(1) | what the student chose (empty = unanswered) |
| CorrectOption | Text(1) | correct answer at the time of the exam |
| IsCorrect | Yes/No | result of automatic marking |
| Marks | Integer | marks that question was worth |

---

## 5. How it works — process by process

### 5.1 Startup sequence

```
Program.Main()
   │
   ├─ Application.EnableVisualStyles()
   ├─ New DatabaseHelper()
   │     └─ ensures ...\bin\Debug\Data exists
   ├─ EnsureDatabase()
   │     ├─ file exists?      → open it
   │     ├─ no file?          → create .accdb with ADOX (ACE provider)
   │     │                      or .mdb with Jet if ACE is missing
   │     ├─ HasSchema()?      → if tables are missing, run CREATE TABLE scripts
   │     └─ SeedData()        → demo users + 10 questions
   │
   └─ Do
        show LoginForm (modal)
          └─ OK?  → open AdminDashboardForm or StudentDashboardForm (modal)
                     └─ “Log Out” returns DialogResult.Retry → back to login
      Loop
```

Relevant code: `Program.vb` (`Main`), `Data\DatabaseHelper.vb`
(`EnsureDatabase`, `CreateNewDatabase`, `ExecuteSchema`, `SeedData`).

If the database cannot be opened, a message box explains the cause (usually the
missing Access database engine) and how to fix it — see §8.

### 5.2 Login process

1. The student/admin types a username and password.
2. `LoginForm.TryLogin()` validates that both fields are filled.
3. `DatabaseHelper.Authenticate()` runs
   `SELECT … FROM tblUsers WHERE [Username] = ? AND [Password] = ?` (parameterised).
4. No row → “Invalid username or password”.
5. A row → an `ExamUser` object is returned, the form closes with `DialogResult.OK`,
   and `Program.Main` opens the dashboard that matches **`Role`**.
6. `chkShowPassword` simply toggles `UseSystemPasswordChar`.

### 5.3 Student dashboard

* `RefreshDashboard()` loads the active questions (`GetExamQuestions()`) and adds up the
  marks, so the label `10 questions | 20 marks | 30 minutes | pass mark 50%` is always true
  even after the administrator edits questions.
* The same call loads the student’s own history (`GetResultsForUser`) into the grid and
  computes the three cards: **exams taken**, **best score**, **average score**.
* **START EXAM** asks for confirmation, then opens `ExamForm` with the question list and
  refreshes the dashboard when it closes.

### 5.4 The exam screen (radio buttons + timer)

```
ExamForm_Load
   ├─ remaining = 30 minutes × 60
   ├─ examTimer.Interval = 1000 ms ; Start()
   └─ ShowQuestion(0)

ShowQuestion(index)
   ├─ writes question text, marks and “QUESTION n OF m”
   ├─ copies the four option texts into rbOptionA…rbOptionD
   ├─ hides an option that is empty and re-flows the remaining ones
   ├─ ticks the radio button that matches a previously saved answer
   └─ updates the progress bar and “Answered: x of y”

User clicks a radio button
   └─ OptionCheckedChanged: answers(QuestionID) = "A"/"B"/"C"/"D"

examTimer_Tick  (every second)
   ├─ remaining −= 1 ; label shows mm:ss
   ├─ last 5 minutes → timer turns red
   └─ remaining = 0 → message “Time is up!” → FinishExam()
```

`_isLoadingQuestion` is a guard flag: when the code ticks a radio button while loading a
question, the `CheckedChanged` event fires but must **not** overwrite the stored answer.
Only real user clicks are recorded.

The **Submit** button warns if some questions are unanswered, then calls `FinishExam()`.
Closing the window with **X** also asks “submit and exit?” — an exam cannot be abandoned
without being marked.

### 5.5 Automatic marking

`FinishExam()` calls `MarkingEngine.Mark(questions, answers)`:

```
FOR EACH question
    selected = answers(question.QuestionID)      ' "" when unanswered
    IF selected = "" THEN
        unanswered += 1 ; marks earned = 0
    ELSE IF selected = question.CorrectOption (ignoring case) THEN
        score += question.Marks ; correct += 1
    ELSE
        wrong += 1 ; marks earned = 0
    END IF
    totalMarks += question.Marks
NEXT
percentage = round(score × 100 / totalMarks, 1)
grade      = GradeFor(percentage)
```

Grade bands (`MarkingEngine.GradeFor`): **A ≥ 80**, **B ≥ 70**, **C ≥ 60**,
**D ≥ 50**, **E ≥ 40**, **F < 40**. Pass mark: 50 % (`AppSettings.PassMarkPercentage`).

**Worked example** — 10 questions × 2 marks = 20 marks:

| Correct | Wrong | Unanswered | Score | Percentage | Grade | Result |
|---|---|---|---|---|---|---|
| 10 | 0 | 0 | 20 / 20 | 100 % | A | Passed |
| 7 | 2 | 1 | 14 / 20 | 70 % | B | Passed |
| 4 | 5 | 1 | 8 / 20 | 40 % | E | Failed |

Unanswered questions score 0 but **still count** in the total, so guessing is never worse
than leaving a blank.

### 5.6 Saving the result

`DatabaseHelper.SaveResult()` writes everything inside **one transaction**:

1. `INSERT INTO tblResults (UserID, StudentName, Score, TotalMarks, Percentage, Grade, DateTaken, TimeTakenSeconds)`
2. `SELECT @@IDENTITY` → the new `ResultID`
3. one `INSERT INTO tblAnswers` per question (question text, chosen option, correct option,
   whether it was right, marks)
4. `COMMIT` — or `ROLLBACK` when anything fails, so a half-saved result is impossible.

The result screen (`ExamSummaryForm`) is then shown with score, percentage, grade,
time used and a colour-coded breakdown of every question (your answer vs correct answer).

### 5.7 Administrator — questions

* **Add / Edit** open `QuestionEditorForm` (text, options A–D, correct option, marks, active).
* Validation: question text ≥ 5 characters, options A and B required, the chosen correct
  option must not be empty, marks 1–100.
* **Delete** removes the question from the bank; past answer sheets keep their own copy of
  the question text, so old reports never break.
* **Activate / Deactivate** flips `tblQuestions.IsActive` — deactivated questions are
  hidden from students but kept for history.
* Double-clicking a row edits it.

### 5.8 Administrator — students

* **Add Student** → `INSERT INTO tblUsers … Role = 'Student'` with validation
  (name, username `3–30` letters/numbers/underscore, password ≥ 5 characters) and a
  friendly message if the username already exists (enforced by the unique index).
* **Reset Password** → asks for a new password (`InputBox`) and runs `UPDATE tblUsers`.
* **Delete Student** → deletes the login only; results stay for reporting.

### 5.9 Administrator — results

* All attempts are listed newest first with student, score, percentage, grade, date and time.
* The **Search** box filters the list (student name, username or grade) using
  `DataTable.DefaultView.RowFilter`.
* **Export to CSV** writes exactly the rows you can see to a `.csv` file that opens in
  Excel — handy for handing marks to a lecturer.

---

## 6. Code map

```
OnlineExamSystem.sln                 Visual Studio solution
└─ OnlineExamSystem\
   ├─ OnlineExamSystem.vbproj        project (net48, WinExe, Option Strict On)
   ├─ App.config
   ├─ Program.vb                     Main(): database check → login loop → dashboards
   ├─ AppSettings.vb                 exam title, time limit, pass mark, default marks
   ├─ Data\
   │  ├─ Models.vb                   ExamUser, ExamQuestion, StudentAnswer, DashboardStats
   │  ├─ DatabaseHelper.vb           ALL Access SQL, database creation + seed data
   │  └─ ComInteropHelper.vb         ADOX call that creates a new .accdb / .mdb
   ├─ Logic\
   │  └─ MarkingEngine.vb            automatic marking + grade bands
   └─ UI\
      ├─ Theme.vb                    colours, fonts, grid styling, message helpers
      ├─ LoginForm.vb / .Designer.vb / .resx
      ├─ StudentDashboardForm.vb / …
      ├─ ExamForm.vb / …             radio buttons + countdown timer
      ├─ ExamSummaryForm.vb / …      instant result + answer breakdown
      ├─ AdminDashboardForm.vb / …   4 tabs: dashboard, questions, students, results
      └─ QuestionEditorForm.vb / …   add / edit question
```

Every form has three files: the code (`FormName.vb`), the generated layout
(`FormName.Designer.vb`) and its resource file (`FormName.resx`) — exactly what Visual
Studio’s designer expects.

---

## 7. Customising the system

| What to change | Where |
|---|---|
| Exam title | `AppSettings.ExamTitle` |
| Time limit (minutes) | `AppSettings.ExamTimeLimitMinutes` |
| Pass mark | `AppSettings.PassMarkPercentage` |
| Grade bands | `MarkingEngine.GradeFor()` |
| Default marks in the editor | `AppSettings.DefaultQuestionMarks` |
| Colours / fonts | `UI\Theme.vb` |
| Questions | Admin dashboard → **Questions** tab (or edit `SeedData` before the first run) |
| Users | Admin dashboard → **Students** tab (admins are seeded in `SeedData`) |

Adding more questions through the UI is the intended way — the exam length, total marks
and dashboard figures update automatically.

---

## 8. Troubleshooting

| Message / problem | Cause | Fix |
|---|---|---|
| *“The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.”* | 64-bit app, 32-bit Office/ACE (or no Access engine) | Install the **Access Database Engine 2016 Redistributable** matching the build (64-bit), **or** Project → Properties → Compile → **Target CPU = x86** |
| *ADOX is not available…* | Office/ACE not installed | Install the Access Database Engine (or full Office) |
| *“The application could not create or open its database”* | See the details in the dialog | Follow the three hints printed in the dialog |
| Database locks / “file in use” | Another copy of the app or Access has the file open | Close the other program |
| I want to start over | — | Close the app, delete `bin\Debug\Data`, run again |
| My results disappeared after **Clean Solution** | `bin` was deleted (database lives in `bin\Debug\Data`) | Back up `ExamSystem.accdb` before cleaning; or change the folder in `DatabaseHelper.New()` |
| Designer shows an error instead of the form | Project not fully loaded / build error | Build the project first (`Ctrl+Shift+B`), then reopen the designer |
| *“The project targets .NET Framework 4.8 which is not installed”* | Missing targeting pack | In Visual Studio Installer add **.NET Framework 4.8 targeting pack** |
| *Operation must use an updateable query* | The `Data` folder is read-only (e.g. in Program Files) | Run from a user folder, or give write permission |

---

## 9. Security notes (and how to improve them)

This is a teaching project, so a few things are intentionally simple:

1. **Passwords are stored as plain text.** For production, store a hash. Small example
   to drop into `DatabaseHelper`:

   ```vb
   Imports System.Security.Cryptography
   Imports System.Text

   Private Shared Function HashPassword(ByVal password As String) As String
       Using sha As SHA256 = SHA256.Create()
           Dim bytes As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(password))
           Return Convert.ToBase64String(bytes)
       End Using
   End Function
   ```

   Use `HashPassword(password)` in `AddStudent`, `UpdateUserPassword` and before comparing
   in `Authenticate`. (Even better: add a per-user salt.)

2. **No limits on attempts.** Add a check in `StudentDashboardForm` such as
   `SELECT COUNT(*) FROM tblResults WHERE UserID = ?` to allow one attempt only.

3. **All SQL is parameterised** (`?` + `OleDbParameter`), so users cannot inject SQL through
   the login box or the question editor.

4. **Randomisation / anti-cheat ideas:** shuffle questions (`OrderBy(Function(q) Rnd())`),
   randomise option order, record a log of focus-loss events, or require a full-screen
   maximised window while the timer runs.

---

## 10. Ideas for extending the project

* Several exams/subjects (add an `tblExams` table and a `ExamID` column).
* Question types (true/false, multiple answers, short text).
* Review mode: teachers open a result and see the answer sheet from `tblAnswers`.
* Print/PDF result slips.
* Charts on the dashboard (`System.Windows.Forms.DataVisualization.Charting`).
* Replace Access with SQLite or SQL Server — only `DatabaseHelper.vb` changes; the forms
  and the marking engine stay the same.

---

## 11. Likely viva / defence questions

**Q: Why is the marking called “automatic”?**
A: Because `MarkingEngine.Mark` compares the stored `CorrectOption` with the option the
student picked, adds the marks and computes the grade without any teacher input. It runs
the moment the student submits or the timer expires.

**Q: Where is the timer? How does it count down?**
A: A `System.Windows.Forms.Timer` with `Interval = 1000` on `ExamForm`; every tick
subtracts one second from `_remainingSeconds`, redraws `mm:ss`, and at zero calls
`FinishExam()`, which marks and saves the paper.

**Q: How do you prevent the same radio-button answer leaking into the next question?**
A: Answers are stored per `QuestionID` in a `Dictionary(Of Integer, String)`. When a
question is displayed, the radio buttons are refreshed from that dictionary; a guard flag
(`_isLoadingQuestion`) stops the programmatic “tick” from being recorded as a user answer.

**Q: Why is `SaveResult` wrapped in a transaction?**
A: A result row plus all its answer rows must be saved together. If one insert fails, the
transaction rolls back, so the database never contains a result without its answer sheet.

**Q: Why parameterised queries?**
A: To stop SQL injection and to let OLEDB handle quotes and dates correctly.

**Q: How is the database created automatically?**
A: `DatabaseHelper.CreateNewDatabase` uses ADOX (`ADOX.Catalog.Create`) to create an empty
Access file, then runs `CREATE TABLE` / `CREATE INDEX` statements and inserts the demo data.

---

*Built for Visual Studio (Windows Forms, VB.NET, .NET Framework 4.8). All code compiles
with Option Strict On; the marking engine was unit-tested before delivery.*
