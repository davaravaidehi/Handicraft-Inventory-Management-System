# Handicraft-Inventory-Management-System

A Windows desktop application built with **C# Windows Forms** and a **Microsoft Access** database to manage the day-to-day work of a handicraft business: items, customers, workers, purchases, sales, returns, stock and reports.

## Table of Contents

- [Features](#features)
- [Modules](#modules)
- [Tech Stack](#tech-stack)
- [Database](#database)
- [Project Structure](#project-structure)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Future Improvements](#future-improvements)
- [Author](#author)

## Features

- Secure login and change-password option
- Splash and welcome screens with an MDI (multi-window) main menu
- Master records for items, customers and workers
- Purchase entry and purchase return entry
- Sales entry and sales return entry
- Live stock details that update with purchases, sales and returns
- Reports for items, customers, workers, purchases, sales, returns and stock
- Built-in utilities: Notepad, Calculator, WordPad and Chrome shortcuts
- Profile, log out and About Us screens

## Modules

| Menu | Option | Purpose |
| --- | --- | --- |
| **Master** | Item Master | Add, update and delete handicraft items |
| | Customer Master | Manage customer details |
| | Worker Master | Manage worker details |
| **Transaction** | Purchase Entry | Record items purchased |
| | Sales Entry | Record items sold to customers |
| | Purchase Return Entry | Record items returned to suppliers |
| | Sales Return Entry | Record items returned by customers |
| **Stock** | Stock Details | View and manage current stock |
| **Report** | Item, Customer, Worker, Purchase, Purchase Return, Sales, Sales Return, Stock | View and search records |
| **Utility** | Notepad, Calculator, WordPad, Chrome | Quick access to common tools |
| **Profile** | Change Password, Log Out | Account options |

## Tech Stack

| Technology | Purpose |
| --- | --- |
| C# | Programming language |
| Windows Forms (.NET Framework 4.5) | Desktop user interface |
| Microsoft Access (`.accdb`) | Database |
| OLEDB (`Microsoft.ACE.OLEDB.12.0`) | Database connection |
| Visual Studio | Development environment |
| Git / GitHub | Version control |

## Database

The application uses the Access database file `login.accdb` in the root folder. Main tables used by the application:

- `login`
- `item_master`
- `customer_master`
- `worker_master`
- `purchase_master`
- `purchase_return_master`
- `sales_master`
- `sales_return_master`
- `stock_master`

## Project Structure

```
Handicraft/
├── Handicraft.sln          # Visual Studio solution
├── login.accdb             # Access database
├── Handicraft/             # Main project
│   ├── Program.cs
│   ├── splash.cs, welcome.cs, login.cs, MDI.cs
│   ├── Item_master.cs, Customer_master.cs, Worker_master.cs
│   ├── Purchase_master.cs, Purchase_return_master.cs
│   ├── Sales_master.cs, Sales_return_master.cs
│   ├── Stock_master.cs
│   ├── *_report.cs         # Report screens
│   └── Resources/          # Images and icons
├── Abc/, images/, photo/   # Item and background images
└── README.md
```

## Getting Started

### Requirements

- Windows 10 or 11
- Visual Studio 2015 or later (with .NET desktop development)
- .NET Framework 4.5
- [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (ACE OLEDB 12.0). Install the 32-bit or 64-bit version to match your Office/Visual Studio setup.

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/davaravaidehi/Handicraft-Inventory-Management-System.git
   ```
2. Open `Handicraft.sln` in Visual Studio.
3. **Update the database path.** The forms connect with a fixed path:
   ```csharp
   Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb
   ```
   Replace it with the location of `login.accdb` on your computer in each form's `.cs` file (use Find and Replace in Files: **Ctrl + Shift + H**).
4. Build the solution (**Ctrl + Shift + B**) and press **F5** to run.

## Usage

1. Start the application and log in with your username and password.
2. Add items, customers and workers from the **Master** menu.
3. Record purchases and sales from the **Transaction** menu.
4. Check available stock under **Stock**.
5. Open the **Report** menu to view records.

**Vaidehi Davara** · [GitHub](https://github.com/davaravaidehi)
