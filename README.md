
# MetaNull.WeeklyCalendarDocx

This repository contains personal scripts, including the source for the [MetaNull.WeeklyCalendarDocx](https://www.powershellgallery.com/packages/MetaNull.WeeklyCalendarDocx) PowerShell module, published on the PowerShell Gallery.

## 📦 Installation

To install the module from the PowerShell Gallery:

```powershell
Install-Module -Name MetaNull.WeeklyCalendarDocx -Scope CurrentUser
```

## 🚀 Usage

After installing, you can import and use the module:

```powershell
Import-Module MetaNull.WeeklyCalendarDocx
```

To generate a weekly calendar document:

```powershell
New-WeeklyCalendar -Year 2025 -FromWeek 10 -NumberOfWeeks 8
```

For more options and help:

```powershell
Get-Help New-WeeklyCalendar -Detailed
```

## 📄 Documentation

- The module provides a command to generate a formatted weekly calendar in Microsoft Word.
- Supports ISO week numbering, multiple languages, and custom formatting.
- See the [PowerShell Gallery page](https://www.powershellgallery.com/packages/MetaNull.WeeklyCalendarDocx) for the latest documentation and examples.

## 🛠️ Development

- The source code for the module is maintained in this repository.

