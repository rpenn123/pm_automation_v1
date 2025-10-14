# User Guide: How to Use the Project Tracking Spreadsheet

Welcome to the Project Tracking Spreadsheet! This guide is here to help you understand how this tool works, what each tab does, and how your data is automatically managed. Our goal is to make your project tracking seamless and efficient.

## Table of Contents
- [The Big Picture: How It All Works](#the-big-picture-how-it-all-works)
- [The Sheets: A Tab-by-Tab Breakdown](#the-sheets-a-tab-by-tab-breakdown)
  - [1. Forecasting Sheet](#1-forecasting-sheet)
  - [2. Upcoming Sheet](#2-upcoming-sheet)
  - [3. Framing Sheet](#3-framing-sheet)
  - [4. Inventory_Elevators Sheet](#4-inventory_elevators-sheet)
  - [5. Dashboard](#5-dashboard)
- [Key Automations: How Data Moves Automatically](#key-automations-how-data-moves-automatically)
  - [Syncing Project Progress](#syncing-project-progress)
- [Frequently Asked Questions (FAQ)](#frequently-asked-questions-faq)

## The Big Picture: How It All Works

This spreadsheet is designed to be your central hub for project management. You start by entering all new projects into the **Forecasting** sheet. As you update certain fields (like a project's status), the spreadsheet automatically moves or copies that project's information to other sheets. This keeps everything organized and ensures that teams are looking at the right information at the right time.

Think of it as a smart to-do list that organizes itself.

## The Sheets: A Tab-by-Tab Breakdown

Here’s a look at what each tab is for and how to use it.

### 1. Forecasting Sheet

**What it's for:** This is where it all begins. The **Forecasting** sheet is your main workspace for entering and managing all active and potential projects. It provides a complete overview of everything in the pipeline.

**How to use it:**
-   **Enter New Projects Here:** When you have a new project, add a new row and fill in the details.
-   **Key Columns to Update:**
    -   **`PROGRESS` (Column G):** Update this to track the project's status. When you change this to **"In Progress"**, the project will automatically be copied to the **Framing** sheet for the construction team.
    -   **`PERMITS` (Column H):** When you update this field to **"approved"**, the project is automatically copied to the **Upcoming** sheet.
    -   **`DELIVERED` (Column L):**  Marking this with a "TRUE" or checking a box (if it's a checkbox) signals that the equipment has been delivered, and the project will be copied to the **Inventory_Elevators** sheet.

### 2. Upcoming Sheet

**What it's for:** This sheet gives you a focused view of projects that have received permit approval and are on the near-term horizon. It’s a cleaner, less-cluttered list of what’s coming up next.

**How it works:**
-   **Automatic Data Transfer:** You don’t need to enter any data here manually. When you set the **`PERMITS`** column in the **Forecasting** sheet to "approved", the key details for that project are automatically copied into this sheet.
-   **Progress Syncing:** The **`PROGRESS`** column in this sheet is automatically kept in sync with the **Forecasting** sheet. If you update the progress in one place, it updates in the other, ensuring everyone is on the same page.

### 3. Framing Sheet

**What it's for:** This sheet is for the construction and framing teams. It lists all projects that are officially "In Progress" and require framing work.

**How it works:**
-   **Automatic Data Transfer:** When the **`PROGRESS`** in the **Forecasting** sheet is set to **"In Progress"**, the project details are automatically copied here. This provides the framing team with a clear, actionable list of their current projects.

### 4. Inventory_Elevators Sheet

**What it's for:** This sheet tracks all projects where the necessary equipment has been delivered to the site. It’s a log of what has been delivered and for which project.

**How it works:**
-   **Automatic Data Transfer:** When you mark a project as **`DELIVERED`** in the **Forecasting** sheet (e.g., by checking a box or entering "TRUE"), that project’s information is automatically added to this sheet.

### 5. Dashboard

**What it's for:** The **Dashboard** provides a high-level, visual summary of all project data. It includes charts and key metrics to give you a quick overview of project timelines and statuses.

**How it works:**
-   **Fully Automated:** This sheet is read-only and updates automatically. There's no need to change anything here; just use it to get a quick pulse on how projects are progressing.

## Key Automations: How Data Moves Automatically

This spreadsheet contains several automations designed to save you time and keep data consistent. Here’s a quick summary:

-   **Project Transfers:** As described above, projects are automatically copied to the **Upcoming**, **Framing**, and **Inventory_Elevators** sheets based on status changes you make in the **Forecasting** sheet.

### Syncing Project Progress

To ensure everyone has the latest information, the **`PROGRESS`** field is synchronized between the **Forecasting** and **Upcoming** sheets.
-   If you update a project's progress in **Forecasting**, it will automatically update in **Upcoming**.
-   If you update it in **Upcoming**, it will also sync back to **Forecasting**.

This two-way sync means you can work in either sheet, and the status will always be consistent.

## Frequently Asked Questions (FAQ)

**Q: What happens if I accidentally delete a row in the `Upcoming`, `Framing`, or `Inventory_Elevators` sheets?**
**A:** These sheets are populated automatically. While you can delete a row, the system doesn't automatically restore it. The best practice is to manage project status from the `Forecasting` sheet. If a project is no longer relevant, consider marking it as "Cancelled" in the `Forecasting` sheet instead of deleting rows from other sheets.

**Q: Can I add my own columns?**
**A:** It is not recommended to add your own columns to the automated sheets, as this could interfere with the scripts. If you need to track additional information, it's best to do so in a separate, personal sheet or consult with the development team to see if it can be added to the main template.

**Q: How quickly do the automations run?**
**A:** The automations trigger almost instantly after you make an edit. You should see the changes in the other sheets within a few seconds.