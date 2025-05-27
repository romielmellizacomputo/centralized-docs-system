# Centralized Docs System (CDS) — General Guide

## Introduction  
The **Centralized Docs System (CDS)** was launched on **October 3, 2024** to solve key challenges faced by QA teams in managing documentation and tasks. Before CDS, teams created new documents every sprint, causing fragmented data that made it difficult to track historical issues, progress, and insights.  

Manual data entries—such as input kiosks, updating bug sheets post-testing, and recording issues from Forge to Google Sheets—were time-consuming and error-prone. CDS was developed to automate these processes, streamline workflows, and improve data accuracy.

By integrating tools like **Google Apps Script**, **GitHub**, **GitLab CI/CD**, and **Google Cloud API**, CDS automates data fetching, updating, and recording. This empowers QA teams to focus on critical activities such as bug reporting, usability discussions, merge request reviews, investigations, and environment testing.

*This system and guide were proposed by Romiel Melliza Computo.*

---

## Core Components of CDS

### 1. Leads/Management CDS  
**Purpose:**  
Provides managers and supervisors with a centralized hub to monitor outputs, timelines, and assignments across all teams.

**Key Features:**  
- Centralized monitoring of team outputs, sprints, and KPIs  
- Task assignment with automatic syncing to Team CDS  
- Access to test-related documents (Test Cases, Test Scripts, etc.)  
- Sprint and milestone timeline management  
- Real-time automated data updates every 5–10 minutes  

**Getting Started:**  
- Access via shared link provided by QA Manager  
- Use dashboard to select teams and sprints for streamlined oversight  

---

### 2. Team CDS  
**Purpose:**  
Consolidates all QA tasks (manual and automated) into a single source of truth, eliminating fragmented data and boosting accountability.

**Key Features:**  
- Centralized QA task documentation  
- Automated data fetching from GitLab/Forge to reduce manual work  
- Clear task ownership and progress tracking  
- Visual performance and accountability metrics  

**Getting Started:**  
- Explore sheets like Help, Dashboard, and Test Cases  
- Understand your roles within the Team CDS framework  

---

### 3. Automated Sprint Report  
**Purpose:**  
Provides a dynamic, weekly overview of team outputs and sprint progress to offer insights into team performance.

**Key Features:**  
- Automatically generated weekly deliverable summaries  
- Sprint leave report for resource tracking  
- Visual progress bars and charts to track performance  
- Customizable dropdowns for team-specific reports  

---

### 4. Automated Portals  
**Purpose:**  
Acts as a centralized Google Sheet for all QA test data, ensuring real-time updates and easy access.

**Key Features:**  
- Centralized management of all QA test data  
- Auto-refresh every 5–10 minutes to keep data current  
- Data extraction by team or feature for targeted analysis  

**Getting Started:**  
- Portal is view-only; do **not** edit directly  
- Use the linked Test Case document for testing activities to keep data accurate  

---

### 5. Automated Test Case Sheets  
**Versions:**  
- V1.1.0: GitLab CI/CD-Driven Test Case Sheets - TEMPLATE (make your own copy)  
- V2.0.0: GitHub-Driven Test Case Sheets - TEMPLATE (make your own copy)  

**Purpose:**  
Eliminates repetitive manual formatting and entry in QA documentation, enabling focus on effective test case creation.

**Key Features:**  
- Automated step numbering and formatting  
- Real-time syncing with Automated Test Case & Scenario Portal  
- Smart templates prevent ID conflicts and ensure consistency  

**Getting Started:**  
- Make a copy of the appropriate template  
- Place it in the correct repository  
- Follow setup instructions to link the sheet to the automation system  

---

## Conclusion  
The **Centralized Docs System (CDS)** improves QA team efficiency by automating and centralizing documentation and task management. Each component helps teams collaborate better, reduce manual errors, and deliver higher-quality results.  

For detailed guides, templates, and resources, please explore the full **CDS repository**.

---

*Developed and maintained by Romiel Melliza Computo.*
