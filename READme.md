# NGO-Campaign-Scheduler

**A VBA-based application to manage high-volume scheduling for animal welfare campaigns, featuring automated slot generation and dynamic communication.**

> **Status:** MVP
> 
> **Role:** volunteer

<div align="center">
  <!--<p align="left"><b>Tech Stack</b></p>-->
  <img src="https://img.shields.io/badge/Excel-217346?style=for-the-badge"  />
  <img src="https://img.shields.io/badge/VBA-gray?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Power_Query-F2C811?style=for-the-badge" />
  <img src="https://img.shields.io/badge/%F0%9F%99%8C%20Volunteering-4CAF50?style=for-the-badge" />
</div>

## Background & Motivation

As a volunteer contributor, I stepped into a scheduling system that was scattered across multiple disconnected sheets with no standardized data structure. Token formats varied from month to month and were sometimes duplicated, which led to overbooking incidents where more animals arrived than the medical team could handle. Tracking empty slots was difficult, and reporting required tedious half-manual counting.  

To solve this, I consolidated all the disparate data into one structured Table, enforcing a unified schema that eliminated inconsistencies.  I then introduced Pivot Tables to summarize capacity by day and hour, giving the manager instant visibility into scheduling conflicts with a simple "Refresh All." 

Finally, I built a VBA application to automate the last mile of logistics—managing individual tokens and communicating with tutors—ensuring that the digital schedule aligned perfectly with physical reality.



**⚡Impact⚡** 

> 

## 🎯 Challenges addressed

* **Limited Capacity:** Specific slots for different animal genders (CF, CM, FF, FM).
* **Communication Friction:** Manually typing confirmation messages to hundreds of tutors is slow and error-prone.
* **Change Management:** Handling cancellations and transfers manually often leads to "Ghost Slots" (unused capacity) or overbooking.
  
  

## 🛠️ The Solution

I engineered a **Relational Token Management System** that bridges the gap between a user-friendly frontend (counts) and a granular database (individual slots).

### 1. Delta-Based Synchronization Engine

* **The Logic:** Users input aggregate demand (e.g., "3 Dogs") on the frontend. The system calculates the delta between the Requested Count and the Database State.
  
  * **Expansion:** If demand > existing tokens, it generates specific new IDs (cm_vsAddRowTo_tbDBTokens) with status "Agendado".
  
  * **Contraction:** If demand < existing tokens, it intelligently cancels the excess specific tokens (cm_vfCancelTokenOn_tbDBTokens), preserving data history rather than deleting rows.

* **Trigger:** Executed via BeforeDoubleClick events, ensuring immediate consistency without manual "Save" buttons.
  
  

### 2. The Transfer Transaction Manager

I built a **UserForm Interface** (fmTransferTokens) to handle ownership changes safely.

* **Relational Integrity:** When tokens are transferred from Person A to Person B, the system:
  
  1. Marks the original token status as **"Transferido"**.
  
  2. Creates a new trace record in tbDBTransfer linking **Old_Schedule_ID** → **New_Schedule_ID**.
  
  3. Updates the Token's Foreign Key (FK_IDAgendamento) to the new owner.

* **Result:** A complete audit trail. We know exactly which slot moved where, preventing "double slots" (people with 1 Dog Book who transfer to someone with 3 Dogs)
  
  

### 3. Dynamic Communication Generator

* **Templating:** Advanced Excel formulas (LET/LAMBDA) dynamically construct WhatsApp messages by parsing tags like `<nome>` and `<senhas>`.

* **Clipboard Automation:** A BeforeRightClick trigger executes vfCopyToClipboard, utilizing the MSForms.DataObject library to bypass the need for manual selection and copying.
  
  

### 4. Robust Architecture (OOP in VBA)

To ensure maintainability, the system uses **Object-Oriented** principles:

* **Grid Abstraction (clRange):** All worksheet references (Columns, Tables) are mapped in a Class Module. If the Excel layout changes, only the Class is updated—the logic remains untouched.

* **Status Standardization (clString):** A dedicated class manages string constants ("Agendado", "Cancelado").
  
  

## 📂 Repository Structure

```text
/NGO-Campaign-Scheduler
│
├── /.Source Code
|   ├── /Buttons
|   |   └── Buttons.bas    # Buttons Event listeners
|   |
│   ├── /Classes
│   │   ├── clRange.cls        # Grid Abstraction Layer
│   │   ├── clString.cls       # Global String Constants
|   |   └── vbASchedule.cls    # Worksheet Event listener
│   │
│   ├── /Forms
│   │   ├── fmTransferTokens.frm    # Slot Transfer Interface
│   │   └── fmTransferTokens.frx    #
│   │
│   ├── /Worksheet function & Controllers
|   |   ├── cm_Buttons.bas             # 
|   |   ├── cm_fmTransferTokens.bas    #
|   |   ├── cm_vbASchedule.bas         #
|   |   ├── cmFunctions.bas            # Universal Functions
|   |   ├── wf_fmTransferTokens.bas    #    
|   |   ├── wf_vbASchedule.bas         #
|   └── Excel Formulas.md    #
│
├── READme
└── Agendamentos.xlsm.zip    # App
```

---

## 🚀 Future Roadmap

* **WhatsApp Automation:** Integration with an API to send the generated messages automatically.
* **Cloud Sync:** Porting the Backend to SharePoint/SQL for multi-user simultaneous editing and deeper integration
