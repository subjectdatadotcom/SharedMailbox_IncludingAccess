# 📬 Create Shared Mailboxes Script

This PowerShell script automates the creation of **Shared Mailboxes** in **Exchange Online** based on a CSV input file.

It supports:
- New shared mailbox creation.
- Assigning FullAccess and SendAs permissions based on user mappings.
- Tracking unmapped users and existing mailboxes.

---

## 📄 Script Overview

- Checks if the **ExchangeOnlineManagement** module is installed and imports it.
- Connects to Exchange Online.
- Loads shared mailbox data from `SharedMailboxes.csv`.
- Checks if a mailbox already exists (optional, commented out).
- Creates new shared mailboxes:
  - **Display Name** is automatically prefixed with **"SPOT-"**.
  - **Primary SMTP Address** is built under the **`target.com`** domain.
- Grants:
  - **FullAccess** permission based on `PrimaryOwnerSMTP`, `OwnersSMTP`, and `MbxAccess` fields.
  - **SendAs** permission based on `SendAccess` field.
- Tracks:
  - Unmapped users that were not found in the mapping file.
  - Already existing mailboxes (optional tracking).

---

## 📂 Input Files

| File | Description |
|:-----|:------------|
| `SharedMailboxes.csv` | List of mailboxes to be created, with alias, display name, owners, access users. |
| `UserMapping.csv` | Mapping file between source (e.g., source) and target (e.g., target) user accounts. |

---

### 🧩 Expected `SharedMailboxes.csv` format example

| alias | displayname | PrimaryOwnerSMTP | OwnersSMTP | MbxAccess | SendAccess |
|:------|:------------|:-----------------|:-----------|:----------|:-----------|
| HRTeam | HR Team Mailbox | jane.doe@source.com | john.doe@source.com | jane.doe@source.com;jack.doe@source.com | john.doe@source.com |

---

### 🧩 Expected `UserMapping.csv` format example

| source Email/UPN | target UPN |
|:-----------------|:------------|
| jane.doe@source.com | jane.doe@target.com |
| john.doe@source.com | john.doe@target.com |
| jack.doe@source.com | jack.doe@target.com |

---

## 📜 Output Files

| File | Description |
|:-----|:------------|
| `ExistingSharedMailboxes.csv` | Mailboxes that already exist in Exchange Online (skipped creation) |
| `UnmappedUsersReport.csv` | Users found in SharedMailboxes.csv who could not be mapped using UserMapping.csv |

---

## ⚙️ Prerequisites

- PowerShell 5.1 or later.
- ExchangeOnlineManagement PowerShell module installed.
- Administrative rights to create shared mailboxes in Exchange Online.

---

## 🚀 How to Run

1. Open PowerShell as Administrator.
2. Place `SharedMailboxes.csv` and `UserMapping.csv` in the **same directory** as the script.
3. Run the script:

```powershell
.\SharedMailbox.ps1
```

4. Follow any authentication prompts to log into Exchange Online.

---

## 📋 Important Notes

- The primary SMTP address is generated using the **alias** field and **target.com** domain.
- **Display Name** is automatically prefixed with **SPOT-**.
- Existing mailbox check is **optional** (currently commented out in the script).
- Archive and auto-expanding archive settings are supported if specified in the CSV.
- All generated reports (existing mailboxes, unmapped users) are saved in the same folder.

---

## 👨‍💻 Author

- SubjectData
- For internal mailbox migration projects.
